'use strict';

const express = require('express');
const crypto = require('crypto');
const fs = require('fs-extra');
const path = require('path');
const ExcelJS = require('exceljs');
const Database = require('better-sqlite3');

const app = express();

const PORT = Number(process.env.PORT || 3000);

// ============================================================================
// CONFIGURAZIONE
// ============================================================================

// IMPORTANTE:
// Il NAS deve essere montato nel container Docker.
//
const NAS_BASE_PATH = process.env.NAS_BASE_PATH;
const DB_PATH = process.env.DB_PATH;

if (!DB_PATH) {
    throw new Error(
        'ERRORE CONFIGURAZIONE: DB_PATH non è impostato. ' +
        'Il server non verrà avviato.'
    );
}

if (!NAS_BASE_PATH) {
    throw new Error(
        'ERRORE CONFIGURAZIONE: NAS_BASE_PATH non è impostato. ' +
        'Il server non verrà avviato.'
    );
}

if (!path.isAbsolute(DB_PATH)) {
    throw new Error(
        `ERRORE CONFIGURAZIONE: DB_PATH deve essere un percorso assoluto: ${DB_PATH}`
    );
}

if (!path.isAbsolute(NAS_BASE_PATH)) {
    throw new Error(
        `ERRORE CONFIGURAZIONE: NAS_BASE_PATH deve essere un percorso assoluto: ${NAS_BASE_PATH}`
    );
}
// Timezone.
// In Docker impostare:
// TZ=Europe/Rome
const TIMEZONE = process.env.TZ || 'Europe/Rome';

const SOGLIA_8_ORE = 8 * 3600;

// FIX 3: intervallo minimo (in secondi) tra due timbrature dello stesso
// dipendente nella stessa giornata. Evita che una tessera tenuta vicino al
// lettore generi più passaggi ravvicinati (chiudendo la giornata a 0 ore).
// Impostare MIN_INTERVALLO_TIMBRATURE_SEC=0 per disattivare il controllo.
const MIN_INTERVALLO_TIMBRATURE_SEC = (() => {
    const grezzo = process.env.MIN_INTERVALLO_TIMBRATURE_SEC;

    if (grezzo === undefined || grezzo === '') {
        return 60;
    }

    const valore = Number(grezzo);

    return Number.isFinite(valore) && valore >= 0
        ? valore
        : 60;
})();

// FIX 4: chiave API opzionale. Se impostata, tutte le rotte /api/* richiedono
// l'header "x-api-key" con lo stesso valore.
const API_KEY = process.env.API_KEY || null;

const MESI = [
    'Gennaio',
    'Febbraio',
    'Marzo',
    'Aprile',
    'Maggio',
    'Giugno',
    'Luglio',
    'Agosto',
    'Settembre',
    'Ottobre',
    'Novembre',
    'Dicembre'
];

// ============================================================================
// EXPRESS
// ============================================================================

app.disable('x-powered-by');

app.use(express.json({
    limit: '100kb'
}));

// ============================================================================
// AUTENTICAZIONE API (FIX 4)
// ============================================================================
//
// Senza autenticazione chiunque sulla rete può inviare timbrature false o
// rinominare i dipendenti. Il controllo è attivo solo se API_KEY è impostata,
// così il firmware ESP32 esistente continua a funzionare finché non viene
// aggiornato per inviare l'header x-api-key.
//
// /health resta pubblico (serve all'healthcheck di Docker).
//
// ============================================================================

function richiediApiKey(req, res, next) {
    if (!API_KEY) {
        return next();
    }

    const fornita = req.get('x-api-key');

    if (typeof fornita === 'string' && fornita.length > 0) {
        const hashFornita = crypto
            .createHash('sha256')
            .update(fornita)
            .digest();

        const hashAttesa = crypto
            .createHash('sha256')
            .update(API_KEY)
            .digest();

        if (crypto.timingSafeEqual(hashFornita, hashAttesa)) {
            return next();
        }
    }

    return res.status(401).json({
        status: 'error',
        message: 'Non autorizzato'
    });
}

app.use('/api', richiediApiKey);

// ============================================================================
// DATABASE SQLITE
// ============================================================================

if (!fs.pathExistsSync(DB_PATH)) {
    throw new Error(
        `ERRORE DATABASE: file non trovato: ${DB_PATH}. ` +
        'Verificare il bind mount di dipendenti.db.'
    );
}

if (!fs.pathExistsSync(NAS_BASE_PATH)) {
    throw new Error(
        `ERRORE NAS: percorso non raggiungibile: ${NAS_BASE_PATH}. ` +
        'Verificare che il NAS sia montato sull\'host e che il bind mount Docker sia attivo.'
    );
}

const db = new Database(DB_PATH);

db.pragma('busy_timeout = 5000');
db.pragma('foreign_keys = ON');

const schemaDipendenti = db.prepare(`
    SELECT name
    FROM sqlite_master
    WHERE type = 'table'
      AND name = 'dipendenti'
`).get();

if (!schemaDipendenti) {
    throw new Error(
        `ERRORE DATABASE: la tabella "dipendenti" non esiste nel DB ${DB_PATH}.`
    );
}

const colonneDipendenti = db.prepare(`
    PRAGMA table_info(dipendenti)
`).all();

const nomiColonne = new Set(
    colonneDipendenti.map(colonna => colonna.name)
);

if (!nomiColonne.has('uid') || !nomiColonne.has('nome')) {
    throw new Error(
        `ERRORE DATABASE: la tabella "dipendenti" deve contenere le colonne "uid" e "nome". ` +
        `Colonne trovate: ${Array.from(nomiColonne).join(', ')}`
    );
}

// FIX 13: l'UPSERT "ON CONFLICT(uid)" funziona solo se uid è PRIMARY KEY
// oppure ha un vincolo/indice UNIQUE. In caso contrario ogni registrazione
// fallirebbe a runtime: meglio accorgersene all'avvio.
const colonnePk = colonneDipendenti.filter(colonna => colonna.pk > 0);

const uidHaVincoloUnico =
    (
        colonnePk.length === 1 &&
        colonnePk[0].name === 'uid'
    ) ||
    db.prepare('PRAGMA index_list(dipendenti)').all().some(indice => {
        if (!indice.unique) {
            return false;
        }

        const colonneIndice = db
            .prepare(`PRAGMA index_info("${indice.name.replace(/"/g, '""')}")`)
            .all();

        return colonneIndice.length === 1 &&
            colonneIndice[0].name === 'uid';
    });

if (!uidHaVincoloUnico) {
    throw new Error(
        'ERRORE DATABASE: la colonna "uid" della tabella "dipendenti" deve essere ' +
        'PRIMARY KEY oppure avere un vincolo UNIQUE (necessario per l\'UPSERT).'
    );
}

const stmtGetDipendente = db.prepare(`
    SELECT nome
    FROM dipendenti
    WHERE uid = ?
`);

const stmtGetDipendenti = db.prepare(`
    SELECT uid, nome
    FROM dipendenti
    ORDER BY nome COLLATE NOCASE
`);

const stmtUpsertDipendente = db.prepare(`
    INSERT INTO dipendenti (uid, nome)
    VALUES (?, ?)
    ON CONFLICT(uid)
    DO UPDATE SET nome = excluded.nome
`);

function ottieniNomeDipendente(uid) {
    const dipendente = stmtGetDipendente.get(uid);

    if (dipendente) {
        return dipendente.nome;
    }

    console.warn(
        `⚠️ UID non presente nel database: [${uid}]`
    );

    return `Sconosciuto (${uid})`;
}

// ============================================================================
// LOCK SCRITTURA EXCEL
// ============================================================================
//
// Evita che due richieste NFC contemporanee leggano e scrivano
// contemporaneamente lo stesso file Excel.
//
// ATTENZIONE:
// Questo lock funziona all'interno dello stesso processo Node.js.
// Se utilizzi più container Node.js contemporaneamente sullo stesso Excel,
// servirà un sistema di lock distribuito.
//
// ============================================================================

let excelQueue = Promise.resolve();

function conLockExcel(fn) {
    const operazione = excelQueue.then(() => fn());

    // Manteniamo la coda attiva anche se un'operazione fallisce.
    excelQueue = operazione.catch(() => {});

    return operazione;
}

// ============================================================================
// DATA E ORA
// ============================================================================
//
// Restituisce:
// - anno
// - mese
// - giorno
// - dataKey -> formato stabile YYYY-MM-DD
// - data -> formato visualizzato DD/MM/YYYY
// - ora -> HH:mm:ss
//
// dataKey è fondamentale per confrontare correttamente i giorni.
// NON utilizziamo direttamente il testo visualizzato nell'Excel.
// ============================================================================

function getDataOraItalia() {
    const now = new Date();

    // FIX 5: hourCycle 'h23' invece di hour12:false. Con hour12:false alcune
    // versioni di ICU restituiscono "24" a mezzanotte (24:00:05 invece di 00:00:05).
    const parts = new Intl.DateTimeFormat('it-IT', {
        timeZone: TIMEZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hourCycle: 'h23'
    }).formatToParts(now);

    const valori = {};

    for (const part of parts) {
        valori[part.type] = part.value;
    }

    const anno = valori.year;
    const mese = Number(valori.month);
    const giorno = valori.day;

    return {
        anno,
        mese,
        giorno,

        // Identificatore tecnico della giornata.
        // Esempio: 2026-09-24
        dataKey: `${anno}-${String(mese).padStart(2, '0')}-${giorno}`,

        // Formato visualizzato nell'Excel.
        // Esempio: 24/09/2026
        data: `${giorno}/${String(mese).padStart(2, '0')}/${anno}`,

        // Ora della timbratura.
        ora: `${valori.hour}:${valori.minute}:${valori.second}`
    };
}

// ============================================================================
// CONVERSIONE DATA EXCEL
// ============================================================================
//
// ExcelJS può restituire:
// - stringa
// - Date
// - altri valori.
//
// Questa funzione converte tutto in:
// YYYY-MM-DD
//
// Esempi:
//
// 24/09/2026 -> 2026-09-24
// 2026-09-24 -> 2026-09-24
// Date(...)  -> 2026-09-24
//
// ============================================================================

function normalizzaDataExcel(value) {
    if (!value) {
        return null;
    }

    // ------------------------------------------------------------
    // Caso 1: ExcelJS restituisce una vera Date
    // ------------------------------------------------------------

    if (value instanceof Date) {
        if (!Number.isFinite(value.getTime())) {
            return null;
        }

        // FIX 6: ExcelJS converte il seriale Excel in una Date UTC "ingenua"
        // (mezzanotte UTC del giorno scritto nella cella). Applicare il fuso
        // orario del server potrebbe spostare la data di un giorno: si
        // usano quindi direttamente i componenti UTC.
        return value.toISOString().slice(0, 10);
    }

    // ------------------------------------------------------------
    // Caso 2: valore testuale
    // ------------------------------------------------------------

    const valore = String(value).trim();

    if (!valore) {
        return null;
    }

    // ------------------------------------------------------------
    // Formato ISO:
    //
    // 2026-09-24
    // ------------------------------------------------------------

    if (/^\d{4}-\d{2}-\d{2}$/.test(valore)) {
        return valore;
    }

    // ------------------------------------------------------------
    // Formato italiano:
    //
    // 24/09/2026
    // ------------------------------------------------------------

    const match = valore.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
    );

    if (match) {
        const giorno = match[1].padStart(2, '0');
        const mese = match[2].padStart(2, '0');
        const anno = match[3];

        return `${anno}-${mese}-${giorno}`;
    }

    return null;
}

// ============================================================================
// UTILITÀ ORE
// ============================================================================

function timeToSeconds(timeStr) {
    if (timeStr === null || timeStr === undefined || timeStr === '') {
        return 0;
    }

    // ExcelJS restituisce spesso le celle formattate come ora come Date,
    // non come stringhe. I componenti UTC rappresentano l'ora memorizzata
    // nel seriale Excel senza applicare il fuso orario del server.
    if (timeStr instanceof Date) {
        if (!Number.isFinite(timeStr.getTime())) {
            return 0;
        }

        return timeStr.getUTCHours() * 3600 +
            timeStr.getUTCMinutes() * 60 +
            timeStr.getUTCSeconds();
    }

    // Supporta anche un valore numerico Excel che rappresenta una frazione
    // di giornata (es. 0,5 = 12:00).
    if (typeof timeStr === 'number') {
        if (!Number.isFinite(timeStr) || timeStr < 0 || timeStr >= 1) {
            return 0;
        }

        return Math.round(timeStr * 24 * 3600);
    }

    if (typeof timeStr !== 'string') {
        return 0;
    }

    const valore = timeStr.trim();
    if (!valore) {
        return 0;
    }

    // FIX 7: si usa il valore già ripulito (prima veniva spezzato timeStr).
    const parts = valore.split(':');

    if (parts.length < 2 || parts.length > 3) {
        return 0;
    }

    const h = Number(parts[0]);
    const m = Number(parts[1]);
    const s = Number(parts[2] || 0);

    if (
        !Number.isFinite(h) ||
        !Number.isFinite(m) ||
        !Number.isFinite(s) ||
        h < 0 ||
        m < 0 || m > 59 ||
        s < 0 || s > 59
    ) {
        return 0;
    }

    return h * 3600 + m * 60 + s;
}

function secondsToTime(totalSec) {
    if (!Number.isFinite(totalSec) || totalSec <= 0) {
        return '00:00:00';
    }

    totalSec = Math.floor(totalSec);

    const h = Math.floor(totalSec / 3600);
    const m = Math.floor((totalSec % 3600) / 60);
    const s = totalSec % 60;

    return [
        String(h).padStart(2, '0'),
        String(m).padStart(2, '0'),
        String(s).padStart(2, '0')
    ].join(':');
}

// ============================================================================
// VALIDAZIONE UID
// ============================================================================

function normalizzaUid(uid) {
    if (typeof uid !== 'string') {
        return null;
    }

    const valore = uid.trim().toUpperCase();

    if (!valore) {
        return null;
    }

    if (valore.length > 100) {
        return null;
    }

    return valore;
}

// ============================================================================
// STRUTTURA EXCEL
// ============================================================================

function creaStrutturaWorksheet(worksheet) {
    worksheet.columns = [
        {
            header: 'Dipendente',
            key: 'dipendente',
            width: 25
        },
        {
            header: 'UID Tessera',
            key: 'uid',
            width: 16
        },
        {
            header: 'Data',
            key: 'data',
            width: 14
        },
        {
            header: '1° Passaggio (Entrata)',
            key: 'p1',
            width: 22
        },
        {
            header: '2° Passaggio (Uscita P.)',
            key: 'p2',
            width: 24
        },
        {
            header: '3° Passaggio (Rientro P.)',
            key: 'p3',
            width: 24
        },
        {
            header: '4° Passaggio (Uscita)',
            key: 'p4',
            width: 22
        },
        {
            header: 'Totale Ore Lavorate',
            key: 'totale',
            width: 22
        },
        {
            header: 'Ore Straordinario (>8h)',
            key: 'straordinari',
            width: 24
        }
    ];

    const headerRow = worksheet.getRow(1);

    // FIX 8: i colori in ExcelJS vanno espressi in ARGB a 8 cifre
    // (con 6 cifre Excel può interpretarli in modo errato).
    headerRow.font = {
        bold: true,
        color: {
            argb: 'FFFFFFFF'
        }
    };

    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
            argb: 'FF1F4E78'
        }
    };

    headerRow.alignment = {
        vertical: 'middle',
        horizontal: 'center'
    };
}

// ============================================================================
// SALVATAGGIO ATOMICO (FIX 1)
// ============================================================================
//
// writeFile() direttamente sul file definitivo, se interrotto (crash,
// riavvio del container, NAS che si disconnette), lascia un .xlsx
// troncato e quindi l'intero mese va perso.
// Si scrive prima su un file temporaneo nella stessa cartella e poi lo si
// rinomina: il rename è atomico e il file originale resta integro finché
// la nuova versione non è completa.
//
// ============================================================================

async function salvaWorkbook(workbook, percorsoFile) {
    const percorsoTemp = `${percorsoFile}.${process.pid}.tmp`;

    try {
        await workbook.xlsx.writeFile(percorsoTemp);
        await fs.rename(percorsoTemp, percorsoFile);
    } catch (error) {
        await fs.remove(percorsoTemp).catch(() => {});
        throw error;
    }
}

// ============================================================================
// RIGA TOTALI MENSILI
// ============================================================================

function rimuoviRigaTotali(worksheet) {
    // Un file già affetto dal problema può contenere più righe totali.
    // Rimuoverne una sola a ogni timbratura le farebbe accumulare nel tempo.
    for (
        let rowNumber = worksheet.rowCount;
        rowNumber >= 2;
        rowNumber--
    ) {
        const valore = worksheet.getRow(rowNumber).getCell(1).value;

        if (
            typeof valore === 'string' &&
            valore.trim().toUpperCase() === 'TOTALI MENSILI'
        ) {
            worksheet.spliceRows(rowNumber, 1);
        }
    }
}

// Elimina solo le righe vuote in coda (comprese le righe modello con
// 00:00:00 nelle colonne dei totali). ExcelJS le considera righe usate:
// lasciarle lì farebbe accodare le nuove giornate molto più in basso.
function eRigaSegnaposto(row) {
    const vuoto = value =>
        value === null ||
        value === undefined ||
        value === '';

    const nessunDatoPresenza = [1, 2, 3, 4, 5, 6, 7]
        .every(colonna => vuoto(row.getCell(colonna).value));

    const totaleVuotoOZero = value =>
        vuoto(value) ||
        value === 0 ||
        value === '00:00:00';

    return nessunDatoPresenza &&
        totaleVuotoOZero(row.getCell(8).value) &&
        totaleVuotoOZero(row.getCell(9).value);
}

function rimuoviSegnapostoFinali(worksheet) {
    for (let rowNumber = worksheet.rowCount; rowNumber >= 2; rowNumber--) {
        const row = worksheet.getRow(rowNumber);

        if (!eRigaSegnaposto(row)) {
            break;
        }

        worksheet.spliceRows(rowNumber, 1);
    }
}

// ============================================================================
// CALCOLO TOTALI MENSILI
// ============================================================================

function calcolaTotaliMensili(worksheet) {
    let totMeseLavoratoSec = 0;
    let totMeseStraordinariSec = 0;

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 1) {
            return;
        }

        // Non contare mai la riga TOTALI MENSILI.
        if (
            row.getCell(1).value ===
            'TOTALI MENSILI'
        ) {
            return;
        }

        totMeseLavoratoSec += timeToSeconds(
            row.getCell(8).value
        );

        totMeseStraordinariSec += timeToSeconds(
            row.getCell(9).value
        );
    });

    return {
        lavorato: totMeseLavoratoSec,
        straordinari: totMeseStraordinariSec
    };
}

// ============================================================================
// AGGIUNTA TOTALI MENSILI
// ============================================================================

function aggiungiTotaliMensili(worksheet) {
    // Prima rimuoviamo eventuali vecchi totali.
    rimuoviRigaTotali(worksheet);

    // I valori salvati in precedenza possono essere errati o avere un tipo
    // Excel diverso da stringa. Ricalcoliamo ogni giorno dai passaggi grezzi.
    ricalcolaOreGiornaliere(worksheet);

    const totali =
        calcolaTotaliMensili(worksheet);

    // Array esplicito: le chiavi definite in worksheet.columns non sono
    // conservate nel file XLSX e non sono disponibili dopo readFile().
    const rigaTotali = worksheet.addRow([
        'TOTALI MENSILI',
        '',
        '',
        '',
        '',
        '',
        '',
        secondsToTime(totali.lavorato),
        secondsToTime(totali.straordinari)
    ]);

    rigaTotali.font = {
        bold: true
    };

    rigaTotali.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
            argb: 'FFD9E1F2'
        }
    };

    rigaTotali.alignment = {
        vertical: 'middle',
        horizontal: 'center'
    };
}

// ============================================================================
// CALCOLO ORE LAVORATE
// ============================================================================
//
// 2 timbrature:
//
// P1 = Entrata
// P2 = Uscita
//
// Totale:
// P2 - P1
//
//
//
// 4 timbrature:
//
// P1 = Entrata
// P2 = Uscita pausa
// P3 = Rientro pausa
// P4 = Uscita
//
// Totale:
// (P2 - P1) + (P4 - P3)
//
// ============================================================================

function calcolaOreLavorate(
    p1,
    p2,
    p3,
    p4
) {
    const t1 = timeToSeconds(p1);
    const t2 = timeToSeconds(p2);
    const t3 = timeToSeconds(p3);
    const t4 = timeToSeconds(p4);

    let secLavorati = 0;

    // ------------------------------------------------------------
    // 2 timbrature
    // ------------------------------------------------------------

    if (
        p1 &&
        p2 &&
        !p3 &&
        !p4
    ) {
        secLavorati =
            Math.max(
                0,
                t2 - t1
            );
    }

    // ------------------------------------------------------------
    // 4 timbrature
    // ------------------------------------------------------------

    else if (
        p1 &&
        p2 &&
        p3 &&
        p4
    ) {
        const mattina =
            Math.max(
                0,
                t2 - t1
            );

        const pomeriggio =
            Math.max(
                0,
                t4 - t3
            );

        secLavorati =
            mattina +
            pomeriggio;
    }

    // ------------------------------------------------------------
    // 1 oppure 3 timbrature
    //
    // Non possiamo ancora determinare il totale.
    // ------------------------------------------------------------

    return secLavorati;
}

function ricalcolaOreGiornaliere(worksheet) {
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 1) {
            return;
        }

        const nome = row.getCell(1).value;
        if (
            typeof nome === 'string' &&
            nome.trim().toUpperCase() === 'TOTALI MENSILI'
        ) {
            return;
        }

        const secLavorati = calcolaOreLavorate(
            row.getCell(4).value,
            row.getCell(5).value,
            row.getCell(6).value,
            row.getCell(7).value
        );

        const secStraordinari = Math.max(
            0,
            secLavorati - SOGLIA_8_ORE
        );

        row.getCell(8).value = secondsToTime(secLavorati);
        row.getCell(9).value = secondsToTime(secStraordinari);
    });
}
// ============================================================================
// GESTIONE PRESENZE
// ============================================================================

async function gestisciPresenze(
    uid,
    nomeDipendente
) {
    // ------------------------------------------------------------
    // DATA E ORA CORRENTI
    // ------------------------------------------------------------
    //
    // FIX 2: l'orario viene fissato quando la richiesta ARRIVA, non quando
    // la coda del lock la libera. Se il NAS è lento, altrimenti le timbrature
    // in attesa verrebbero registrate con un orario successivo.
    //
    const oraAttuale =
        getDataOraItalia();

    return conLockExcel(async () => {

        const anno =
            oraAttuale.anno;

        const meseNumero =
            oraAttuale.mese;

        const meseNome =
            MESI[meseNumero - 1];

        const dataOggi =
            oraAttuale.data;

        const dataOggiKey =
            oraAttuale.dataKey;

        const oraTimbratura =
            oraAttuale.ora;

        // ------------------------------------------------------------
        // CARTELLA ANNO
        // ------------------------------------------------------------

        await fs.access(NAS_BASE_PATH);

        const cartellaAnno =
            path.join(
                NAS_BASE_PATH,
                anno
            );

        await fs.ensureDir(
            cartellaAnno
        );

        // ------------------------------------------------------------
        // FILE MENSILE
        // ------------------------------------------------------------

        const nomeFile =
            `Presenze_${meseNome}_${anno}.xlsx`;

        const percorsoFile =
            path.join(
                cartellaAnno,
                nomeFile
            );

        console.log('');
        console.log(
            `📅 Data: ${dataOggi}`
        );

        console.log(
            `🔑 DataKey: ${dataOggiKey}`
        );

        console.log(
            `📄 File: ${percorsoFile}`
        );

        // ------------------------------------------------------------
        // CARICA / CREA WORKBOOK
        // ------------------------------------------------------------

        const workbook =
            new ExcelJS.Workbook();

        let worksheet;

        if (
            await fs.pathExists(
                percorsoFile
            )
        ) {
            console.log(
                '📂 File Excel esistente: caricamento...'
            );

            await workbook.xlsx.readFile(
                percorsoFile
            );

            worksheet =
                workbook.getWorksheet(
                    'Presenze'
                );

            // Se il foglio non esiste,
            // lo creiamo.
            if (!worksheet) {
                console.warn(
                    '⚠️ Foglio "Presenze" non trovato. Creazione...'
                );

                worksheet =
                    workbook.addWorksheet(
                        'Presenze'
                    );

                creaStrutturaWorksheet(
                    worksheet
                );
            }
        }

        else {
            console.log(
                '🆕 File Excel non esistente: creazione...'
            );

            worksheet =
                workbook.addWorksheet(
                    'Presenze'
                );

            creaStrutturaWorksheet(
                worksheet
            );
        }

        // ------------------------------------------------------------
        // RIMUOVI TEMPORANEAMENTE I TOTALI
        // ------------------------------------------------------------
        //
        // Questo è fondamentale:
        // la riga TOTALI MENSILI non deve essere
        // considerata durante la ricerca del dipendente.
        //
        // ------------------------------------------------------------

        rimuoviRigaTotali(
            worksheet
        );

        // Rimuove i segnaposto finali prima di accodare i dati reali.
        rimuoviSegnapostoFinali(worksheet);

        // ------------------------------------------------------------
        // CERCA LA RIGA DEL DIPENDENTE
        // NELLA GIORNATA CORRENTE
        // ------------------------------------------------------------

        let targetRow = null;

        worksheet.eachRow(
            (row, rowNumber) => {

                if (rowNumber <= 1) {
                    return;
                }

                const valoreUid =
                    row.getCell(2).value;

                const valoreData =
                    row.getCell(3).value;

                // Normalizzazione UID
                const rowUid =
                    String(
                        valoreUid || ''
                    ).trim().toUpperCase();

                // Normalizzazione data.
                //
                // Può essere:
                // 24/09/2026
                // 2026-09-24
                // Date
                //
                const rowDataKey =
                    normalizzaDataExcel(
                        valoreData
                    );

                // FIX 9: rimosso il log di debug stampato per OGNI riga a
                // ogni timbratura (centinaia di righe di log per passaggio
                // a fine mese, con crescita incontrollata dei log Docker).

                // ----------------------------------------------------
                // CONFRONTO
                // ----------------------------------------------------
                //
                // STESSO UID
                // +
                // STESSO GIORNO
                //
                // = aggiorna questa riga
                //
                // ----------------------------------------------------

                if (
                    rowUid === uid &&
                    rowDataKey === dataOggiKey
                ) {
                    targetRow = row;

                    console.log(
                        `✅ Trovata riga odierna: ${rowNumber}`
                    );
                }
            }
        );

        // ------------------------------------------------------------
        // NUOVA GIORNATA
        // ------------------------------------------------------------

        if (!targetRow) {

            console.log(
                '🆕 Nessuna riga per oggi: creazione nuova riga.'
            );

            targetRow = worksheet.addRow([
                nomeDipendente,
                uid,
                dataOggi,
                oraTimbratura,
                '',
                '',
                '',
                '00:00:00',
                '00:00:00'
            ]);

            console.log(
                `➕ Nuova riga creata: ${targetRow.number}`
            );
        }

        // ------------------------------------------------------------
        // GIORNATA ESISTENTE
        // ------------------------------------------------------------

        else {

            console.log(
                `🔄 Aggiornamento riga ${targetRow.number}`
            );

            // --------------------------------------------------------
            // FIX 3: ANTI DOPPIA LETTURA
            // --------------------------------------------------------
            //
            // Se l'ultima timbratura è troppo recente, quella attuale
            // viene ignorata (nulla viene modificato né salvato).
            //
            // --------------------------------------------------------

            if (MIN_INTERVALLO_TIMBRATURE_SEC > 0) {

                const ultimaTimbraturaSec = Math.max(
                    timeToSeconds(targetRow.getCell(4).value),
                    timeToSeconds(targetRow.getCell(5).value),
                    timeToSeconds(targetRow.getCell(6).value),
                    timeToSeconds(targetRow.getCell(7).value)
                );

                const differenzaSec =
                    timeToSeconds(oraTimbratura) -
                    ultimaTimbraturaSec;

                if (
                    differenzaSec >= 0 &&
                    differenzaSec < MIN_INTERVALLO_TIMBRATURE_SEC
                ) {
                    console.warn(
                        `⚠️ Timbratura ravvicinata ignorata per ${nomeDipendente} ` +
                        `(${uid}): ${differenzaSec}s dall'ultimo passaggio`
                    );

                    return {
                        success: false,
                        limitReached: false,
                        duplicate: true,
                        file: percorsoFile
                    };
                }
            }

            const p2 =
                targetRow.getCell(5);

            const p3 =
                targetRow.getCell(6);

            const p4 =
                targetRow.getCell(7);

            // --------------------------------------------------------
            // 2° PASSAGGIO
            // --------------------------------------------------------

            if (!p2.value) {

                p2.value =
                    oraTimbratura;

                console.log(
                    '🟡 2° passaggio registrato.'
                );
            }

            // --------------------------------------------------------
            // 3° PASSAGGIO
            // --------------------------------------------------------

            else if (!p3.value) {

                p3.value =
                    oraTimbratura;

                console.log(
                    '🟠 3° passaggio registrato.'
                );
            }

            // --------------------------------------------------------
            // 4° PASSAGGIO
            // --------------------------------------------------------

            else if (!p4.value) {

                p4.value =
                    oraTimbratura;

                console.log(
                    '🔴 4° passaggio registrato.'
                );
            }

            // --------------------------------------------------------
            // OLTRE 4 PASSAGGI
            // --------------------------------------------------------

            else {

                console.warn(
                    `⚠️ Limite di 4 timbrature raggiunto per ${nomeDipendente} (${uid})`
                );

                // Ricrea i totali senza modificare
                // la riga della giornata.
                aggiungiTotaliMensili(
                    worksheet
                );

                await salvaWorkbook(
                    workbook,
                    percorsoFile
                );

                return {
                    success: false,
                    limitReached: true,
                    duplicate: false,
                    file: percorsoFile
                };
            }
        }

        // ------------------------------------------------------------
        // LETTURA PASSAGGI
        // ------------------------------------------------------------

        const p1 =
            targetRow.getCell(4).value;

        const p2 =
            targetRow.getCell(5).value;

        const p3 =
            targetRow.getCell(6).value;

        const p4 =
            targetRow.getCell(7).value;

        // ------------------------------------------------------------
        // CALCOLO ORE
        // ------------------------------------------------------------

        const secLavorati =
            calcolaOreLavorate(
                p1,
                p2,
                p3,
                p4
            );

        const secStraordinari =
            secLavorati >
            SOGLIA_8_ORE
                ? secLavorati -
                  SOGLIA_8_ORE
                : 0;

        targetRow.getCell(8).value =
            secondsToTime(
                secLavorati
            );

        targetRow.getCell(9).value =
            secondsToTime(
                secStraordinari
            );

        targetRow.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };

        // ------------------------------------------------------------
        // STAMPA DEBUG
        // ------------------------------------------------------------

        console.log(
            `⏱️ P1: ${p1 || '-'}`
        );

        console.log(
            `⏱️ P2: ${p2 || '-'}`
        );

        console.log(
            `⏱️ P3: ${p3 || '-'}`
        );

        console.log(
            `⏱️ P4: ${p4 || '-'}`
        );

        console.log(
            `⏱️ Totale: ${secondsToTime(secLavorati)}`
        );

        console.log(
            `⏱️ Straordinario: ${secondsToTime(secStraordinari)}`
        );

        // ------------------------------------------------------------
        // TOTALI MENSILI
        // ------------------------------------------------------------

        aggiungiTotaliMensili(
            worksheet
        );

        // ------------------------------------------------------------
        // SALVATAGGIO
        // ------------------------------------------------------------

        console.log(
            `💾 Salvataggio Excel: ${percorsoFile}`
        );

        // FIX 1: salvataggio atomico (file temporaneo + rename).
        await salvaWorkbook(
            workbook,
            percorsoFile
        );

        console.log(
            '✅ Excel salvato correttamente.'
        );

        return {
            success: true,
            limitReached: false,
            duplicate: false,
            file: percorsoFile
        };
    });
}

// ============================================================================
// API NFC
// ============================================================================

app.post(
    '/api/nfc',
    async (req, res, next) => {

        try {

            const uid =
                normalizzaUid(
                    req.body?.uid
                );

            if (!uid) {

                return res.status(400).json({
                    status: 'error',
                    message:
                        'UID mancante o non valido'
                });
            }

            const nomeDipendente =
                ottieniNomeDipendente(
                    uid
                );

            const timestamp =
                getDataOraItalia().ora;

            console.log('');
            console.log(
                '================================='
            );

            console.log(
                `🟢 [${timestamp}] TIMBRATURA RICEVUTA`
            );

            console.log(
                `👤 Dipendente: ${nomeDipendente}`
            );

            console.log(
                `🆔 UID: ${uid}`
            );

            const esito =
                await gestisciPresenze(
                    uid,
                    nomeDipendente
                );

            // --------------------------------------------------------
            // TIMBRATURA RAVVICINATA IGNORATA (FIX 3)
            // --------------------------------------------------------

            if (
                esito.duplicate
            ) {

                console.log(
                    '⚠️ Risposta ESP32: timbratura ravvicinata ignorata.'
                );

                console.log(
                    '================================='
                );

                return res.status(200).json({
                    status:
                        'duplicate_ignored',

                    dipendente:
                        nomeDipendente,

                    receivedUid:
                        uid,

                    message:
                        'Timbratura ravvicinata ignorata'
                });
            }

            // --------------------------------------------------------
            // LIMITE 4 TIMBRATURE
            // --------------------------------------------------------

            if (
                esito.limitReached
            ) {

                console.log(
                    '⚠️ Risposta ESP32: limite raggiunto.'
                );

                console.log(
                    '================================='
                );

                return res.status(200).json({
                    status:
                        'limit_reached',

                    dipendente:
                        nomeDipendente,

                    receivedUid:
                        uid,

                    message:
                        'Limite massimo di 4 timbrature giornaliere raggiunto'
                });
            }

            // --------------------------------------------------------
            // SUCCESSO
            // --------------------------------------------------------

            console.log(
                `💾 Salvato su Excel NAS: ${esito.file}`
            );

            console.log(
                '================================='
            );

            return res.status(200).json({
                status:
                    'success',

                dipendente:
                    nomeDipendente,

                receivedUid:
                    uid
            });
        }

        catch (error) {
            next(error);
        }
    }
);

// ============================================================================
// API DIPENDENTI - INSERIMENTO / MODIFICA
// ============================================================================

app.post(
    '/api/dipendenti',
    (req, res) => {

        try {

            const uid =
                normalizzaUid(
                    req.body?.uid
                );

            const nome =
                typeof req.body?.nome === 'string'
                    ? req.body.nome.trim()
                    : '';

            if (!uid || !nome) {

                return res.status(400).json({
                    status: 'error',
                    message:
                        'Parametri UID e Nome richiesti'
                });
            }

            if (nome.length > 150) {

                return res.status(400).json({
                    status: 'error',
                    message:
                        'Nome troppo lungo'
                });
            }

            stmtUpsertDipendente.run(
                uid,
                nome
            );

            console.log(
                `✅ Dipendente registrato: [${uid}] -> ${nome}`
            );

            return res.status(200).json({
                status: 'success',
                message:
                    `UID ${uid} associato a ${nome}`
            });
        }

        catch (error) {

            console.error(
                '❌ Errore registrazione dipendente:',
                error
            );

            return res.status(500).json({
                status: 'error',
                message:
                    'Errore durante la registrazione del dipendente'
            });
        }
    }
);

// ============================================================================
// API DIPENDENTI - LISTA
// ============================================================================

app.get(
    '/api/dipendenti',
    (req, res) => {

        try {

            const lista =
                stmtGetDipendenti.all();

            return res.status(200).json(
                lista
            );
        }

        catch (error) {

            console.error(
                '❌ Errore lettura dipendenti:',
                error
            );

            return res.status(500).json({
                status: 'error',
                message:
                    'Errore durante la lettura dei dipendenti'
            });
        }
    }
);

// ============================================================================
// HEALTH CHECK
// ============================================================================

app.get(
    '/health',
    async (req, res) => {

        try {

            // Controlla che il NAS sia effettivamente
            // accessibile (anche in scrittura) dal container.
            // FIX 12: verifica anche il permesso di scrittura.
            await fs.access(
                NAS_BASE_PATH,
                fs.constants.R_OK | fs.constants.W_OK
            );

            // Controlla anche che il DB risponda.
            db.prepare(
                'SELECT 1'
            ).get();

            // FIX 14: l'endpoint è pubblico, quindi non espone più i
            // percorsi interni di database e NAS.
            return res.status(200).json({
                status: 'ok',

                timezone:
                    TIMEZONE
            });
        }

        catch (error) {

            console.error(
                '❌ Health check fallito:',
                error
            );

            return res.status(503).json({
                status: 'error',

                message:
                    'NAS o database non disponibili'
            });
        }
    }
);

// ============================================================================
// ERROR HANDLER
// ============================================================================

app.use(
    (error, req, res, next) => {

        console.error(
            '❌ Errore non gestito:',
            error
        );

        if (res.headersSent) {
            return next(error);
        }

        // FIX 10: gli errori del client (es. JSON malformato, payload troppo
        // grande) devono restituire 4xx, non 500.
        const status =
            Number.isInteger(error.status) &&
            error.status >= 400 &&
            error.status < 500
                ? error.status
                : 500;

        return res.status(status).json({
            status: 'error',
            message:
                status === 500
                    ? 'Errore interno del server'
                    : 'Richiesta non valida'
        });
    }
);

// ============================================================================
// ARRESTO GRADUALE (FIX 11)
// ============================================================================
//
// Docker invia SIGTERM allo stop/riavvio del container. Node, come PID 1,
// senza handler non si ferma in modo pulito: una scrittura Excel in corso
// verrebbe interrotta. Qui si smette di accettare richieste, si attende
// il completamento della coda Excel e si chiude il database.
//
// ============================================================================

let server = null;
let arrestoInCorso = false;

function arrestoGraduale(segnale) {
    if (arrestoInCorso) {
        return;
    }

    arrestoInCorso = true;

    console.log(
        `🛑 Segnale ${segnale} ricevuto: arresto in corso...`
    );

    // Uscita forzata se l'arresto non termina entro 15 secondi.
    setTimeout(() => {
        console.error('❌ Arresto forzato per timeout.');
        process.exit(1);
    }, 15000).unref();

    const concludi = async () => {
        try {
            await excelQueue;
        } catch (_) {
            // L'errore è già stato gestito dalla richiesta.
        }

        try {
            db.close();
        } catch (_) {
            // Ignorato: il processo sta terminando.
        }

        console.log('👋 Server arrestato.');
        process.exit(0);
    };

    if (server) {
        server.close(() => {
            concludi();
        });

        if (typeof server.closeIdleConnections === 'function') {
            server.closeIdleConnections();
        }
    } else {
        concludi();
    }
}

process.on('SIGTERM', () => arrestoGraduale('SIGTERM'));
process.on('SIGINT', () => arrestoGraduale('SIGINT'));

// ============================================================================
// AVVIO SERVER
// ============================================================================

async function avviaServer() {
    try {
        await fs.access(DB_PATH);

        // FIX 12: il NAS deve essere anche scrivibile, non solo raggiungibile
        // (un mount in sola lettura passerebbe il controllo ma fallirebbe
        // a ogni timbratura).
        await fs.access(
            NAS_BASE_PATH,
            fs.constants.R_OK | fs.constants.W_OK
        );

        db.prepare('SELECT 1').get();

        console.log('');
        console.log('==============================================');
        console.log('        CONFIGURAZIONE SERVER OK');
        console.log('==============================================');
        console.log(`🗄️ Database: ${DB_PATH}`);
        console.log(`📁 NAS: ${NAS_BASE_PATH}`);
        console.log(`🌍 Timezone: ${TIMEZONE}`);
        console.log(`🔌 Porta: ${PORT}`);
        console.log(
            `⏳ Intervallo minimo tra timbrature: ${MIN_INTERVALLO_TIMBRATURE_SEC}s`
        );
        console.log('----------------------------------------------');
        console.log('✓ Database presente e valido');
        console.log('✓ NAS raggiungibile e scrivibile');

        if (API_KEY) {
            console.log('✓ Autenticazione API attiva (x-api-key)');
        } else {
            console.warn(
                '⚠️ API_KEY non impostata: le rotte /api/* sono PUBBLICHE.'
            );
        }

        console.log('==============================================');

        server = app.listen(
            PORT,
            '0.0.0.0',
            () => {
                console.log(
                    `🚀 Server avviato sulla porta ${PORT}`
                );
            }
        );

    } catch (error) {
        console.error('');
        console.error('==============================================');
        console.error('        SERVER NON AVVIATO');
        console.error('==============================================');
        console.error(error.message);
        console.error('==============================================');

        process.exit(1);
    }
}

avviaServer();
