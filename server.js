'use strict';

const express = require('express');
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

    const parts = new Intl.DateTimeFormat('it-IT', {
        timeZone: TIMEZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
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
        const parts = new Intl.DateTimeFormat('it-IT', {
            timeZone: TIMEZONE,
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        }).formatToParts(value);

        const valori = {};

        for (const part of parts) {
            valori[part.type] = part.value;
        }

        return `${valori.year}-${valori.month}-${valori.day}`;
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
    if (!timeStr) {
        return 0;
    }

    if (typeof timeStr !== 'string') {
        return 0;
    }

    const parts = timeStr.split(':');

    if (parts.length < 2 || parts.length > 3) {
        return 0;
    }

    const h = Number(parts[0]);
    const m = Number(parts[1]);
    const s = Number(parts[2] || 0);

    if (
        !Number.isFinite(h) ||
        !Number.isFinite(m) ||
        !Number.isFinite(s)
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

    headerRow.font = {
        bold: true,
        color: {
            argb: 'FFFFFF'
        }
    };

    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
            argb: '1F4E78'
        }
    };

    headerRow.alignment = {
        vertical: 'middle',
        horizontal: 'center'
    };
}

// ============================================================================
// RIGA TOTALI MENSILI
// ============================================================================

function trovaRigaTotali(worksheet) {
    for (
        let rowNumber = worksheet.rowCount;
        rowNumber >= 2;
        rowNumber--
    ) {
        const row = worksheet.getRow(rowNumber);

        if (
            row.getCell(1).value ===
            'TOTALI MENSILI'
        ) {
            return rowNumber;
        }
    }

    return null;
}

function rimuoviRigaTotali(worksheet) {
    const rowNumber = trovaRigaTotali(worksheet);

    if (rowNumber !== null) {
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

    const totali =
        calcolaTotaliMensili(worksheet);

    const rigaTotali = worksheet.addRow({
        dipendente: 'TOTALI MENSILI',
        uid: '',
        data: '',
        p1: '',
        p2: '',
        p3: '',
        p4: '',
        totale: secondsToTime(
            totali.lavorato
        ),
        straordinari: secondsToTime(
            totali.straordinari
        )
    });

    rigaTotali.font = {
        bold: true
    };

    rigaTotali.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
            argb: 'D9E1F2'
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

// ============================================================================
// GESTIONE PRESENZE
// ============================================================================

async function gestisciPresenze(
    uid,
    nomeDipendente
) {
    return conLockExcel(async () => {

        // ------------------------------------------------------------
        // DATA E ORA CORRENTI
        // ------------------------------------------------------------

        const oraAttuale =
            getDataOraItalia();

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
                    ).trim();

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

                // DEBUG
                console.log(
                    `🔎 Riga ${rowNumber}: UID=${rowUid}, DATA=${rowDataKey}`
                );

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

            targetRow =
                worksheet.addRow({
                    dipendente:
                        nomeDipendente,

                    uid:
                        uid,

                    data:
                        dataOggi,

                    p1:
                        oraTimbratura,

                    p2:
                        '',

                    p3:
                        '',

                    p4:
                        '',

                    totale:
                        '00:00:00',

                    straordinari:
                        '00:00:00'
                });

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

                await workbook.xlsx.writeFile(
                    percorsoFile
                );

                return {
                    success: false,
                    limitReached: true,
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

        await workbook.xlsx.writeFile(
            percorsoFile
        );

        console.log(
            '✅ Excel salvato correttamente.'
        );

        return {
            success: true,
            limitReached: false,
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
            // accessibile dal container.
            await fs.access(
                NAS_BASE_PATH
            );

            // Controlla anche che il DB risponda.
            db.prepare(
                'SELECT 1'
            ).get();

            return res.status(200).json({
                status: 'ok',

                database:
                    DB_PATH,

                nas:
                    NAS_BASE_PATH,

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
                    'NAS o database non disponibili',

                nas:
                    NAS_BASE_PATH,

                database:
                    DB_PATH
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

        return res.status(500).json({
            status: 'error',
            message:
                'Errore interno del server'
        });
    }
);

// ============================================================================
// AVVIO SERVER
// ============================================================================

async function avviaServer() {
    try {
        await fs.access(DB_PATH);
        await fs.access(NAS_BASE_PATH);

        db.prepare('SELECT 1').get();

        console.log('');
        console.log('==============================================');
        console.log('        CONFIGURAZIONE SERVER OK');
        console.log('==============================================');
        console.log(`🗄️ Database: ${DB_PATH}`);
        console.log(`📁 NAS: ${NAS_BASE_PATH}`);
        console.log(`🌍 Timezone: ${TIMEZONE}`);
        console.log(`🔌 Porta: ${PORT}`);
        console.log('----------------------------------------------');
        console.log('✓ Database presente e valido');
        console.log('✓ NAS raggiungibile');
        console.log('==============================================');

        app.listen(
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
