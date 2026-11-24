'use strict';

const express = require('express');
const crypto = require('crypto');
const fs = require('fs-extra');
const path = require('path');
const ExcelJS = require('exceljs');
const Database = require('better-sqlite3');

const app = express();

const PORT = Number(process.env.PORT || 3000);

// CONFIGURATION
// IMPORTANT:
// The NAS must be mounted in the Docker container.
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
// Set in Docker:
// TZ=Europe/Rome
const TIMEZONE = process.env.TZ || 'Europe/Rome';

const SOGLIA_8_ORE = 8 * 3600;
// FIX 3: minimum interval (in seconds) between two clock-ins by the same
// employee on the same day. Prevents a card held near the
// reader from generating multiple close-together scans (closing the day with 0 hours).
// Set MIN_INTERVALLO_TIMBRATURE_SEC=0 to disable the check.
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
// FIX 4: optional API key. If set, all /api/* routes require
// the "x-api-key" header with the same value.
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

// EXPRESS
app.disable('x-powered-by');

app.use(express.json({
    limit: '100kb'
}));


// API AUTHENTICATION (FIX 4)
//
// Without authentication, anyone on the network can submit fake clock-ins or
// rename employees. This check is active only if API_KEY is set,
// so the existing ESP32 firmware continues to work until it is
// updated to send the x-api-key header.
//
// /health remains public (used by Docker health checks).

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

// SQLITE DATABASE
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
// FIX 13: the "ON CONFLICT(uid)" UPSERT works only if uid is a PRIMARY KEY
// or has a UNIQUE constraint/index. Otherwise, every registration
// would fail at runtime, so it is better to catch this at startup.
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
        'UID non presente nel database: [${uid}]`
    );

    return `Sconosciuto (${uid})`;
}

// EXCEL WRITE LOCK
//
// Prevents two concurrent NFC requests from reading and writing
// the same Excel file at the same time.
//
// WARNING:
// This lock works within the same Node.js process.
// If you use multiple Node.js containers on the same Excel file at the same time,
// a distributed locking system will be needed.
let excelQueue = Promise.resolve();

function conLockExcel(fn) {
    const operazione = excelQueue.then(() => fn());
    // Keep the queue active even if an operation fails.
    excelQueue = operazione.catch(() => {});

    return operazione;
}

// DATE AND TIME
//
// Returns:
// - year
// - month
// - day
// - dataKey -> stable format YYYY-MM-DD
// - date -> displayed format DD/MM/YYYY
// - time -> HH:mm:ss
//
// dataKey is essential for correctly comparing days.
// Do NOT use the text displayed in Excel directly.
function getDataOraItalia() {
    const now = new Date();
    // FIX 5: use hourCycle 'h23' instead of hour12:false. With hour12:false, some
    // ICU versions return "24" at midnight (24:00:05 instead of 00:00:05).
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
        // Technical identifier for the day.
        // Example: 2026-09-24
        dataKey: `${anno}-${String(mese).padStart(2, '0')}-${giorno}`,
        // Format displayed in Excel.
        // Example: 24/09/2026
        data: `${giorno}/${String(mese).padStart(2, '0')}/${anno}`,
        // Clock-in time.
        ora: `${valori.hour}:${valori.minute}:${valori.second}`
    };
}

// EXCEL DATE CONVERSION
//
// ExcelJS may return:
// - string
// - Date
// - other values.
//
// This function converts everything to:
// YYYY-MM-DD
//
// Examples:
//
// 24/09/2026 -> 2026-09-24
// 2026-09-24 -> 2026-09-24
// Date(...)  -> 2026-09-24
function normalizzaDataExcel(value) {
    if (!value) {
        return null;
    }
    // Case 1: ExcelJS returns an actual Date
    if (value instanceof Date) {
        if (!Number.isFinite(value.getTime())) {
            return null;
        }
        // FIX 6: ExcelJS converts the Excel serial value to a "naive" UTC Date
        // (UTC midnight of the day written in the cell). Applying the server
        // timezone could shift the date by one day, so the UTC components
        // are used directly.
        return value.toISOString().slice(0, 10);
    }
    // Case 2: text value
    const valore = String(value).trim();

    if (!valore) {
        return null;
    }
    // ISO format:
    //
    // 2026-09-24
    if (/^\d{4}-\d{2}-\d{2}$/.test(valore)) {
        return valore;
    }
    // Italian format:
    //
    // 24/09/2026
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

// TIME UTILITIES
function timeToSeconds(timeStr) {
    if (timeStr === null || timeStr === undefined || timeStr === '') {
        return 0;
    }
    // ExcelJS often returns cells formatted as time as Date objects,
    // not strings. The UTC components represent the time stored
    // in the Excel serial value without applying the server timezone.
    if (timeStr instanceof Date) {
        if (!Number.isFinite(timeStr.getTime())) {
            return 0;
        }

        return timeStr.getUTCHours() * 3600 +
            timeStr.getUTCMinutes() * 60 +
            timeStr.getUTCSeconds();
    }
    // Also supports an Excel numeric value representing a fraction
    // of a day (e.g. 0.5 = 12:00).
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
    // FIX 7: use the already-trimmed value (previously timeStr was split).
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

// UID VALIDATION
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

// EXCEL STRUCTURE
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
    // FIX 8: colors in ExcelJS must be expressed as 8-digit ARGB
    // (with 6 digits, Excel may interpret them incorrectly).
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

// ATOMIC SAVE (FIX 1)
//
// Writing directly to the final file with writeFile(), if interrupted (crash,
// container restart, NAS disconnect), leaves a truncated .xlsx
// file, causing the entire month to be lost.
// First write to a temporary file in the same folder, then
// rename it: the rename is atomic, and the original file remains intact until
// the new version is complete.
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

// MONTHLY TOTALS ROW
function rimuoviRigaTotali(worksheet) {
    // A file already affected by this issue may contain multiple totals rows.
    // Removing only one at each clock-in would cause them to accumulate over time.
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
// Removes only trailing empty rows (including template rows with
// 00:00:00 in the totals columns). ExcelJS considers them used rows:
// leaving them would cause new days to be appended much farther down.
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

// MONTHLY TOTALS CALCULATION
function calcolaTotaliMensili(worksheet) {
    let totMeseLavoratoSec = 0;
    let totMeseStraordinariSec = 0;

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 1) {
            return;
        }
        // Never count the MONTHLY TOTALS row.
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

// ADD MONTHLY TOTALS
function aggiungiTotaliMensili(worksheet) {
    // First remove any existing totals.
    rimuoviRigaTotali(worksheet);
    // Previously saved values may be incorrect or have a type
    // other than a string in Excel. Recalculate each day from the raw clock-ins.
    ricalcolaOreGiornaliere(worksheet);

    const totali =
        calcolaTotaliMensili(worksheet);
    // Explicit array: the keys defined in worksheet.columns are not
    // preserved in the XLSX file and are unavailable after readFile().
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

// CALCULATE HOURS WORKED
//
// 2 clock-ins:
//
// P1 = Clock-in
// P2 = Clock-out
//
// Total:
// P2 - P1
//
//
//
// 4 clock-ins:
//
// P1 = Clock-in
// P2 = Clock-out pausa
// P3 = Break end
// P4 = Clock-out
//
// Total:
// (P2 - P1) + (P4 - P3)
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
    // 2 clock-ins
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
    // 4 clock-ins
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
    // 1 or 3 clock-ins
    //
    // The total cannot be determined yet.
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

// ATTENDANCE MANAGEMENT
async function gestisciPresenze(
    uid,
    nomeDipendente
) {
    // CURRENT DATE AND TIME
    //
    // FIX 2: the time is fixed when the request ARRIVES, not when
    // the lock queue releases it. Otherwise, if the NAS is slow, pending clock-ins
    // would be recorded with a later time.
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
        // YEAR FOLDER
        await fs.access(NAS_BASE_PATH);

        const cartellaAnno =
            path.join(
                NAS_BASE_PATH,
                anno
            );

        await fs.ensureDir(
            cartellaAnno
        );
        // MONTHLY FILE
        const nomeFile =
            `Presenze_${meseNome}_${anno}.xlsx`;

        const percorsoFile =
            path.join(
                cartellaAnno,
                nomeFile
            );

        console.log('');
        console.log(
            `Date: ${dataOggi}`
        );

        console.log(
            `DateKey: ${dataOggiKey}`
        );

        console.log(
            `File: ${percorsoFile}`
        );
        // LOAD / CREATE WORKBOOK
        const workbook =
            new ExcelJS.Workbook();

        let worksheet;

        if (
            await fs.pathExists(
                percorsoFile
            )
        ) {
            console.log(
                'Existing Excel file: loading...'
            );

            await workbook.xlsx.readFile(
                percorsoFile
            );

            worksheet =
                workbook.getWorksheet(
                    'Presenze'
                );
            // If the worksheet does not exist,
            // create it.
            if (!worksheet) {
                console.warn(
                    'Foglio "Presenze" non trovato. Creazione...'
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
                'Excel file does not exist: creating it...'
            );

            worksheet =
                workbook.addWorksheet(
                    'Presenze'
                );

            creaStrutturaWorksheet(
                worksheet
            );
        }
        
        // TEMPORARILY REMOVE TOTALS
        //
        // This is essential:
        // the MONTHLY TOTALS row must not be
        // included when searching for the employee.
        rimuoviRigaTotali(
            worksheet
        );
        // Remove trailing placeholders before appending real data.
        rimuoviSegnapostoFinali(worksheet);
        // FIND THE EMPLOYEE ROW
        // FOR THE CURRENT DAY
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
                // UID normalization
                const rowUid =
                    String(
                        valoreUid || ''
                    ).trim().toUpperCase();
                // Date normalization.
                //
                // May be:
                // 24/09/2026
                // 2026-09-24
                // Date
                const rowDataKey =
                    normalizzaDataExcel(
                        valoreData
                    );

                // FIX 9: removed the debug log printed for EVERY row on
                // each clock-in (hundreds of log lines per scan
                // at the end of the month, causing uncontrolled growth of Docker logs).
                // COMPARISON
                //
                // SAME UID
                // +
                // SAME DAY
                //
                // = update this row
                if (
                    rowUid === uid &&
                    rowDataKey === dataOggiKey
                ) {
                    targetRow = row;

                    console.log(
                        `Found today’s row: ${rowNumber}`
                    );
                }
            }
        );
        // NEW DAY
        if (!targetRow) {

            console.log(
                'No row for today: creating a new row.'
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
                `➕ New row created: ${targetRow.number}`
            );
        }
        // EXISTING DAY
        else {

            console.log(
                'Updating row ${targetRow.number}`
            );
            // FIX 3: DUPLICATE SCAN PREVENTION
            //
            // If the last clock-in was too recent, the current one
            // is ignored (nothing is changed or saved).
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
                        `Timbratura ravvicinata ignorata per ${nomeDipendente} ` +
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
            // 2ND SCAN
            if (!p2.value) {

                p2.value =
                    oraTimbratura;

                console.log(
                    '2nd scan recorded.'
                );
            }
            // 3RD SCAN
            else if (!p3.value) {

                p3.value =
                    oraTimbratura;

                console.log(
                    '3rd scan recorded.'
                );
            }
            // 4TH SCAN
            else if (!p4.value) {

                p4.value =
                    oraTimbratura;

                console.log(
                    '4th scan recorded.'
                );
            }
            // MORE THAN 4 SCANS
            else {

                console.warn(
                    `Limite di 4 timbrature raggiunto per ${nomeDipendente} (${uid})`
                );
                // Recreate the totals without changing
                // the day’s row.
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
        // READ SCANS
        const p1 =
            targetRow.getCell(4).value;

        const p2 =
            targetRow.getCell(5).value;

        const p3 =
            targetRow.getCell(6).value;

        const p4 =
            targetRow.getCell(7).value;
        // CALCULATE HOURS
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
        // DEBUG OUTPUT
        console.log(
            `P1: ${p1 || '-'}`
        );

        console.log(
            `P2: ${p2 || '-'}`
        );

        console.log(
            `P3: ${p3 || '-'}`
        );

        console.log(
            `P4: ${p4 || '-'}`
        );

        console.log(
            `Total: ${secondsToTime(secLavorati)}`
        );

        console.log(
            `Overtime: ${secondsToTime(secStraordinari)}`
        );
        // MONTHLY TOTALS
        aggiungiTotaliMensili(
            worksheet
        );
        // SAVING
        console.log(
            `Saving Excel file: ${percorsoFile}`
        );
        // FIX 1: atomic save (temporary file + rename).
        await salvaWorkbook(
            workbook,
            percorsoFile
        );

        console.log(
            'Excel file saved successfully.'
        );

        return {
            success: true,
            limitReached: false,
            duplicate: false,
            file: percorsoFile
        };
    });
}
// NFC API
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
                `[${timestamp}] CLOCK-IN RECEIVED`
            );

            console.log(
                `Employee: ${nomeDipendente}`
            );

            console.log(
                `UID: ${uid}`
            );

            const esito =
                await gestisciPresenze(
                    uid,
                    nomeDipendente
                );
            // CLOSE-TOGETHER CLOCK-IN IGNORED (FIX 3)
            if (
                esito.duplicate
            ) {

                console.log(
                    'ESP32 response: close-together clock-in ignored.'
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
            // 4 CLOCK-IN LIMIT
            if (
                esito.limitReached
            ) {

                console.log(
                    'ESP32 response: limit reached.'
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
            // SUCCESS
            console.log(
                `Saved to NAS Excel file: ${esito.file}`
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

// EMPLOYEE API - ADD / UPDATE
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
                `Employee registered: [${uid}] -> ${nome}`
            );

            return res.status(200).json({
                status: 'success',
                message:
                    `UID ${uid} associato a ${nome}`
            });
        }

        catch (error) {

            console.error(
                'Errore registrazione dipendente:',
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

// EMPLOYEE API - LIST
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
                'Errore lettura dipendenti:',
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

// HEALTH CHECK
app.get(
    '/health',
    async (req, res) => {

        try {
            // Check that the NAS is actually
            // accessible (including for writing) from the container.
            // FIX 12: also check write permissions.
            await fs.access(
                NAS_BASE_PATH,
                fs.constants.R_OK | fs.constants.W_OK
            );
            // Also check that the database responds.
            db.prepare(
                'SELECT 1'
            ).get();
            // FIX 14: the endpoint is public, so it no longer exposes the
            // internal database and NAS paths.
            return res.status(200).json({
                status: 'ok',

                timezone:
                    TIMEZONE
            });
        }

        catch (error) {

            console.error(
                'Health check fallito:',
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

// ERROR HANDLER
app.use(
    (error, req, res, next) => {

        console.error(
            'Errore non gestito:',
            error
        );

        if (res.headersSent) {
            return next(error);
        }
        // FIX 10: client errors (e.g. malformed JSON, oversized payloads)
        // must return 4xx, not 500.
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

// GRACEFUL SHUTDOWN (FIX 11)
//
// Docker sends SIGTERM when stopping/restarting the container. Node, as PID 1,
// does not shut down cleanly without a handler: an ongoing Excel write
// would be interrupted. Here, the server stops accepting requests, waits for
// the Excel queue to finish, and closes the database.
let server = null;
let arrestoInCorso = false;

function arrestoGraduale(segnale) {
    if (arrestoInCorso) {
        return;
    }

    arrestoInCorso = true;

    console.log(
        `Segnale ${segnale} ricevuto: arresto in corso...`
    );
    // Force exit if shutdown does not finish within 15 seconds.
    setTimeout(() => {
        console.error('Arresto forzato per timeout.');
        process.exit(1);
    }, 15000).unref();

    const concludi = async () => {
        try {
            await excelQueue;
        } catch (_) {
            // The error has already been handled by the request.
        }

        try {
            db.close();
        } catch (_) {
            // Ignored: the process is shutting down.
        }

        console.log('Server stopped.');
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

// SERVER STARTUP
async function avviaServer() {
    try {
        await fs.access(DB_PATH);
        // FIX 12: the NAS must also be writable, not just reachable
        // (a read-only mount would pass the check but fail
        // on every clock-in).
        await fs.access(
            NAS_BASE_PATH,
            fs.constants.R_OK | fs.constants.W_OK
        );

        db.prepare('SELECT 1').get();

        console.log('');
        console.log('==============================================');
        console.log('        SERVER CONFIGURATION OK');
        console.log('==============================================');
        console.log(`Database: ${DB_PATH}`);
        console.log(`NAS: ${NAS_BASE_PATH}`);
        console.log(`Timezone: ${TIMEZONE}`);
        console.log(`Port: ${PORT}`);
        console.log(
            `Minimum interval between clock-ins: ${MIN_INTERVALLO_TIMBRATURE_SEC}s`
        );
        console.log('----------------------------------------------');
        console.log('✓ Database found and valid');
        console.log('✓ NAS reachable and writable');

        if (API_KEY) {
            console.log('✓ API authentication enabled (x-api-key)');
        } else {
            console.warn(
                'API_KEY non impostata: le rotte /api/* sono PUBBLICHE.'
            );
        }

        console.log('==============================================');

        server = app.listen(
            PORT,
            '0.0.0.0',
            () => {
                console.log(
                    `Server started on port ${PORT}`
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
