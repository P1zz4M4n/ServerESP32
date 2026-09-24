const express = require('express');
const fs = require('fs-extra');
const path = require('path');
const ExcelJS = require('exceljs');
const Database = require('better-sqlite3');

const app = express();
const PORT = 3000;

// Percorso base del NAS (usa la lettera dell'unità mappata o il percorso UNC)
const NAS_BASE_PATH = '\\\\192.168.1.106\\NAS\\tabella presenze\\Lettura NFC';

app.use(express.json());

// =========================================================================
// INIZIALIZZAZIONE DATABASE SQLITE LOCALE
// =========================================================================
const db = new Database('dipendenti.db');

db.exec(`
    CREATE TABLE IF NOT EXISTS dipendenti (
        uid TEXT PRIMARY KEY,
        nome TEXT NOT NULL
    )
`);

function ottieniNomeDipendente(uid) {
    const stmt = db.prepare('SELECT nome FROM dipendenti WHERE uid = ?');
    const dipendente = stmt.get(uid);
    return dipendente ? dipendente.nome : `Sconosciuto (${uid})`;
}

const MESI = [
    'Gennaio', 'Febbraio', 'Marzo', 'Aprile',
    'Maggio', 'Giugno', 'Luglio', 'Agosto',
    'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
];

function timeToSeconds(timeStr) {
    if (!timeStr) return 0;
    const [h, m, s] = timeStr.split(':').map(Number);
    return h * 3600 + m * 60 + (s || 0);
}

function secondsToTime(totalSec) {
    if (totalSec <= 0 || isNaN(totalSec)) return '00:00:00';
    const h = Math.floor(totalSec / 3600).toString().padStart(2, '0');
    const m = Math.floor((totalSec % 3600) / 60).toString().padStart(2, '0');
    const s = Math.floor(totalSec % 60).toString().padStart(2, '0');
    return `${h}:${m}:${s}`;
}

async function gestisciPresenze(uid, nomeDipendente) {
    const oraAttuale = new Date();
    
    const anno = oraAttuale.getFullYear().toString();
    const meseNome = MESI[oraAttuale.getMonth()];
    const dataOggi = oraAttuale.toLocaleDateString('it-IT');
    const oraTimbratura = oraAttuale.toLocaleTimeString('it-IT');

    const cartellaAnno = path.join(NAS_BASE_PATH, anno);
    await fs.ensureDir(cartellaAnno);

    const nomeFile = `Presenze_${meseNome}_${anno}.xlsx`;
    const percorsoFile = path.join(cartellaAnno, nomeFile);

    const workbook = new ExcelJS.Workbook();
    let worksheet;

    if (await fs.pathExists(percorsoFile)) {
        await workbook.xlsx.readFile(percorsoFile);
        worksheet = workbook.getWorksheet('Presenze');
    } else {
        worksheet = workbook.addWorksheet('Presenze');

        worksheet.columns = [
            { header: 'Dipendente', key: 'dipendente', width: 25 },
            { header: 'UID Tessera', key: 'uid', width: 16 },
            { header: 'Data', key: 'data', width: 14 },
            { header: '1° Passaggio (Entrata)', key: 'p1', width: 20 },
            { header: '2° Passaggio (Uscita P.)', key: 'p2', width: 22 },
            { header: '3° Passaggio (Rientro P.)', key: 'p3', width: 22 },
            { header: '4° Passaggio (Uscita)', key: 'p4', width: 20 },
            { header: 'Totale Ore Lavorate', key: 'totale', width: 20 },
            { header: 'Ore Straordinario (>8h)', key: 'straordinari', width: 22 }
        ];

        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    }

    const lastRowIndex = worksheet.rowCount;
    if (lastRowIndex > 1) {
        const lastRow = worksheet.getRow(lastRowIndex);
        if (lastRow.getCell(1).value === 'TOTALI MENSILI') {
            worksheet.spliceRows(lastRowIndex, 1);
        }
    }

    let targetRow = null;
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
            const rowUid = row.getCell(2).value;
            const rowData = row.getCell(3).value;
            if (rowUid === uid && rowData === dataOggi) {
                targetRow = row;
            }
        }
    });

    if (!targetRow) {
        // 1° PASSAGGIO
        targetRow = worksheet.addRow({
            dipendente: nomeDipendente,
            uid: uid,
            data: dataOggi,
            p1: oraTimbratura,
            p2: '', p3: '', p4: '',
            totale: '00:00:00',
            straordinari: '00:00:00'
        });
    } else {
        // GESTIONE DEI PASSAGGI SUCCESSIVI
        if (!targetRow.getCell(5).value) {
            targetRow.getCell(5).value = oraTimbratura; // 2° Passaggio
        } else if (!targetRow.getCell(6).value) {
            targetRow.getCell(6).value = oraTimbratura; // 3° Passaggio
        } else if (!targetRow.getCell(7).value) {
            targetRow.getCell(7).value = oraTimbratura; // 4° Passaggio
        } else {
            // 🛑 OLTRE IL 4° PASSAGGIO: Ripristina i totali e restituisce il flag di blocco
            console.warn(`\n⚠️ Limite raggiunto: 4 timbrature già eseguite oggi per ${nomeDipendente} (${uid})`);
            
            await ripristinaTotaliMensili(worksheet);
            await workbook.xlsx.writeFile(percorsoFile);
            
            return { success: false, limitReached: true, file: percorsoFile };
        }
    }

    // --- CALCOLO ORE E STRAORDINARI ---
    const p1 = targetRow.getCell(4).value;
    const p2 = targetRow.getCell(5).value;
    const p3 = targetRow.getCell(6).value;
    const p4 = targetRow.getCell(7).value;

    let secLavorati = 0;
    if (p1 && p4) {
        const totaleGiornata = timeToSeconds(p4) - timeToSeconds(p1);
        let pausaPranzo = 0;
        if (p2 && p3) pausaPranzo = timeToSeconds(p3) - timeToSeconds(p2);
        secLavorati = Math.max(0, totaleGiornata - pausaPranzo);
    } else if (p1 && p2 && !p3 && !p4) {
        secLavorati = timeToSeconds(p2) - timeToSeconds(p1);
    }

    const SOGLIA_8_ORE = 8 * 3600;
    const secStraordinari = secLavorati > SOGLIA_8_ORE ? secLavorati - SOGLIA_8_ORE : 0;

    targetRow.getCell(8).value = secondsToTime(secLavorati);
    targetRow.getCell(9).value = secondsToTime(secStraordinari);
    targetRow.alignment = { vertical: 'middle', horizontal: 'center' };

    await ripristinaTotaliMensili(worksheet);
    await workbook.xlsx.writeFile(percorsoFile);
    
    return { success: true, limitReached: false, file: percorsoFile };
}

async function ripristinaTotaliMensili(worksheet) {
    let totMeseLavoratoSec = 0;
    let totMeseStraordinariSec = 0;

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
            totMeseLavoratoSec += timeToSeconds(row.getCell(8).value);
            totMeseStraordinariSec += timeToSeconds(row.getCell(9).value);
        }
    });

    const rigaTotali = worksheet.addRow({
        dipendente: 'TOTALI MENSILI',
        uid: '', data: '', p1: '', p2: '', p3: '', p4: '',
        totale: secondsToTime(totMeseLavoratoSec),
        straordinari: secondsToTime(totMeseStraordinariSec)
    });

    rigaTotali.font = { bold: true };
    rigaTotali.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9E1F2' } };
    rigaTotali.alignment = { vertical: 'middle', horizontal: 'center' };
}

// =========================================================================
// ROUTE HTTP
// =========================================================================

app.post('/api/nfc', async (req, res, next) => {
    try {
        const { uid } = req.body;
        if (!uid) return res.status(400).json({ status: 'error', message: 'UID mancante' });

        const nomeDipendente = ottieniNomeDipendente(uid);
        const timestamp = new Date().toLocaleTimeString('it-IT');

        console.log(`\n=================================`);
        console.log(`🟢 [${timestamp}] TIMBRATURA RICEVUTA`);
        console.log(`👤 Dipendente: ${nomeDipendente}`);
        console.log(`🆔 UID: ${uid}`);

        const esito = await gestisciPresenze(uid, nomeDipendente);

        if (esito.limitReached) {
            console.log(`⚠️ Risposta inviata all'ESP32: Limite 4 timbrature raggiunto.`);
            console.log(`=================================`);
            
            // Invia lo stato che l'ESP32 si aspetta per il limite raggiunto
            return res.status(200).json({
                status: 'limit_reached',
                dipendente: nomeDipendente,
                receivedUid: uid,
                message: 'Limite massimo di 4 timbrature giornaliere raggiunto'
            });
        }

        console.log(`💾 Salvato su Excel NAS: ${esito.file}`);
        console.log(`=================================`);

        res.status(200).json({
            status: 'success',
            dipendente: nomeDipendente,
            receivedUid: uid
        });
    } catch (error) {
        next(error);
    }
});

app.post('/api/dipendenti', (req, res) => {
    const { uid, nome } = req.body;

    if (!uid || !nome) {
        return res.status(400).json({ status: 'error', message: 'Parametri UID e Nome richiesti' });
    }

    const stmt = db.prepare(`
        INSERT INTO dipendenti (uid, nome) 
        VALUES (?, ?) 
        ON CONFLICT(uid) DO UPDATE SET nome = excluded.nome
    `);
    
    stmt.run(uid, nome);

    console.log(`\n✅ Dipendente Registrato: [${uid}] -> ${nome}`);
    res.status(200).json({ status: 'success', message: `UID ${uid} associato a ${nome}` });
});

app.get('/api/dipendenti', (req, res) => {
    const stmt = db.prepare('SELECT * FROM dipendenti');
    const lista = stmt.all();
    res.status(200).json(lista);
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Server avviato sulla porta ${PORT}`);
    console.log(`🗄️ Database dipendenti caricato: dipendenti.db`);
});