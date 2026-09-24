# ServerESP32

Server Node.js per la **rilevazione delle presenze tramite tessere NFC**.

Un dispositivo ESP32 con lettore NFC legge il codice (UID) di una tessera e lo invia al server con una richiesta HTTP POST. Il server riconosce il dipendente, registra la timbratura in un foglio Excel mensile salvato su un NAS e calcola automaticamente ore lavorate e straordinari.

## Come funziona

1. L'ESP32 legge una tessera e invia l'UID a `POST /api/nfc`.
2. Il server cerca l'UID nel database SQLite locale (`dipendenti.db`) per ricavare il nome del dipendente. Se l'UID non è registrato, la timbratura viene comunque salvata come `Sconosciuto (<uid>)`.
3. La timbratura viene scritta nel file Excel del mese corrente.
4. Il server risponde all'ESP32 con l'esito, così il dispositivo può segnalarlo all'utente.

### Le timbrature giornaliere

Ogni dipendente può timbrare **fino a 4 volte al giorno**, nell'ordine:

| Passaggio | Significato |
|-----------|-------------|
| 1° | Entrata |
| 2° | Uscita per la pausa |
| 3° | Rientro dalla pausa |
| 4° | Uscita |

Dal quinto passaggio in poi il server non registra nulla e risponde con lo stato `limit_reached`.

### Calcolo delle ore

- Con 4 passaggi: ore lavorate = (uscita − entrata) − (rientro − uscita per la pausa).
- Con solo entrata e 2° passaggio (senza rientro né uscita): ore lavorate = 2° passaggio − entrata.
- **Straordinario**: le ore lavorate oltre le 8 ore giornaliere.
- In fondo al foglio, la riga **TOTALI MENSILI** somma ore lavorate e straordinari del mese. Viene ricalcolata a ogni timbratura.

### File Excel

I file vengono creati automaticamente, con questa struttura:

```
<cartella NAS>/<anno>/Presenze_<Mese>_<anno>.xlsx
```

Esempio: `2026/Presenze_Settembre_2026.xlsx`. Il foglio si chiama `Presenze` e contiene le colonne: Dipendente, UID Tessera, Data, i quattro passaggi, Totale Ore Lavorate, Ore Straordinario.

## Tecnologie

- [Node.js](https://nodejs.org/) (versione 18 o superiore)
- [Express](https://expressjs.com/) 5, per il server HTTP
- [better-sqlite3](https://github.com/WiseLibs/better-sqlite3), per il database dei dipendenti
- [ExcelJS](https://github.com/exceljs/exceljs), per la scrittura dei fogli Excel
- [fs-extra](https://github.com/jprichardson/node-fs-extra), per la gestione di file e cartelle

## Installazione

```bash
git clone https://github.com/P1zz4M4n/ServerESP32.git
cd ServerESP32
npm install
```

## Configurazione

Le impostazioni si trovano all'inizio di `server.js`:

| Costante | Descrizione | Valore attuale |
|----------|-------------|----------------|
| `PORT` | Porta su cui ascolta il server | `3000` |
| `NAS_BASE_PATH` | Cartella del NAS dove salvare i file Excel | percorso di rete (UNC) |

Modifica `NAS_BASE_PATH` con il percorso della tua cartella. Deve essere raggiungibile dalla macchina su cui gira il server, con permessi di lettura e scrittura.

## Avvio

```bash
npm start
```

All'avvio il server ascolta su tutte le interfacce di rete (`0.0.0.0`) sulla porta configurata e crea il database `dipendenti.db` nella cartella di lavoro, se non esiste.

## API

### `POST /api/nfc`

Registra una timbratura. Chiamato dall'ESP32.

Richiesta:

```json
{ "uid": "04A1B2C3" }
```

Risposte:

| Codice | Corpo | Significato |
|--------|-------|-------------|
| 200 | `{"status": "success", "dipendente": "Mario Rossi", "receivedUid": "04A1B2C3"}` | Timbratura registrata |
| 200 | `{"status": "limit_reached", "dipendente": "...", "receivedUid": "...", "message": "..."}` | Già eseguite 4 timbrature oggi |
| 400 | `{"status": "error", "message": "UID mancante"}` | Campo `uid` assente |
| 500 | (errore del server) | Problema durante la scrittura, ad esempio NAS non raggiungibile |

Esempio:

```bash
curl -X POST http://localhost:3000/api/nfc \
  -H "Content-Type: application/json" \
  -d '{"uid": "04A1B2C3"}'
```

### `POST /api/dipendenti`

Associa un UID a un nome. Se l'UID esiste già, aggiorna il nome.

```json
{ "uid": "04A1B2C3", "nome": "Mario Rossi" }
```

Risponde con `200` in caso di successo o `400` se manca uno dei due campi.

### `GET /api/dipendenti`

Restituisce l'elenco dei dipendenti registrati, come array JSON di oggetti `{ "uid": "...", "nome": "..." }`.

## Note di sicurezza

Gli endpoint **non sono protetti da autenticazione**. Chiunque raggiunga il server sulla rete può registrare timbrature o leggere l'elenco dei dipendenti. Il progetto è pensato per essere usato solo in una rete locale fidata.

## Struttura del progetto

```
ServerESP32/
├── server.js       # Tutto il server: API, database, gestione Excel
├── package.json
└── dipendenti.db   # Creato al primo avvio
```

## Licenza

ISC
