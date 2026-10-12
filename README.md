# ServerESP32

Server Node.js per la **rilevazione delle presenze tramite tessere NFC**.

Un dispositivo ESP32 con lettore NFC legge l'UID di una tessera e lo invia al server con una richiesta HTTP POST. Il server identifica il dipendente tramite un database SQLite e registra la timbratura nel foglio Excel del mese corrente, salvato su un NAS. Calcola inoltre le ore lavorate e gli straordinari.

## Funzionamento

1. L'ESP32 legge una tessera e invia l'UID a `POST /api/nfc`.
2. Il server cerca l'UID nel database SQLite `dipendenti.db`.
3. Se l'UID non è registrato, la timbratura viene salvata come `Sconosciuto (<uid>)`.
4. La timbratura viene scritta nel file Excel del mese corrente sul NAS.
5. Il server risponde all'ESP32 con l'esito dell'operazione.

### Timbrature giornaliere

Ogni dipendente può timbrare fino a quattro volte al giorno:

| Passaggio | Significato |
| --- | --- |
| 1° | Entrata |
| 2° | Uscita per la pausa |
| 3° | Rientro dalla pausa |
| 4° | Uscita |

Dal quinto passaggio il server non registra altre timbrature e restituisce `limit_reached`.

### Calcolo delle ore

- Con quattro passaggi: `(uscita - entrata) - (rientro - uscita per la pausa)`.
- Con solo entrata e uscita per la pausa: `secondo passaggio - entrata`.
- Gli straordinari sono le ore lavorate oltre le otto ore giornaliere.
- La riga `TOTALI MENSILI` somma le ore lavorate e gli straordinari del mese ed è ricalcolata a ogni timbratura.

### File Excel

I file sono creati nella cartella configurata, con questa struttura:

```text
<NAS_BASE_PATH>/<anno>/Presenze_<Mese>_<anno>.xlsx
```

Esempio: `2026/Presenze_Settembre_2026.xlsx`. Il foglio si chiama `Presenze` e contiene: Dipendente, UID Tessera, Data, quattro passaggi, Totale Ore Lavorate e Ore Straordinario.

## Tecnologie

- Node.js 18 o superiore
- Express 5
- better-sqlite3
- ExcelJS
- fs-extra
- Docker Compose per il deployment containerizzato

## Avvio in sviluppo senza Docker

```bash
git clone https://github.com/P1zz4M4n/ServerESP32.git
cd ServerESP32
npm install
npm start
```

Il server ascolta su `0.0.0.0` alla porta configurata. Per l'avvio senza Docker, verificare in `server.js` i valori effettivamente letti dal processo e predisporre il database e il percorso NAS accessibile in scrittura.

## Deployment con Docker Compose

Il template pubblico è `docker-compose.example.yml`. La configurazione effettiva di ciascun host va tenuta in un `docker-compose.yml` locale e ignorato da Git, insieme al file `.env`.

### Configurare un'installazione

Dalla cartella del progetto:

```bash
cp docker-compose.example.yml docker-compose.yml
cp .env.example .env
```

Modificare `.env` con i percorsi dell'host. Esempio:

```dotenv
HOST_PORT=3000
TZ=Europe/Rome
DB_FILE=/percorso/locale/dati/dipendenti.db
NAS_MOUNT=/mnt/NAS
NAS_BASE_PATH="/NAS/Cartella dipendenti/tabella presenze/Lettura NFC"
```

`DB_FILE` deve indicare un database esistente e inizializzato. `NAS_MOUNT` deve essere una directory esistente sull'host. Per usare un NAS, montarlo prima sul sistema operativo; Compose condivide la directory montata, ma non configura la connessione CIFS. `NAS_BASE_PATH` è il percorso visto dal container e deve corrispondere alla cartella effettivamente presente nel NAS.

Il template monta il database in `/data/dipendenti.db` e il NAS in `/NAS`. Il database è configurato in sola lettura; il NAS è accessibile in scrittura per creare e aggiornare i file Excel.

Verificare e avviare il servizio:

```bash
docker compose config -q
docker compose up -d --build
docker compose ps
```

Per controllare i log:

```bash
docker compose logs --tail 100 server
```

### Nota importante: database in sola lettura

Il Compose di produzione fornito per questo deployment monta `dipendenti.db` con `read_only: true`. Questo protegge il file dalle scritture del container, ma è incompatibile con `POST /api/dipendenti` se l'endpoint deve inserire o aggiornare dipendenti nel database. In tal caso occorre rimuovere `read_only: true` dal volume e assicurarsi che l'utente nel container possa scrivere nel file e nella directory SQLite. Se il database deve restare immutabile durante il funzionamento, mantenere la modalità di sola lettura e aggiornare il database con una procedura amministrativa separata.

Verificare quale comportamento si vuole prima di distribuire una nuova configurazione. Il template può essere adattato di conseguenza.

## NAS su Linux

Il mount CIFS va configurato sull'host, non nel container. Un esempio generico di riga `/etc/fstab` è:

```fstab
//<NAS_HOST>/<SHARE> /mnt/NAS cifs credentials=/root/.nas-credentials,vers=3.0,uid=<UID>,gid=<GID>,file_mode=0664,dir_mode=0775,_netdev,nofail,x-systemd.automount,x-systemd.mount-timeout=30s 0 0
```

Sostituire i segnaposto con i valori dell'ambiente. Il file `/root/.nas-credentials` deve restare fuori dal repository e accessibile solo a root, ad esempio con permessi `600`. Il mount deve essere verificato prima di avviare l'applicazione: una directory `/mnt/NAS` esistente non garantisce da sola che il filesystem di rete sia montato.

Per un server che deve avviare Docker solo quando il NAS è disponibile, configurare sul sistema host una dipendenza systemd dal mount e una verifica della sorgente CIFS. La configurazione systemd e le credenziali sono specifiche dell'host e non vanno incluse nel repository pubblico. Dopo aver corretto o spostato una cartella sul NAS, aggiornare `NAS_BASE_PATH` se il percorso interno è cambiato.

Comandi utili sul server:

```bash
findmnt /mnt/NAS
systemctl status mnt-NAS.mount mnt-NAS.automount --no-pager
systemctl status docker.service --no-pager
docker compose ps
```

## Separare il template pubblico dalla configurazione del server

Nel `.gitignore` del repository mantenere almeno:

```gitignore
# Configurazioni specifiche dell'host
/docker-compose.yml
/.env

# Dati locali
*.db
*.sqlite
*.sqlite3
```

`docker-compose.example.yml` e `.env.example` restano versionati e visibili su GitHub. I file reali `docker-compose.yml` e `.env` restano sull'host. Se `docker-compose.yml` è già tracciato, aggiungerlo al `.gitignore` non basta: va rimosso dall'indice Git in una migrazione, conservando prima una copia della configurazione usata dal server.

Per applicare aggiornamenti dell'applicazione su un host già configurato:

```bash
git pull --ff-only
docker compose config -q
docker compose up -d --build
docker compose ps
```

Il pull aggiorna il codice e il template pubblico, non ricrea da solo il container. Il Compose locale ignorato e `.env` non vengono sovrascritti. Controllare `git status` prima di aggiornare e conservare una copia del database prima di interventi che lo riguardano.

## API

### `POST /api/nfc`

Registra una timbratura. Richiesta:

```json
{ "uid": "04A1B2C3" }
```

| Codice | Risposta | Significato |
| --- | --- | --- |
| 200 | `{"status":"success","dipendente":"Mario Rossi","receivedUid":"04A1B2C3"}` | Timbratura registrata |
| 200 | `{"status":"limit_reached","dipendente":"...","receivedUid":"...","message":"..."}` | Quattro timbrature già eseguite oggi |
| 400 | `{"status":"error","message":"UID mancante"}` | Campo `uid` assente |
| 500 | Errore del server | Errore durante la registrazione, ad esempio NAS non raggiungibile |

Esempio:

```bash
curl -X POST http://localhost:3000/api/nfc \
  -H "Content-Type: application/json" \
  -d '{"uid":"04A1B2C3"}'
```

### `POST /api/dipendenti`

Associa un UID a un nome o aggiorna il nome già associato:

```json
{ "uid": "04A1B2C3", "nome": "Mario Rossi" }
```

Restituisce `200` in caso di successo o `400` se manca uno dei campi richiesti. Per il deployment Docker, consultare la nota sulla modalità di sola lettura del database.

### `GET /api/dipendenti`

Restituisce l'elenco dei dipendenti registrati come array JSON di oggetti `{ "uid": "...", "nome": "..." }`.

## Sicurezza e dati

- Gli endpoint non sono protetti da autenticazione. Limitare l'accesso alla rete fidata e non pubblicare la porta su Internet senza aggiungere autenticazione e protezioni adeguate.
- Non inserire password, credenziali CIFS, `.env` o database reali nel repository.
- Conservare copie di backup del database e dei file Excel presenti sul NAS.
- Proteggere l'accesso a database, fogli presenze e API: contengono dati personali e di presenza.

## Struttura essenziale

```text
ServerESP32/
├── server.js
├── package.json
├── package-lock.json
├── Dockerfile
├── docker-compose.example.yml
├── .env.example
└── README.md
```

Il database operativo e i file Excel sono dati runtime: non fanno parte del template pubblico.

## Licenza

ISC
