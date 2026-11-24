# ServerESP32
 
Node.js server for **attendance tracking with NFC cards**.
 
An ESP32 device with an NFC reader reads a card's UID and sends it to the server with an HTTP POST request. The server identifies the employee through a SQLite database and records the punch in the Excel sheet for the current month, stored on a NAS. It also calculates worked hours and overtime.
 
## How it works
 
1. The ESP32 reads a card and sends the UID to `POST /api/nfc`.
2. The server looks the UID up in the SQLite database `dipendenti.db`.
3. If the UID is not registered, the punch is still saved, as `Sconosciuto (<uid>)` ("Unknown").
4. The punch is written to the Excel file of the current month on the NAS.
5. The server answers the ESP32 with the outcome of the operation.
The punch time is the moment the request **arrives** at the server, not the moment the file is written. Requests are processed one at a time (in arrival order) so the Excel file is never read and written concurrently.
 
### Daily punches
 
Each employee can punch up to four times a day:
 
| Punch | Meaning |
| --- | --- |
| 1st | Clock in |
| 2nd | Clock out for the break |
| 3rd | Back from the break |
| 4th | Clock out |
 
From the fifth punch on, the server records nothing more and returns `limit_reached`.
 
**Duplicate-read protection.** A card left on the reader, or read twice in a row, would otherwise consume several punches within seconds and close the day with zero hours. For this reason, if the same employee punches again less than `MIN_INTERVALLO_TIMBRATURE_SEC` seconds (default **60**) after their last punch of the day, the new punch is ignored: nothing is written and the server returns `duplicate_ignored`. Set the variable to `0` to disable the check.
 
### Hours calculation
 
Daily hours are recomputed from the raw punches on every request:
 
| Punches in the day | Worked hours |
| --- | --- |
| 1 or 3 | `00:00:00` (the day is not closed yet) |
| 2 | `2nd − 1st` |
| 4 | `(2nd − 1st) + (4th − 3rd)` |
 
- Overtime is the worked time beyond eight hours per day.
- The `TOTALI MENSILI` ("Monthly totals") row sums the worked hours and overtime of **all rows** of the month and is rebuilt on every punch.
- Hours are stored as `HH:MM:SS` text; totals can exceed 24 hours (for example `164:30:00`).
### Excel files
 
Files are created in the configured folder with this structure:
 
```text
<NAS_BASE_PATH>/<year>/Presenze_<Month>_<year>.xlsx
```
 
Example: `2026/Presenze_Settembre_2026.xlsx` (month names are in Italian). The sheet is called `Presenze` and contains: Dipendente (employee), UID Tessera (card UID), Data (date, `DD/MM/YYYY`), the four punches, Totale Ore Lavorate (worked hours) and Ore Straordinario (overtime).
 
**Safe saving.** The file is never overwritten in place. The server writes a temporary file in the same folder (`<file>.xlsx.<pid>.tmp`) and then renames it over the real one, so a crash, a container restart or a NAS disconnection in the middle of a write cannot leave a truncated workbook. After a hard crash a leftover `.tmp` file may remain; it is safe to delete. If the workbook is open in Excel on another computer, the replacement may be refused by the NAS: the request then fails with `500` and the punch is not saved. Open a **copy** of the file when you need to consult it.
 
## Technologies
 
- Node.js 18 or later
- Express 5
- better-sqlite3
- ExcelJS
- fs-extra
- Docker Compose for containerized deployment
## Configuration
 
The server is configured with environment variables.
 
| Variable | Required | Default | Description |
| --- | --- | --- | --- |
| `DB_PATH` | yes | – | Absolute path of the SQLite database (the file must already exist). |
| `NAS_BASE_PATH` | yes | – | Absolute path of the folder where the Excel files are stored. Must be readable **and writable**. |
| `PORT` | no | `3000` | Listening port. |
| `TZ` | no | `Europe/Rome` | Time zone used for dates and times of the punches. |
| `API_KEY` | no | *(empty)* | If set, every `/api/*` route requires the header `x-api-key: <API_KEY>`. See [Security and data](#security-and-data). |
| `MIN_INTERVALLO_TIMBRATURE_SEC` | no | `60` | Minimum seconds between two punches of the same employee. `0` disables the check. |
 
### Database schema
 
The `dipendenti` table must contain the columns `uid` and `nome`, and `uid` must be a `PRIMARY KEY` or have a `UNIQUE` constraint (it is required by the upsert used to register employees). Minimal example:
 
```sql
CREATE TABLE dipendenti (
    uid  TEXT PRIMARY KEY,
    nome TEXT NOT NULL
);
```
 
UIDs are trimmed and converted to uppercase (maximum 100 characters); names are trimmed (maximum 150 characters).
 
### Startup checks
 
At startup the server verifies the configuration and **refuses to start** (exit code 1, with an explanatory message) if:
 
- `DB_PATH` or `NAS_BASE_PATH` is missing or is not an absolute path;
- the database file does not exist, the `dipendenti` table is missing, it lacks `uid`/`nome`, or `uid` has no uniqueness constraint;
- the NAS folder is not reachable or is **not writable** (a read-only mount is rejected on purpose, since every punch would fail).
If `API_KEY` is not set, the server prints a warning that the `/api/*` routes are public.
 
## Running in development without Docker
 
```bash
git clone https://github.com/P1zz4M4n/ServerESP32.git
cd ServerESP32
npm install
 
export DB_PATH=/absolute/path/to/dipendenti.db
export NAS_BASE_PATH=/absolute/path/to/writable/folder
export TZ=Europe/Rome
npm start
```
 
The server listens on `0.0.0.0` on the configured port.
 
## Deployment with Docker Compose
 
The public template is `docker-compose.example.yml`. The actual configuration of each host must be kept in a local `docker-compose.yml` that is ignored by Git, together with the `.env` file.
 
### Configuring an installation
 
From the project folder:
 
```bash
cp docker-compose.example.yml docker-compose.yml
cp .env.example .env
```
 
Edit `.env` with the host paths. Example:
 
```dotenv
HOST_PORT=3000
TZ=Europe/Rome
DB_FILE=/local/path/data/dipendenti.db
NAS_MOUNT=/mnt/NAS
NAS_BASE_PATH="/NAS/Employees folder/attendance sheet/NFC reading"
 
# Optional
API_KEY=
MIN_INTERVALLO_TIMBRATURE_SEC=60
```
 
`DB_FILE` must point to an existing, initialized database. `NAS_MOUNT` must be an existing directory on the host. To use a NAS, mount it on the operating system first; Compose shares the mounted directory but does not configure the CIFS connection. `NAS_BASE_PATH` is the path as seen **by the container** and must match the folder that actually exists on the NAS.
 
The template mounts the database at `/data/dipendenti.db` and the NAS at `/NAS`. The database is configured read-only; the NAS is writable so that the Excel files can be created and updated.
 
The two optional variables must also be passed to the container. Add them to the `environment` section of your local `docker-compose.yml` (and to `.env.example` if you want them in the public template):
 
```yaml
    environment:
      API_KEY: ${API_KEY:-}
      MIN_INTERVALLO_TIMBRATURE_SEC: ${MIN_INTERVALLO_TIMBRATURE_SEC:-60}
```
 
**Graceful shutdown.** On `SIGTERM`/`SIGINT` (for example `docker compose stop` or a restart) the server stops accepting requests, waits for the Excel write in progress, closes the database and exits. If this takes more than 15 seconds it exits anyway. Docker's default grace period is 10 seconds, so add the following to the service to let the server finish by itself:
 
```yaml
    stop_grace_period: 20s
```
 
**Health check (optional).** `GET /health` can be used as a Docker health check:
 
```yaml
    healthcheck:
      test: ["CMD", "node", "-e", "fetch('http://localhost:'+(process.env.PORT||3000)+'/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"]
      interval: 30s
      timeout: 5s
      retries: 3
```
 
Check and start the service:
 
```bash
docker compose config -q
docker compose up -d --build
docker compose ps
```
 
To read the logs:
 
```bash
docker compose logs --tail 100 server
```
 
### Important note: read-only database
 
The production Compose file provided for this deployment mounts `dipendenti.db` with `read_only: true`. This protects the file from writes made by the container, but it is incompatible with `POST /api/dipendenti` if the endpoint must insert or update employees in the database: in that case the endpoint answers `500`. Punching (`POST /api/nfc`) and listing employees keep working, because they only read the database.
 
To allow writes, remove `read_only: true` from the volume and make sure the user inside the container can write to the file **and to its directory**. Since SQLite creates a journal file next to the database during writes, it is better to mount the **directory** that contains `dipendenti.db` rather than the single file; with a single-file bind mount the journal ends up outside the volume. If the database must stay immutable while the server runs, keep the read-only mode and update the database with a separate administrative procedure.
 
Decide which behavior you want before deploying a new configuration. The template can be adapted accordingly.
 
## NAS on Linux
 
The CIFS mount must be configured on the host, not in the container. A generic example of an `/etc/fstab` line is:
 
```fstab
//<NAS_HOST>/<SHARE> /mnt/NAS cifs credentials=/root/.nas-credentials,vers=3.0,uid=<UID>,gid=<GID>,file_mode=0664,dir_mode=0775,_netdev,nofail,x-systemd.automount,x-systemd.mount-timeout=30s 0 0
```
 
Replace the placeholders with the values of your environment. The file `/root/.nas-credentials` must stay out of the repository and be readable only by root, for example with `600` permissions. Verify the mount before starting the application: an existing `/mnt/NAS` directory does not by itself guarantee that the network filesystem is mounted.
 
For a server that must start Docker only when the NAS is available, configure on the host a systemd dependency on the mount and a check of the CIFS source. The systemd configuration and the credentials are host-specific and must not be included in the public repository. After fixing or moving a folder on the NAS, update `NAS_BASE_PATH` if the internal path has changed.
 
Useful commands on the server:
 
```bash
findmnt /mnt/NAS
systemctl status mnt-NAS.mount mnt-NAS.automount --no-pager
systemctl status docker.service --no-pager
docker compose ps
```
 
## Separating the public template from the server configuration
 
In the repository's `.gitignore`, keep at least:
 
```gitignore
# Host-specific configuration
/docker-compose.yml
/.env
 
# Local data
*.db
*.sqlite
*.sqlite3
```
 
`docker-compose.example.yml` and `.env.example` stay versioned and visible on GitHub. The real `docker-compose.yml` and `.env` files stay on the host. If `docker-compose.yml` is already tracked, adding it to `.gitignore` is not enough: it must be removed from the Git index in a migration, after saving a copy of the configuration used by the server.
 
To apply application updates on an already configured host:
 
```bash
git pull --ff-only
docker compose config -q
docker compose up -d --build
docker compose ps
```
 
The pull updates the code and the public template, but does not recreate the container by itself. The ignored local Compose file and `.env` are not overwritten. Check `git status` before updating, and keep a copy of the database before any operation that touches it.
 
## API
 
All `/api/*` routes require the `x-api-key` header when `API_KEY` is set. `/health` is always public.
 
Error responses always have the form `{"status":"error","message":"..."}`. The `message` strings are returned by the server in Italian, as shown below.
 
### `POST /api/nfc`
 
Records a punch. Request:
 
```json
{ "uid": "04A1B2C3" }
```
 
| Code | Response | Meaning |
| --- | --- | --- |
| 200 | `{"status":"success","dipendente":"Mario Rossi","receivedUid":"04A1B2C3"}` | Punch recorded. If the card is not registered, `dipendente` is `Sconosciuto (<uid>)`. |
| 200 | `{"status":"duplicate_ignored","dipendente":"...","receivedUid":"...","message":"Timbratura ravvicinata ignorata"}` | Same employee punched less than `MIN_INTERVALLO_TIMBRATURE_SEC` seconds ago; nothing was written. |
| 200 | `{"status":"limit_reached","dipendente":"...","receivedUid":"...","message":"Limite massimo di 4 timbrature giornaliere raggiunto"}` | Four punches already made today. |
| 400 | `{"status":"error","message":"UID mancante o non valido"}` | `uid` missing, not a string, empty or longer than 100 characters. |
| 400 | `{"status":"error","message":"Richiesta non valida"}` | Malformed JSON body (other 4xx client errors, such as a body larger than 100 KB, are reported the same way with their own status code). |
| 401 | `{"status":"error","message":"Non autorizzato"}` | `API_KEY` is set and the `x-api-key` header is missing or wrong. |
| 500 | `{"status":"error","message":"Errore interno del server"}` | Error while recording, for example NAS unreachable or workbook locked or corrupted. The punch is **not** saved. |
 
Example:
 
```bash
curl -X POST http://localhost:3000/api/nfc \
  -H "Content-Type: application/json" \
  -H "x-api-key: <API_KEY>" \
  -d '{"uid":"04A1B2C3"}'
```
 
(Omit the `x-api-key` header if `API_KEY` is not set.)
 
### `POST /api/dipendenti`
 
Associates a UID with a name, or updates the name already associated:
 
```json
{ "uid": "04A1B2C3", "nome": "Mario Rossi" }
```
 
| Code | Meaning |
| --- | --- |
| 200 | `{"status":"success","message":"UID 04A1B2C3 associato a Mario Rossi"}` |
| 400 | Missing `uid` or `nome` (`Parametri UID e Nome richiesti`), or name longer than 150 characters (`Nome troppo lungo`). |
| 401 | Missing or wrong API key. |
| 500 | Database error, for example the database is mounted read-only. |
 
Rows already written in past Excel files keep the name they had when they were recorded. For the Docker deployment, see the note on the database read-only mode.
 
### `GET /api/dipendenti`
 
Returns the list of registered employees as a JSON array of `{ "uid": "...", "nome": "..." }` objects, sorted by name. It answers `401` without a valid API key and `500` on database errors.
 
### `GET /health`
 
Checks that the NAS folder is reachable and writable and that the database responds. It does not require the API key and does not expose internal paths.
 
| Code | Response |
| --- | --- |
| 200 | `{"status":"ok","timezone":"Europe/Rome"}` |
| 503 | `{"status":"error","message":"NAS o database non disponibili"}` |
 
## ESP32 reader integration
 
The reader firmware is maintained separately from this repository. To work with this server it must:
 
- send `POST /api/nfc` with `Content-Type: application/json` and the body `{"uid":"<UPPERCASE HEX>"}`;
- send the `x-api-key` header if the server has `API_KEY` set;
- use an HTTP read timeout long enough for the server to read and write the workbook on the NAS (the current firmware uses 15 s, plus a 5 s connection timeout);
- handle all the `status` values and HTTP codes listed above, and not treat an unknown response as a success.
Reference behavior of the current firmware:
 
| Server response / condition | Reader signal |
| --- | --- |
| `success` | Blue LED for about 1 s, 1 short beep |
| `success` with `Sconosciuto (...)` | Blue and red LEDs alternating, 3 medium beeps (punch saved, card not registered) |
| `duplicate_ignored` | Blue LED, 2 very short ticks |
| `limit_reached` | Blue and red LEDs together, 2 medium beeps |
| `400`, invalid response, unknown status | Red LED, 3 fast beeps |
| `401` | Red LED, 1 long beep + 2 short beeps |
| `500` | Red LED, 4 long beeps |
| Connection to the server failed | Red LED, 2 long beeps (punch not received) |
| No complete response (timeout, connection lost) | Red LED, 3 long beeps (**outcome uncertain**: the punch may have been saved) |
| Wi-Fi not connected | Red LED on, 1 very long beep |
 
A card left on the reader is registered only once; it must be removed before the next punch.
 
## Troubleshooting
 
When the reader signals an error, the server log (`docker compose logs --tail 200 server`) tells where the problem is. Look at the time of the failed punch:
 
| In the server log | Meaning | Where to look |
| --- | --- | --- |
| No `TIMBRATURA RICEVUTA` line at that time | The request never reached the server | Wi-Fi, access points, network between reader and server |
| `TIMBRATURA RICEVUTA` followed by `❌ Errore non gestito` | Server-side error; the log shows the error code | NAS mount, Excel file locked or corrupted, permissions |
| `TIMBRATURA RICEVUTA` followed by `✅ Excel salvato correttamente`, but the reader showed an error | The punch **was** saved but the response was slow or lost | Time between the two log lines (NAS speed), network quality |
| `⚠️ Timbratura ravvicinata ignorata` | Duplicate read ignored (normal) | – |
 
You can also open the Excel file of the day: if the punch column is filled, the punch was saved. If an employee punched again after an error, the repeated punch appears minutes after the first one.
 
## Security and data
 
- Routes under `/api` can be protected with a shared key: set `API_KEY` on the server and send it in the `x-api-key` header from the reader. The key is sent over plain HTTP, so it only protects against unauthorized clients on a trusted network; it does not protect against eavesdropping. Without `API_KEY` the endpoints are **public**: limit access to the trusted network and do not publish the port on the Internet without adding adequate protection (TLS, reverse proxy, firewall).
- `/health` is public by design (for Docker health checks) and exposes no internal paths.
- Do not put passwords, CIFS credentials, `.env` files or real databases in the repository.
- Keep backup copies of the database and of the Excel files on the NAS.
- Protect access to the database, the attendance sheets and the API: they contain personal and attendance data.
## Known limitations
 
- The write queue works inside a single Node.js process: do not run more than one container on the same Excel files.
- The server answers only after the Excel file has been written, so a slow or hanging NAS delays the answer to the reader and every request queued behind it.
- A day is identified by calendar date: a shift crossing midnight is split in two rows.
- The `TOTALI MENSILI` row adds up the hours of all employees together, not per employee.
- Unknown cards are saved as `Sconosciuto (<uid>)`; registering the card later does not rename the rows already written.
## Essential structure
 
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
 
The operational database and the Excel files are runtime data: they are not part of the public template.
 
## License
 
ISC
 
