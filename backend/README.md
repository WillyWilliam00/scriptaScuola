## ScriptaScuola Backend

Il backend di **ScriptaScuola** espone una serie di **API REST** multi‑tenant per la gestione di:
- autenticazione e setup iniziale della scuola,
- docenti e relativi limiti di copie,
- utenti dell’istituto (admin, collaboratori),
- registrazioni delle fotocopie effettuate.

È scritto in **TypeScript** usando **Express**, con **PostgreSQL** (es. Neon) come database e **Drizzle ORM** per l’accesso ai dati.

### Indice

- [Stack backend](#stack-backend)
- [Struttura principale](#struttura-principale)
- [Rotte principali](#rotte-principali)
  - [Autenticazione – `/api/auth`](#autenticazione--apiauth)
  - [Docenti – `/api/docenti`](#docenti--apidocenti)
  - [Utenti – `/api/utenti`](#utenti--apiutenti)
  - [Registrazioni copie – `/api/registrazioni-copie`](#registrazioni-copie--apiregistrazioni-copie)
  - [Istituti – `/api/istituti`](#istituti--apiistituti)
- [Modello dati (alto livello)](#modello-dati-alto-livello)
- [Setup & avvio](#setup--avvio)
  - [Variabili d’ambiente](#variabili-dambiente)
  - [Installazione dipendenze](#installazione-dipendenze)
  - [Avvio in sviluppo](#avvio-in-sviluppo)
- [Test](#test)
- [Note architetturali](#note-architetturali)

### Stack backend

- Node.js + TypeScript
- Express (API REST)
- PostgreSQL (`pg`) ospitato su Neon
- Drizzle ORM (`drizzle-orm`, `drizzle-kit`)
- Autenticazione JWT (`jsonwebtoken`) + refresh token
- Hashing password con `bcrypt`
- Validazione con Zod (schemi definiti in `shared/validation.ts`)
- Test di integrazione con Vitest + Supertest

### Struttura principale

- `src/index.ts`: entry point dell’app Express.
  - Registra i middleware globali (CORS, `express.json()`).
  - Espone le rotte:
    - `/api/auth` (pubbliche, autenticazione e setup).
    - `/api/docenti` (protette, admin/collaboratori).
    - `/api/utenti` (protette, solo admin).
    - `/api/registrazioni-copie` (protette, admin/collaboratori).
    - `/api/istituti` (protette, solo admin).
  - Applica i middleware `authMiddleware`, `tenantStoreMiddleware`, `requireRole` e `errorHandler`.
- `src/route/*.ts`: singole route (authentication, docenti, utenti, registrazioni copie).
- `src/db/tenantStore.ts`: contiene la logica di accesso ai dati per un singolo istituto (`istitutoId`), con metodi per docenti, utenti, registrazioni copie, ecc.
- `src/db/schema.ts`: definizioni delle tabelle Drizzle (istituti, utenti, docenti, registrazioni, token di refresh, ...).
- `src/middleware/auth.ts`: middleware per:
  - estrarre e verificare il JWT,
  - applicare `requireRole('admin')` o altri ruoli,
  - gestire gli errori applicativi.
- `tests/db/test-setup.ts`, `tests/db/test-seed.ts`: setup e seed del database di test.
- `tests/routes/*.integration.test.ts`, `tests/middleware/*.test.ts`, `tests/db/utils/*.test.ts`: test di integrazione delle rotte e test di unità sugli helper di backend.

### Rotte principali

#### Autenticazione – `/api/auth`

- `POST /api/auth/setup-scuola`
  - Crea un **nuovo istituto** e il relativo **utente admin** in un’unica operazione.
  - Body con due blocchi:
    - `istituto`: dati della scuola (`codiceIstituto`, nome, ecc.).
    - `utente`: dati dell’admin (email, password, ...).
  - Usa `insertIstitutoSchema` e `registerSchema` per validazione.
  - Crea lo `tenantStore` per l’istituto e registra l’admin con ruolo `admin`.

- `POST /api/auth/login`
  - Login con `identifier` (email o username) e `password`. Limite: 10 tentativi ogni 15 minuti per IP.
  - Verifica credenziali, genera:
    - access token JWT (15 minuti), restituito nel corpo e tenuto dal frontend solo in memoria,
    - refresh token (7 giorni), inviato **solo** in un cookie `httpOnly`, `SameSite=Strict`, `Path=/api/auth`. Nel database ne resta solo l'hash SHA-256.
  - Risposta: `token`, dati utente (id, email, username, ruolo), dati istituto.

- `POST /api/auth/logout`
  - Legge il refresh token dal cookie, lo revoca (`revokedAt`) e cancella il cookie. Risponde 200 anche senza cookie.

- `POST /api/auth/refresh`
  - Legge il refresh token dal cookie (il corpo della richiesta è ignorato).
  - Lo invalida, ne genera uno nuovo nel cookie e restituisce un **nuovo access token**.

- `POST /api/auth/change-password` (autenticata)
  - Revoca tutte le sessioni dell'utente e ne apre una nuova (nuovo cookie) per chi ha fatto la richiesta.

Il frontend chiama l'API su `/api` del proprio dominio: Vercel inoltra le richieste a Render (`frontend/vercel.json`), in sviluppo lo fa il proxy di Vite. Così il cookie è di prima parte e i browser non lo scartano.

#### Docenti – `/api/docenti`

Protette da `authMiddleware` + `tenantStoreMiddleware` (accesso per istituto).

- `GET /api/docenti`
  - Restituisce una **lista paginata** di docenti con filtri/ordinamenti.
  - Usa `docentiQuerySchema` per validare query string.
  - Risposta con `{ data, pagination }`.

- `GET /api/docenti/export`
  - Esporta **tutti i docenti che corrispondono ai filtri correnti** (ignorando paginazione) in formato **CSV** (apribile in Excel).
  - Usa gli stessi parametri di `GET /api/docenti` (filtri + sort), validati da `docentiQuerySchema`.
  - Restituisce un file con intestazioni leggibili (Nome, Cognome, Limite copie, Copie effettuate, Copie rimanenti, date creazione/aggiornamento).

- `POST /api/docenti/new-docente` (solo admin)
  - Crea un nuovo docente per l’istituto corrente.
  - Body validato con `insertDocenteSchema` (senza `istitutoId`, gestito da store).

- `PUT /api/docenti/update-docente/:id` (solo admin)
  - Aggiorna dati di un docente esistente (`nome`, `cognome`, `limiteCopie`, ecc.).

- `DELETE /api/docenti/delete-docente/:id` (solo admin)
  - Elimina un docente dell’istituto corrente.
  - Risposta con `id` del docente eliminato.

- `POST /api/docenti/bulk-import` (solo admin)
  - Importa multipli docenti da un file Excel in un'unica operazione.
  - Body validato con `bulkImportDocentiSchema` (array di docenti con `nome`, `cognome`, `limiteCopie` opzionale, `copieEffettuate` opzionale).
  - Risposta con messaggio di riepilogo: totale creati, totale con copie già effettuate.

- `DELETE /api/docenti/delete-all` (solo admin)
  - Elimina tutti i docenti dell'istituto corrente.
  - Risposta con messaggio e conteggio dei docenti eliminati.

#### Utenti – `/api/utenti`

Protette da `authMiddleware` + `tenantStoreMiddleware` + `requireRole('admin')`.

- `GET /api/utenti`
  - Restituisce lista paginata degli utenti dell’istituto.
  - Query validata con `utentiQuerySchema` (paginazione, filtri per ruolo/identifier, ordinamento) da `shared/validation.ts`.

- `GET /api/utenti/export`
  - Esporta **tutti gli utenti che corrispondono ai filtri correnti** (ignorando paginazione) in formato **CSV**.
  - Usa gli stessi parametri di `GET /api/utenti` (identifier, ruolo, sort), validati da `utentiQuerySchema`.
  - Nel CSV viene incluso un identificativo leggibile (email per admin, username per collaboratori) e le date di creazione/aggiornamento.

- `POST /api/utenti/new-utente`
  - Crea un nuovo utente (tipicamente collaboratore o admin).
  - Body validato con `createUtenteSchema` (ruolo, email/username, password).

- `PUT /api/utenti/update-utente/:id`
  - Aggiorna i dati di un utente (es. username, password).
  - `id` validato con `uuidParamSchema`, body validato con `modifyUtenteSchema`.

- `DELETE /api/utenti/delete-utente/:id`
  - Elimina un utente e restituisce l’`id` eliminato.
  - `id` validato con `uuidParamSchema`.

#### Registrazioni copie – `/api/registrazioni-copie`

Protette da `authMiddleware` + `tenantStoreMiddleware`.

- `GET /api/registrazioni-copie`
  - (Rotta tipicamente riservata ad admin) restituisce lista paginata di registrazioni con filtri.
  - Query validata con `registrazioniCopieQuerySchema` (paginazione, filtri per docenteId/utenteId, ordinamento).

- `GET /api/registrazioni-copie/export`
  - Esporta **tutte le registrazioni che corrispondono ai filtri correnti** (ignorando paginazione) in formato **CSV**.
  - Usa gli stessi filtri e parametri di sort di `GET /api/registrazioni-copie`, validati da `registrazioniCopieQuerySchema`.
  - Ogni riga contiene dati del docente, dell’utente (se presente), numero di copie, note e date di creazione/aggiornamento.

- `POST /api/registrazioni-copie/new-registrazione`
  - Crea una **nuova registrazione di copie** effettuate da un docente.
  - Body:
    - `docenteId`: docente a cui attribuire le copie.
    - `copieEffettuate`: numero di copie.
    - `utenteId`: utente (admin/collaboratore) che ha registrato l’operazione.
    - `note` opzionali.
  - Body validato con `insertRegistrazioneSchema`.
  - Lo store:
    - verifica che docente e utente appartengano all’istituto,
    - controlla i limiti di copie,
    - salva la registrazione.

- `DELETE /api/registrazioni-copie/delete-registrazione/:id`
  - Elimina una registrazione dell’istituto.
  - Risposta: `{ id }` della registrazione cancellata.
  - `id` validato con `idParamSchema`.

- `DELETE /api/registrazioni-copie/delete-all` (solo admin)
  - Elimina tutte le registrazioni di copie dell'istituto corrente.
  - Risposta con messaggio di conferma.

#### Istituti – `/api/istituti`

Protette da `authMiddleware` + `tenantStoreMiddleware` + `requireRole('admin')`.

- `DELETE /api/istituti/delete-istituto`
  - Elimina l’istituto dell’utente autenticato (istitutoId dal token, via tenantStore).
  - Risposta con conferma dell’eliminazione.

### Modello dati (alto livello)

In `tenantStore` e nello schema Drizzle troverai entità come:

- `istituti`: scuole/istituti configurati in ScriptaScuola.
- `utenti`: utenti dell’istituto (`admin`, `collaboratore`), con password hashata e ruolo.
- `docenti`: docenti associati a un istituto, con `limiteCopie` e contatori.
- `registrazioniCopie`: singole registrazioni di copie effettuate, legate a docente, utente e istituto.
- `refreshTokens`: token di refresh per la gestione sicura delle sessioni.

Ogni tabella è filtrata per `istitutoId` per garantire l’isolamento dei dati tra istituti.

### Setup & avvio

#### Variabili d’ambiente

Creare un file `.env` in `backend` (o configurare le variabili nel tuo ambiente) con almeno:

- `DATABASE_URL`: connection string PostgreSQL (Neon o altro).
- `JWT_SECRET`: chiave segreta per firmare i JWT. Obbligatoria: senza, il backend non parte.
- `PORT`: porta HTTP su cui esporre il backend (default 3000).
- `TRUST_PROXY`: numero di proxy davanti al backend (default 2: Vercel, che inoltra `/api`, più il proxy di Render). Serve al limite dei tentativi su login e registrazione per leggere l'IP reale del client. Dopo un deploy si verifica con `GET /api/health/ip`, che deve mostrare il proprio IP pubblico.
- `FRONTEND_ORIGIN`: origine del sito in produzione (es. `https://scriptascuola.it`). Le richieste da altre origini ricevono 403.
- `ALLOWED_ORIGINS`: altre origini ammesse, separate da virgola (es. `http://192.168.1.10:3000` per provare il sito da un telefono in rete locale). Facoltativa. Fuori da `NODE_ENV=production` è sempre ammesso anche `http://localhost:3000`.
- `COOKIE_SECURE`: `false` per disattivare l'attributo Secure del cookie di sessione in locale (serve solo con Safari su http). Di default il cookie è Secure fuori dai test.
- Eventuali altre variabili richieste dal tuo setup DB.

#### Installazione dipendenze

Nella root hai già eseguito `npm install`. Se vuoi lavorare solo sul backend:

```bash
cd backend
npm install
```

#### Avvio in sviluppo

```bash
cd backend
npm run dev
```

L’app ascolterà su `http://localhost:3000` (o la porta configurata).

### Test

Per eseguire i test (unitari/integrati) del backend:

```bash
cd backend
npm test
```

I test usano **Vitest + Supertest** e sono organizzati nella cartella `tests/`:
- `tests/routes/`: test di integrazione per le principali rotte API
- `tests/middleware/`: test per middleware come `auth`
- `tests/db/utils/`: test unitari per gli helper di accesso dati (paginazione, sort, limiti, ecc.)
- `tests/db/`: setup e seed del database di test (`test-setup.ts`, `test-seed.ts`)
- `tests/utils/`: helper condivisi per i test (es. `test-helpers.ts`)

Le rotte principali coperte dai test di integrazione sono:
- autenticazione (`/api/auth`)
- gestione docenti (`/api/docenti`)
- gestione utenti (`/api/utenti`)
- registrazioni copie (`/api/registrazioni-copie`)

**Configurazione dei test:**
- I test di integrazione vengono eseguiti in sequenza (`singleFork: true` in `vitest.config.ts`) per evitare conflitti quando più suite accedono al database condiviso
- Il database viene pulito con `TRUNCATE CASCADE` prima di ogni suite di test tramite `cleanTestDb()` in `tests/db/test-setup.ts`
- I dati di test vengono popolati tramite `seedTestData()` in `tests/db/test-seed.ts`

### Note architetturali

- Il backend è pensato per essere **multi‑tenant per istituto**: ogni richiesta autenticata porta con sé un `istitutoId` (dal JWT) e lo `tenantStore` opera solo sui dati di quell’istituto.
- La validazione centralizzata tramite Zod (in `shared`) mantiene **consistenza** tra ciò che accetta il backend e ciò che il frontend invia/si aspetta.

