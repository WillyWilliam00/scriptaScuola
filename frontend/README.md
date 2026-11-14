## ScriptaScuola Frontend

Il frontend di **ScriptaScuola** è una **single page application** in **React + TypeScript** che fornisce l’interfaccia per:
- registrare le fotocopie effettuate,
- gestire i docenti e i relativi limiti,
- gestire le utenze (admin/collaboratori),
- eseguire login e registrazione iniziale dell’istituto.

È sviluppato con **Vite** e utilizza **Tailwind CSS**, componenti UI (shadcn, Base UI) e **TanStack Query** per la gestione dei dati lato client.

### Indice

- [Stack frontend](#stack-frontend)
- [Struttura principale](#struttura-principale)
- [Integrazione con il backend ScriptaScuola](#integrazione-con-il-backend-scriptascuola)
- [Setup & avvio](#setup--avvio)
  - [Prerequisiti](#prerequisiti)
  - [Installazione dipendenze](#installazione-dipendenze)
  - [Avvio in sviluppo](#avvio-in-sviluppo)
- [Build di produzione](#build-di-produzione)
- [Note di design](#note-di-design)

### Stack frontend

- React 19 + TypeScript
- Vite
- React Router (`react-router-dom`) per il routing client‑side
- Tailwind CSS, shadcn/UI, Base UI per il design dei componenti
- TanStack Query / Form / Table per:
  - gestione asincrona dei dati (`QueryClient`),
  - validazione e gestione dei form (via Zod e adapter TanStack),
  - tabelle e viste tabellari.
- Integrazione con le API backend di ScriptaScuola tramite HTTP client (es. Axios).

### Struttura principale

- `src/main.tsx`
  - Monta l’app React nel DOM con `StrictMode`.
  - Importa gli stili globali (`index.css`).
  - Renderizza `App`.

- `src/App.tsx`
  - Configura:
    - `QueryClientProvider` di TanStack Query.
    - `BrowserRouter` e le varie route principali dell’app.
  - Routing:
    - `/` → **landing page** (`LandingPage`).
    - `/login` → pagina di login.
    - `/register` → pagina di registrazione/setup iniziale.
    - Rotte protette racchiuse in `AppLayout`:
      - `/registra-copie` → vista registrazione copie (`RegistraCopie`).
      - `/gestione-docenti` → gestione docenti e limiti (`GestioneDocenti`).
      - `/gestione-utenze` → gestione utenze (admin/collaboratori) (`GestioneUtenze`).
      - `/visualizza-registrazioni` → visualizzazione registrazioni copie (`VisualizzaRegistrazioni`).
      - `/profilo` → profilo utente (`ProfiloUtente`).
    - Route catch-all `*` → pagina 404 (`NotFound`).

- Componenti principali (indicativi):
  - `components/AppLayout.tsx`: layout principale con navigazione, header, ecc.
  - `components/LandingPage.tsx`, `components/landing/*`: landing page con Hero, Features, FAQ, CTA.
  - `components/RegistraCopie.tsx`: form e flusso per registrare nuove fotocopie (selezione docente, numero copie, note, ecc.).
  - `components/GestioneDocenti.tsx`: lista docenti con modali (NuovoDocenteModal, EditDocenteModal, ViewDocenteModal) e dialoghi (EliminaDocenteDialog, EliminaTuttiDocentiDialog, ResetConteggioDialog).
  - `components/ImportDocentiDialog.tsx`: dialog per importare docenti da file Excel con anteprima dei dati.
  - `components/GestioneUtenze.tsx`: lista utenti con modali (NuovoUtenteModal, EditUtenteModal, ViewUtenteModal) e dialoghi (EliminaUtenteDialog, DeleteUtenteAlertDialog).
  - `components/VisualizzaRegistrazioni.tsx`: visualizzazione registrazioni copie con tabella e filtri.
  - `components/ProfiloUtente.tsx`: profilo utente autenticato.
  - `components/NotFound.tsx`: pagina 404 per route non trovate.
  - `components/Login.tsx` e `components/Register.tsx`: flussi di autenticazione e setup istituto.

> Nota: i nomi esatti e la struttura interna dei componenti possono evolvere, ma l’idea principale è mantenere una separazione chiara tra aree funzionali (copie, docenti, utenze, auth).

### Integrazione con il backend ScriptaScuola

- Il frontend comunica con le API esposte dal backend:
  - `/api/auth` per login, setup scuola, refresh token, logout.
  - `/api/docenti` per elenco/gestione docenti, importazione bulk da Excel, eliminazione multipla.
  - `/api/utenti` per elenco/gestione utenti (admin).
  - `/api/registrazioni-copie` per creare/listare/eliminare registrazioni di copie, eliminazione multipla.
  - `/api/istituti` per gestione istituto (es. eliminazione).
- Funzionalità aggiuntive implementate:
  - Importazione bulk docenti da Excel (`POST /api/docenti/bulk-import`) con anteprima dati.
  - Eliminazione di tutti i docenti (`DELETE /api/docenti/delete-all`).
  - Eliminazione di tutte le registrazioni copie (`DELETE /api/registrazioni-copie/delete-all`).
  - **Export CSV (apribile in Excel) delle tabelle principali**, rispettando i filtri attivi e ignorando la paginazione:
    - Docenti: `GET /api/docenti/export`, usato dall’hook `useExportDocenti()` (`hooks/use-docenti.ts`) e dal componente `GestioneDocenti`.
    - Utenti: `GET /api/utenti/export`, usato dall’hook `useExportUtenti()` (`hooks/use-utenti.ts`) e dal componente `GestioneUtenze`.
    - Registrazioni copie: `GET /api/registrazioni-copie/export`, usato dall’hook `useExportRegistrazioni()` (`hooks/use-registrazioni.ts`) e dal componente `VisualizzaRegistrazioni`.
    - Gli hook lato frontend costruiscono la query string a partire dallo stato dei filtri, chiamano l’endpoint `/export` con `responseType: 'blob'` e in caso di successo generano e scaricano automaticamente un file `.csv` con nome predefinito (es. `docenti-YYYY-MM-DD.csv`), mostrando toast di conferma o errore tramite `sonner`.
- Le chiamate HTTP sono incapsulate nell’istanza Axios `api` (`src/lib/api.ts`), che gestisce automaticamente:
  - base URL configurabile,
  - header `Authorization` con token JWT,
  - refresh token e gestione errori 401.
- TanStack Query si occupa di:
  - caching delle risposte,
  - refetch automatico quando necessario,
  - gestione di loading/error state,
  - gestione delle mutation (create/update/delete/export) con invalidazione selettiva delle query correlate.

### Setup & avvio

#### Prerequisiti

- Node.js (versione LTS recente)
- Backend di ScriptaScuola in esecuzione (vedi `backend/README.md`)

#### Installazione dipendenze

Se hai già eseguito `npm install` nella root del monorepo, le dipendenze del frontend saranno già installate.
In alternativa, puoi installarle esplicitamente:

```bash
cd frontend
npm install
```

#### Avvio in sviluppo

```bash
cd frontend
npm run dev
```

Per impostazione predefinita Vite avvierà l’app su `http://localhost:5173` (o porta configurata).

### Build di produzione

Per creare una build ottimizzata:

```bash
cd frontend
npm run build
```

Puoi poi usare `npm run preview` per un’anteprima locale della build.

### Test

```bash
cd frontend
npm test               # tutti i test, una volta
npm run test:watch     # riesegue i test a ogni modifica
npm run test:coverage  # report di copertura in coverage/ (per ora senza soglie)
```

- **Vitest** con **jsdom** (browser simulato) e **Testing Library**; la configurazione è in `vitest.config.ts` e riusa gli alias di Vite.
- La rete è simulata con **MSW**: ogni test registra le risposte che gli servono con `server.use(...)`; una chiamata senza risposta prevista fa fallire il test. Non serve il backend.
- I test stanno accanto al file che provano (`api.ts` → `api.test.ts`); gli aiuti comuni sono in `src/test/` (`createWrapper`, `loginAs`, `makeJwt`, `stubDownload`).
- Regola: si prova il comportamento visibile (richieste inviate, stato della sessione, messaggi mostrati), non le funzioni interne.
- I test girano anche in CI, nel job del frontend.

### Note di design

- L’interfaccia è pensata per essere **semplice e veloce da usare in segreteria scolastica**:
  - accesso rapido alla registrazione copie,
  - navigazione chiara tra gestione docenti e utenze,
  - feedback immediato sugli errori (es. limiti di copie superati, permessi mancanti).
- L’uso combinato di Tailwind CSS, shadcn/UI e TanStack mira a:
  - mantenere un design moderno e coerente,
  - velocizzare lo sviluppo di componenti riutilizzabili,
  - ridurre il codice boilerplate nella gestione dello stato dei dati.

