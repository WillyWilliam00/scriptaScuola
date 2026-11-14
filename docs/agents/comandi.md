# Comandi

Tutti i comandi partono dalla radice del repo, salvo dove indicato.

## Installazione
```bash
npm ci            # installa i tre workspace dal package-lock.json
```
Per aggiungere una dipendenza: `npm install <pacchetto> -w backend` (o `-w frontend`, `-w shared`). Il lock è unico, alla radice.

Con npm 10.9 `npm audit fix` va in crash: usare `npx npm@11 audit fix`, mai `--force` (porta a downgrade di versioni major).

## Database locale
```bash
cd docker && docker compose up -d   # Postgres su 5432, pgAdmin su 5050
```
Le credenziali vengono da un `.env` in `docker/` (`POSTGRES_USER`, `POSTGRES_PASSWORD`, `POSTGRES_DB`, `PGADMIN_DEFAULT_EMAIL`, `PGADMIN_DEFAULT_PASSWORD`). Crea **due** database: uno per sviluppare e uno solo per i test.

Variabili del backend (`backend/.env`, mai committato): l'elenco completo è in `backend/README.md`. Le principali: `DATABASE_URL`, `DATABASE_URL_TEST`, `JWT_SECRET`, `FRONTEND_ORIGIN`, `TRUST_PROXY`, `COOKIE_SECURE`.

Schema: `npx drizzle-kit push` dentro `backend/` applica `src/db/schema.ts` al database di `DATABASE_URL`. ⚠️ Controlla prima a quale database punta.

## Sviluppo
```bash
npm run dev -w backend     # http://localhost:3001 (nodemon + tsx)
npm run dev -w frontend    # http://localhost:3000, /api inoltrato a localhost:3001
npm run seed -w backend    # istituto di prova
```

## Controlli (gli stessi della CI)
```bash
npm run lint -w frontend
npm run build -w frontend          # tsc -b && vite build
npm test -w shared                 # nessun database
cd backend && npx tsc --noEmit     # tipi del backend
cd backend && npx vitest run --no-file-parallelism
```

### ⚠️ Test del backend
`backend/tests/db/test-setup.ts` usa `DATABASE_URL_TEST` **oppure, se manca, `DATABASE_URL`** e svuota tutte le tabelle con `TRUNCATE ... RESTART IDENTITY CASCADE`. Prima di lanciarli:
- `DATABASE_URL_TEST` deve esistere e puntare a un database usa e getta;
- lo schema deve essere applicato a quel database (`DATABASE_URL=<url di test> npx drizzle-kit push`);
- in una sessione cloud senza Postgres, non lanciarli: dillo e lascia che li esegua la CI.

I file di test girano uno alla volta (`fileParallelism: false` in `vitest.config.ts`) perché condividono il database.

## CI
`.github/workflows/ci.yml` gira su ogni PR verso `develop` o `main` e su ogni push a quei branch. Tre job: **Frontend** (lint e build), **Shared** (test), **Backend** (tipi, `drizzle-kit push` su un Postgres temporaneo, test). Una PR si unisce solo con i tre job verdi.

La review automatica (`.github/workflows/claude-review.yml`) commenta le PR verso `develop`; le regole sono in `.github/review-rules.md`. `/review` in un commento la riavvia a mano, `/review completa` rifà la review intera.
