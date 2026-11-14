# Struttura del repo

Monorepo npm (`package.json` alla radice con `workspaces: shared, frontend, backend`). Un solo `package-lock.json`, alla radice.

```
backend/      API REST (Express 5, TypeScript, Drizzle, PostgreSQL)
frontend/     sito (React 19, Vite, Tailwind 4, TanStack)
shared/       schemi Zod e tipi comuni a backend e frontend
docker/       docker-compose con Postgres e pgAdmin per lo sviluppo locale
docs/agents/  questa documentazione
.github/      CI, review automatica delle PR, template di issue e PR
.claude/      skill di Claude Code del progetto
TODO.md       vecchia roadmap: il piano attuale è nell'issue #20
```

## Backend (`backend/src/`)

| Percorso | Cosa contiene |
| --- | --- |
| `index.ts` | crea l'app Express: helmet, CORS, `trust proxy`, montaggio delle rotte e dei middleware, `errorHandler`. Esporta `app` per i test. |
| `route/` | una rotta per area: `authentication.ts` (`/api/auth`), `docenti.ts`, `utenti.ts`, `registrazioniCopie.ts`, `istituti.ts`, `health.ts`. |
| `middleware/auth.ts` | `authMiddleware`, `requireRole`, `asyncHandler`, `ErrorWithStatus`, `errorHandler`. |
| `middleware/tenantStore.ts` | mette in `req.tenantStore` lo store dell'istituto del token. |
| `middleware/rateLimit.ts` | limiti ai tentativi (disattivati con `NODE_ENV=test`) e `getTrustProxy()`. |
| `middleware/refreshCookie.ts` | cookie httpOnly del refresh token. |
| `db/schema.ts` | schema Drizzle: `istituti`, `docenti`, `utenti`, `registrazioni_copie`, `refresh_tokens`, `richieste_reset_password`. |
| `db/tenantStore.ts` | **unico accesso ai dati di un istituto**: ogni query filtra per `istitutoId`. |
| `db/index.ts` | connessione `db` (driver `pg`, `drizzle-orm/node-postgres`). |
| `db/utils/` | funzioni pure: limiti, paginazione, ordinamento, hash dei refresh token, password temporanee. |
| `scripts/seed-istituto-prova.ts` | dati di prova (`npm run seed -w backend`). |

Test in `backend/tests/`: `routes/*.integration.test.ts` (con database, tramite Supertest su `app`), `db/utils/` e `middleware/` (unitari). Preparazione del database di test in `tests/db/test-setup.ts` e `tests/db/test-seed.ts`, login di comodo in `tests/utils/test-helpers.ts`.

Non ci sono migrazioni: lo schema si applica con `drizzle-kit push` (`backend/drizzle.config.ts`).

## Frontend (`frontend/src/`)

| Percorso | Cosa contiene |
| --- | --- |
| `App.tsx` | rotte (React Router) e `QueryClient`. |
| `components/<area>/` | componenti per funzione: `auth`, `RegistraCopie`, `gestione-docenti`, `gestione-utenze`, `visualizza-registrazioni`, `profilo`, `landing`, `layout`, `table`, `common`. |
| `components/ui/` | componenti generati da shadcn: non modificarli a mano senza motivo. |
| `hooks/` | hook di TanStack Query per area (`use-docenti.ts`, `use-utenti.ts`, `use-registrazioni.ts`, ...). |
| `lib/api.ts` | istanza Axios su `/api`, rinnovo automatico della sessione con coda delle richieste. |
| `lib/auth-api.ts` | chiamate di login, logout, refresh, cambio password. |
| `store/auth-store.ts` | Zustand: utente e istituto in `localStorage`, access token **solo in memoria**. |
| `vercel.json` (in `frontend/`) | CSP, header di sicurezza, rewrite `/api/*` verso il backend su Render, poi fallback su `index.html`. |

Alias: `@` → `frontend/src`, `@shared` → `shared` (in `vite.config.ts` e `tsconfig.app.json`).

## Shared (`shared/`)

- `validation.ts`: schemi Zod (input delle rotte, query di paginazione, form).
- `types.ts`: tipi comuni (`JwtPayload`, `Utente`, codici errore come `CAMBIO_PASSWORD_OBBLIGATORIO`).
- `validation.test.ts`: test degli schemi.

Il backend importa con percorsi relativi (`../../../shared/validation.js`), il frontend con `@shared/validation`.
