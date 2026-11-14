# Stack

Versioni indicative: quelle esatte sono in `package-lock.json`.

## Frontend
- React 19, React Router 7, Vite 7, TypeScript 5.9.
- Tailwind CSS 4, shadcn/ui (Radix e Base UI), icone Hugeicons, framer-motion.
- TanStack Query 5 (dati dal server), TanStack Form 1 (form), TanStack Table 8 (tabelle).
- Zustand 5 (stato di autenticazione), Axios (HTTP), ExcelJS (export).
- ESLint 9 con `react-hooks`, `react-refresh` e `@tanstack/eslint-plugin-query`.
- Nessun test del frontend per ora (issue #28).

## Backend
- Node.js 22 (quello della CI), Express 5, TypeScript, `tsx` e `nodemon` in sviluppo.
- Drizzle ORM con driver `pg`; `drizzle-kit push` per lo schema.
- Zod 4 (tramite `shared`), JWT (`jsonwebtoken`), `bcrypt`, `helmet`, `cors`, `cookie-parser`, `express-rate-limit`.
- Test: Vitest 4 e Supertest.

## Database
- PostgreSQL: Neon in produzione, `docker/docker-compose.yml` in locale (Postgres 16 nella CI).

## Deploy
- Frontend su **Vercel** (`frontend/vercel.json`). Il browser chiama `/api` sul dominio del sito e Vercel inoltra a `https://scriptascuola.onrender.com`: cookie di prima parte e `connect-src 'self'`.
- Backend su **Render**. Dietro due proxy (Vercel e Render): `TRUST_PROXY=2` di default.
- Il branch di produzione di Vercel e Render va verificato (issue #37): deve essere `main`.

## Servizi esterni previsti
- Email transazionali con **Resend** (dominio `scriptascuola.it`), login con Google e Microsoft tramite **Arctic**: piano nell'issue #39.
