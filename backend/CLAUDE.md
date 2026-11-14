# Backend: regole locali

Vedi anche `../docs/agents/convenzioni.md` (sezione Backend) e `../docs/agents/linee-guida.md`.

- ⚠️ `npm test` svuota il database di `DATABASE_URL_TEST`, o di `DATABASE_URL` se manca. Mai contro il database reale.
- Dati di un istituto solo tramite `req.tenantStore` (`src/db/tenantStore.ts`); ogni query filtra per `istitutoId`, che viene da `req.user`.
- Rotte: `asyncHandler`, input validato con gli schemi di `../shared/validation.ts`, errori in italiano. Montaggio e middleware in `src/index.ts`.
- `requireRole('admin')` sulle rotte che modificano, eliminano o esportano: guarda le rotte vicine.
- Mai `passwordHash` o token nelle risposte: colonne esplicite in `select` e `returning`.
- Controlli sui limiti di copie nella stessa transazione della scrittura; niente nuovi `tx: any`.
- Modifiche a `src/db/schema.ts`: nessuna migrazione nel repo, si usa `drizzle-kit push`. Aggiorna `cleanTestDb` in `tests/db/test-setup.ts` per le tabelle nuove e scrivi nella PR cosa va applicato in produzione.
- Test: integrazione in `tests/routes/*.integration.test.ts` con Supertest su `app`; ogni file pulisce e semina il database come gli altri (`cleanTestDb`, `test-seed.ts`). I file girano uno alla volta.
- Variabili d'ambiente nuove: documentale in `README.md`, dai un default sensato e ricorda nella PR di impostarle su Render.
