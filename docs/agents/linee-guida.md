# Linee guida

## Cosa non fare mai
- Lanciare test, seed o `drizzle-kit push` contro un database reale (Neon di produzione). I test del backend cancellano tutto.
- Committare `.env`, chiavi, token o connection string. Nella CI si usano solo valori finti.
- Saltare, disattivare o indebolire test (`skip`, `only`, asserzioni tolte) per far passare la CI.
- Usare `any`, `as any`, `@ts-ignore` o `eslint-disable` per far tacere un errore. `tx: any` in `tenantStore.ts` ha già nascosto un bug (#21).
- Leggere o scrivere dati di un istituto con `db` direttamente in una rotta, o prendere `istitutoId` da body, query o parametri.
- Fare push su `main` o `develop`, riscrivere la storia di un branch altrui, unire una PR con la CI rossa.
- `npm audit fix --force` (porta a downgrade major).

## Sicurezza
- Rotte nuove dietro `authMiddleware` (e `tenantStoreMiddleware` se toccano dati di un istituto), con `requireRole` dove serve.
- Rotte pubbliche di autenticazione con il limite ai tentativi (`middleware/rateLimit.ts`).
- Mai loggare o restituire password, `passwordHash`, refresh token, token di reset.
- Messaggi di login che non rivelano se un utente esiste.
- Il refresh token vive solo nel cookie httpOnly; l'access token solo in memoria nel frontend.
- Un'origine nuova del frontend va aggiunta sia agli origin CORS (`backend/src/index.ts`) sia, se serve, al `connect-src` della CSP in `frontend/vercel.json`.

## Definizione di «fatto»
1. Lint, build del frontend, tipi del backend e test (shared e backend) verdi, in locale o nella CI.
2. Test nuovi per la logica nuova o corretta; per un bug, un test che fallisce senza la correzione.
3. Le modifiche all'interfaccia sono **provate nel browser** (per esempio con Playwright), non solo con test sugli schemi.
4. Il «Fatto quando» dell'issue è soddisfatto punto per punto, o la PR dice cosa manca e perché.
5. La PR dice **cosa è stato verificato e come** (letto nel codice, eseguito in locale, provato nel browser, solo nella CI) e cosa no.
6. Passi manuali scritti nella PR: schema da applicare, variabili da impostare su Render o Vercel.

## Trappole note
- **Validatore `onMount` di TanStack Form**: l'errore calcolato al mount resta in `canSubmit` e il bottone resta disabilitato anche con dati validi. Nei form nuovi usa `onChange`/`onSubmit`. Alcuni modali di docenti e registrazioni lo usano ancora: se li tocchi, provali nel browser.
- **Select di Radix in jsdom**: richiede polyfill (`hasPointerCapture`, `scrollIntoView`); spesso conviene provarlo con Playwright.
- **`.for('update')` di Drizzle** (lock di riga per i limiti, #22): verifica che funzioni con il driver `pg` in uso con un test di concorrenza vero.
- **`trust proxy`**: con il valore sbagliato il limite ai tentativi vede un solo IP. Dopo un deploy controlla `GET /api/health/ip`.
- **Ordine delle rewrite in `vercel.json`**: `/api/(.*)` deve restare prima del fallback su `index.html`.
- **Sessioni cloud senza Postgres**: i test del backend non si possono eseguire; dillo nella PR e affidati alla CI.
- **npm 10.9** va in crash su `audit fix`: usa `npx npm@11 audit fix`.
