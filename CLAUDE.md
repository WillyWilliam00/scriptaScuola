# ScriptaScuola: guida per Claude Code

Gestionale web per le scuole: conta le fotocopie dei docenti rispetto a un limite, con utenti `admin` e `collaboratore` e dati separati per istituto. Monorepo npm con tre workspace: `backend` (Express + Drizzle + PostgreSQL), `frontend` (React + Vite), `shared` (schemi Zod e tipi usati da entrambi).

Scrivi **in italiano**: interfaccia, messaggi di errore, commenti, commit, PR e issue.

## Regole essenziali

1. **Mai toccare un database reale.** I test del backend svuotano le tabelle (`TRUNCATE ... CASCADE`) e, se manca `DATABASE_URL_TEST`, usano `DATABASE_URL`. Prima di `npm test` nel backend controlla che punti a un database di test.
2. **Isolamento tra istituti.** L'`istitutoId` arriva solo dal token (`req.user`). I dati di un istituto si leggono e si scrivono solo tramite `req.tenantStore` (`backend/src/db/tenantStore.ts`), mai con `db` direttamente nelle rotte.
3. **Validazioni in `shared/validation.ts`**, usate sia dal backend sia dai form. Se cambi lì, aggiorna backend e frontend nella stessa PR.
4. **Branch da `develop`, PR verso `develop`**, una issue per PR, `Closes #N` nel corpo. Mai push su `main`. Dettagli in `docs/agents/flusso-git-e-issue.md`.
5. **Niente scorciatoie per far passare la CI**: niente test saltati o indeboliti, niente `any`, `as any`, `@ts-ignore` o `eslint-disable` nuovi.
6. **Mai segreti nel repo** (`.env` è ignorato) e mai token o password nei log o nelle risposte JSON.
7. **Di' sempre cosa hai verificato e cosa no**: letto nel codice, eseguito in locale, provato nel browser. Le modifiche all'interfaccia si provano nel browser (Playwright), non solo con i test.

## Comandi rapidi (dalla radice)

```bash
npm ci                       # installa tutti i workspace
npm run dev -w backend       # API su http://localhost:3001
npm run dev -w frontend      # sito su http://localhost:3000 (proxy /api verso 3001)
npm run lint -w frontend     # ESLint
npm run build -w frontend    # tsc -b + vite build
npm test -w shared           # test degli schemi, nessun database
npx tsc --noEmit -p backend  # tipi del backend
npm test -w backend          # test di integrazione: SOLO con un database di test
```

«Fatto» vuol dire: lint, build, tipi e test verdi, come nella CI (`.github/workflows/ci.yml`).

## Dove trovare il resto

Leggi il file che serve prima di lavorare in quell'area:

- `docs/agents/struttura.md`: mappa del monorepo e file chiave.
- `docs/agents/stack.md`: tecnologie, versioni, deploy.
- `docs/agents/comandi.md`: installazione, database locale, test, CI.
- `docs/agents/convenzioni.md`: come si scrive il codice qui.
- `docs/agents/flusso-git-e-issue.md`: branch, PR, issue, merge, rilasci.
- `docs/agents/linee-guida.md`: cosa non fare, sicurezza, definizione di «fatto», trappole note.
- `backend/CLAUDE.md`, `frontend/CLAUDE.md`, `shared/CLAUDE.md`: regole locali di ogni workspace.
- `.github/review-rules.md`: cosa controlla la review automatica delle PR.
- Piano dei lavori: issue #20 (con le dipendenze tra issue).
