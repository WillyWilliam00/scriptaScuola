# Flusso git e issue

## Branch
```
branch di lavoro ──PR──▶ develop ──PR di rilascio──▶ main
   (CI verde)          (integrazione)              (produzione)
```
- `develop` è il branch predefinito del repo: ogni lavoro parte da lì e torna lì con una PR.
- `main` è la produzione (Vercel e Render pubblicano da `main`, da confermare in #37). Non si fa mai push diretto su `main` né su `develop`.
- Nome del branch: libero e descrittivo (`fix/limite-modifica-registrazione`); le sessioni di Claude Code usano il branch che viene loro assegnato.

## Pull request
- **Una issue, una PR.** Nel corpo `Closes #N`: con `develop` predefinito l'issue si chiude al merge in `develop`.
- Base sempre `develop` (salvo rilasci e hotfix, sotto).
- Compila il template (`.github/pull_request_template.md`): cosa cambia, come è stato verificato e cosa no, eventuali passi manuali (schema del database, variabili su Render o Vercel).
- Si unisce solo con la CI verde (tre job) e dopo aver letto la review automatica. I punti 🔴 si correggono o si spiegano in risposta; la review ricontrolla da sola dopo il push.
- PR piccole: una modifica a `shared/validation.ts` che serve a più lavori va in una PR separata, unita per prima.
- Strategia di merge (da impostare in #37): **squash** per le PR verso `develop`, **merge commit** per i rilasci verso `main` (con lo squash i due branch divergono e i conflitti si ripetono). Branch cancellato dopo il merge.

## Rilascio (`develop` → `main`)
1. `develop` è verde e provato.
2. PR da `develop` a `main` con l'elenco delle PR incluse e i passi manuali (schema, variabili d'ambiente).
3. Prima del merge: applicare al database di produzione le modifiche di `schema.ts`, se ci sono, e impostare le variabili nuove su Render e Vercel.
4. Merge con **merge commit**. Dopo il deploy: `GET /api/health` e una prova di login.

## Hotfix
Branch da `main`, PR verso `main`, poi una PR (o un merge) da `main` verso `develop`, così la correzione non si perde al rilascio successivo.

## Issue
- Il piano generale è l'issue madre **#20**: ogni lavoro è una sua sotto-issue, da spuntare quando è chiusa.
- Prima di iniziare leggi la sezione **⛔ Dipendenze** dell'issue e la tabella in #20: se una dipendenza è aperta, non partire.
- Struttura di una buona issue (template in `.github/ISSUE_TEMPLATE/`): **Problema**, **Cosa fare**, **⛔ Dipendenze**, **Fatto quando**, più una riga su cosa è stato verificato (letto nel codice, riprodotto, non verificato).
- Se un'issue è datata rispetto al codice, aggiornala o scrivilo nella PR, invece di seguire il testo alla lettera.

## Lavoro in parallelo
Più sessioni possono lavorare insieme se non toccano gli stessi file:
- **backend**: `route/`, `db/tenantStore.ts`, `db/schema.ts` (lo schema è il punto più delicato: una sola PR alla volta);
- **frontend**: una cartella di `components/` per area funzionale, con i suoi hook;
- **shared**: `validation.ts` e `types.ts` toccano entrambi i lati: PR piccola, separata, fatta prima;
- `package-lock.json`: conflitti frequenti ma meccanici; si risolvono rigenerandolo con `npm install` dopo aver preso `develop`.

Prima di aprire una PR, aggiorna il branch con `develop` (`git merge origin/develop`) e rilancia i controlli.
