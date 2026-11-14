# Convenzioni

## Lingua
Tutto in italiano: testi dell'interfaccia, messaggi di errore delle API, commenti, nomi di dominio (`docenti`, `registrazioniCopie`, `istitutoId`), commit, PR e issue. I nomi tecnici standard restano in inglese (`router`, `handler`, `hook`).

## Backend
- **Multi-istituto.** Ogni rotta protetta è montata in `index.ts` con `authMiddleware` e `tenantStoreMiddleware`. Le rotte usano `req.tenantStore.<area>.<metodo>()`; ogni metodo di `tenantStore.ts` filtra per `istitutoId`, anche in conteggi, sottoquery ed export. Un id che arriva dal client (docente, utente, registrazione) va verificato come appartenente all'istituto.
- **Permessi.** `/api/utenti` e `/api/istituti` sono solo per admin (in `index.ts`). In `docenti.ts` e `registrazioniCopie.ts` creazione, modifica, eliminazione, export e liste delle registrazioni hanno `requireRole('admin')` sulla singola rotta: una rotta nuova segue le vicine.
- **Rotte.** Handler dentro `asyncHandler(...)`; input validato con lo schema di `shared` (`schema.parse(req.body)` o `req.query`): un `ZodError` diventa 400 nell'`errorHandler`.
- **Errori.** Lancia un `Error` con `status` (`ErrorWithStatus`) o rispondi `res.status(n).json({ error: '...' })`. Messaggi in italiano, senza dettagli tecnici.
- **Risposte.** Mai `passwordHash` o token: seleziona colonne esplicite e non passare a `res.json` una riga intera di `utenti`.
- **Transazioni.** Controlli di limite e scritture che dipendono da una lettura stanno nella stessa transazione. Tipa `tx` (non aggiungere altri `tx: any`).
- **Schema.** Una tabella nuova va anche in `cleanTestDb` (`tests/db/test-setup.ts`). Ogni modifica a `schema.ts` va segnalata nella PR, dicendo se è distruttiva e cosa va applicato al database di produzione.

## Shared
- Gli schemi di `validation.ts` servono sia al backend sia ai form: cambiarli tocca entrambi i lati.
- Gli schemi di modifica e creazione rifiutano i campi extra (`.strict()` o `z.strictObject`, issue #35): fai lo stesso negli schemi nuovi.
- Ogni schema nuovo o cambiato ha i suoi test in `validation.test.ts`.

## Frontend
- **Dati dal server** solo con TanStack Query, tramite gli hook in `src/hooks/`. Dopo una mutazione si invalidano le query giuste; la chiave contiene tutti i filtri e la pagina.
- **Form** con TanStack Form e gli schemi di `shared`. Per le utenze: `useUtenteForm` e `UtenteFormFields` in `components/gestione-utenze/`. Niente validatori `onMount` nei form nuovi (vedi le trappole in `linee-guida.md`).
- **HTTP** solo con `api` di `lib/api.ts` (base `/api`, rinnovo della sessione automatico). Mai `fetch` o Axios diretti verso il backend.
- **Sessione.** L'access token resta in memoria; in `localStorage` solo utente e istituto (vedi `store/auth-store.ts`).
- **Componenti** nella cartella della propria area; componenti di base da `components/ui/` (shadcn).
- **React.** Regole degli hook rispettate senza `eslint-disable`; `useEffect` solo per sincronizzarsi con l'esterno; `key` stabili; bottoni con sola icona con `aria-label`.

## Commit
Formato `tipo(area): descrizione`, in italiano, al presente:
- `feat(auth): reset password dei collaboratori con password temporanea`
- `fix(validation): createUtenteSchema rifiuta i campi extra`
- tipi usati: `feat`, `fix`, `refactor`, `test`, `docs`, `chore`, `style`, `ci`.
