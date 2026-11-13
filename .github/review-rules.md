# Regole per la review automatica delle PR

Queste regole guidano il reviewer automatico (`.github/workflows/claude-review.yml`). Valgono anche per chi fa una review a mano.

## Obiettivo

Trovare **regressioni e bug reali** introdotti dalla PR: qualcosa che prima funzionava e ora no, dati di un istituto visibili a un altro, un controllo di sicurezza che salta. Non è una review di stile.

Segnala un problema solo se sai indicare **lo scenario concreto in cui si rompe** (quale richiesta, quale utente, quale dato). Se è solo un sospetto, mettilo in «Da verificare» e di' cosa controllare. Meglio tre segnalazioni fondate che dieci generiche.

## Contesto del progetto

- Monorepo: `backend` (Express 5, Drizzle, PostgreSQL su Neon, JWT), `frontend` (React, Vite, TanStack Query, Zustand), `shared` (schemi Zod e tipi usati da entrambi).
- Multi-istituto: ogni dato appartiene a un istituto. L'accesso ai dati passa da `backend/src/db/tenantStore.ts`, creato con l'`istitutoId` del token.
- Ruoli: `admin` e `collaboratore`.
- Deploy: frontend su Vercel (`frontend/vercel.json`), backend su Render.
- Il piano degli interventi è nell'issue #20; l'autenticazione nelle issue #39 e #41.

## Cosa controllare

### 1. Isolamento tra istituti (massima priorità)
- Ogni `select`, `update` e `delete` su `docenti`, `utenti`, `registrazioni_copie` filtra per `istitutoId`. Vale anche per sottoquery, conteggi e query di export.
- `istitutoId` arriva sempre da `req.user` (token), mai da body, query o parametri.
- Le rotte nuove sono montate in `backend/src/index.ts` con `authMiddleware` e `tenantStoreMiddleware`.
- Le rotte non usano `db` direttamente per i dati di un istituto: passano da `req.tenantStore`.
- Un id ricevuto dal client (docente, utente, registrazione) viene sempre verificato come appartenente all'istituto prima di usarlo.

### 2. Permessi
- Le rotte che modificano, eliminano o esportano dati hanno `requireRole('admin')` quando serve. Confronta con le rotte vicine dello stesso file.
- Una rotta non diventa pubblica per errore (ordine dei middleware, router montato sotto `/api/auth`).

### 3. Limiti di copie
- Ogni modifica a somme, limiti o registrazioni: il totale è calcolato giusto (in modifica va esclusa la registrazione stessa), il confronto è tra numeri, tutto avviene nella stessa transazione.
- Attenzione a `tx: any`: nasconde errori di tipo, per esempio usare la riga `{ totale }` come se fosse un numero.
- Due richieste simultanee non devono poter superare il limite.

### 4. Autenticazione e sessioni
- `passwordHash` e refresh token non escono mai dalle risposte né dai log. Sospetta di `.returning()` senza colonne e di `select()` su `utenti` passato a `res.json`.
- I refresh token sono salvati solo come hash, ruotati a ogni uso e revocati al cambio password.
- Le rotte pubbliche di autenticazione hanno il limite ai tentativi.
- Cookie di sessione: `httpOnly`, `Secure`, `SameSite`, `Path` corretti; controllo dell'`Origin` su refresh e logout.
- I messaggi di errore del login non rivelano se un utente esiste.
- Le scelte restano coerenti con il piano in #39 e #41.

### 5. Contratto tra backend e frontend
- Se cambiano `shared/validation.ts` o `shared/types.ts`, sono aggiornati sia il backend sia il frontend (hook in `frontend/src/hooks`, `lib/api.ts`, `lib/auth-api.ts`, form).
- Se cambia la forma di una risposta o il percorso di una rotta, il frontend che la consuma è aggiornato nella stessa PR.

### 6. Schema del database
- Nel repo non ci sono migrazioni: ogni modifica a `backend/src/db/schema.ts` va segnalata sempre, dicendo se è distruttiva (colonna rimossa o rinominata, `NOT NULL` senza default, nuovo vincolo `unique` su dati esistenti) e cosa va applicato al database prima del deploy.
- Una tabella nuova va aggiunta anche a `cleanTestDb` in `backend/tests/db/test-setup.ts`.
- `onDelete` coerente con le altre tabelle.

### 7. Configurazione e deploy
- `frontend/vercel.json`: CSP (`connect-src` verso il backend), ordine delle rewrite.
- CORS e `trust proxy` in `backend/src/index.ts`.
- Una variabile d'ambiente nuova è documentata e ha un comportamento sensato se manca. Ricorda nel commento che va impostata su Render o Vercel.
- Dipendenze nuove: servono davvero e stanno nel workspace giusto.

### 8. Test
- Logica cambiata senza un test che la copre: indica quale test manca.
- Test rimossi, saltati (`skip`, `only`) o con asserzioni indebolite.
- I test di integrazione condividono un database: un test nuovo pulisce e semina i dati come gli altri.

### 9. Frontend e regole di React
- Dopo una mutazione le query giuste di TanStack Query vengono invalidate; le chiavi delle query includono tutti i filtri e la pagina.
- L'interceptor di refresh in `lib/api.ts` non entra in ciclo e non perde le richieste in coda.
- `ProtectedRoute` e `PublicRoute` continuano a proteggere le pagine; l'interfaccia non mostra a un collaboratore azioni da admin.
- La validazione dei form usa gli schemi di `shared`.
- Regole degli hook: niente hook dentro condizioni, cicli o dopo un `return`; dipendenze di `useEffect`, `useMemo` e `useCallback` complete, senza disattivare la regola con un commento.
- `useEffect` usato solo per sincronizzarsi con l'esterno: non per calcolare dati derivati dallo stato o dalle props, non per copiare i dati di una query in uno stato locale, non per reagire a un evento che si può gestire nel suo handler.
- Effetti con timer, listener o richieste hanno la funzione di pulizia.
- Stato non mutato direttamente; nessuno stato duplicato che può andare fuori sincronia (per esempio una lista filtrata salvata nello stato).
- Liste con `key` stabile (l'id del dato, non l'indice) quando gli elementi si possono aggiungere, togliere o riordinare.
- Niente `dangerouslySetInnerHTML`, `eval`, `innerHTML`; niente token o dati sensibili in `localStorage` oltre a quanto previsto da #41.
- Niente `as any` o `@ts-ignore` nuovi per far tacere un errore di tipo: segnala cosa nascondono.
- Il lint del frontend (con le regole di React e di TanStack Query) gira già nel workflow `lint-frontend.yml` sui file cambiati: controlla comunque questi punti, ma segnala soprattutto i commenti `eslint-disable` nuovi e i casi che il lint non può vedere.
- Bottoni con solo un'icona hanno `aria-label`; la navigazione usa `Link` e non `onClick` con `navigate()` quando è un semplice collegamento.

### 10. Aderenza all'issue
Le issue del progetto descrivono cosa fare e quando il lavoro è finito («Cosa fare», «Fatto quando»). Se la PR ne cita una:
- **Manca**: una richiesta dell'issue non è fatta o è fatta a metà. Cita la riga dell'issue.
- **Diverso**: è fatta in un modo diverso da quello deciso nell'issue (per esempio una libreria o un flusso scartati).
- **In più**: la PR cambia cose che l'issue non chiede. Segnalalo solo se il cambiamento è rischioso o dovrebbe stare in una PR separata.
Se la PR non cita nessuna issue, salta questa sezione senza segnalarlo.

## Cosa ignorare
- Stile, formattazione, nomi, ordine degli import.
- I componenti generati in `frontend/src/components/ui/`.
- `package-lock.json`, salvo una dipendenza inattesa.
- Testi della landing page e file di documentazione.
- Problemi che esistevano già e che la PR non tocca. Se uno è grave, una riga in fondo con «Preesistente».
- Non eseguire test, build o installazioni.

## Formato del commento

Un solo commento, in italiano, senza preamboli:

```
## Review automatica

**Esito:** nessun problema trovato | N punti da guardare (X bloccanti)

### 🔴 Bloccante
- **`percorso/file.ts:riga` Titolo breve.** Cosa si rompe e in quale scenario. Come sistemarlo.

### 🟡 Da verificare
- ...

### 🔵 Suggerimento
- ...

_Controllato: elenco breve delle aree guardate. Non controllato: cosa è rimasto fuori e perché._
```

- **Bloccante**: bug certo, perdita o esposizione di dati, sicurezza, regressione di una funzione esistente.
- **Da verificare**: rischio plausibile che non puoi confermare dal codice.
- **Suggerimento**: miglioramento utile ma non necessario. Al massimo due.
- Al massimo sei voci in tutto. Ometti le sezioni vuote.
- Se non trovi nulla: solo le righe «Esito» e «Controllato».
- Commenti sulla riga solo per i bloccanti, con lo stesso testo.
