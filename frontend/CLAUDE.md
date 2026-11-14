# Frontend: regole locali

Vedi anche `../docs/agents/convenzioni.md` (sezione Frontend) e `../docs/agents/linee-guida.md`.

- Chiamate al backend solo con `api` di `src/lib/api.ts` (base `/api`); dati dal server solo con TanStack Query, tramite gli hook di `src/hooks/`.
- Dopo una mutazione invalida le query giuste; le chiavi includono filtri e pagina.
- Form con TanStack Form e gli schemi di `@shared/validation`. Utenze: `useUtenteForm` e `UtenteFormFields`. Niente validatori `onMount` nei form nuovi.
- Access token solo in memoria (`src/store/auth-store.ts`); in `localStorage` niente token.
- Componenti nella cartella della propria area in `src/components/`; `src/components/ui/` è generato da shadcn.
- Testi dell'interfaccia in italiano; bottoni con sola icona con `aria-label`; `Link` per i collegamenti.
- Prima di chiudere: `npm run lint -w frontend` e `npm run build -w frontend`, e prova la modifica nel browser (Playwright, `npm run dev` su 3000 con il backend su 3001).
- `vercel.json`: CSP e rewrite. La rewrite `/api/(.*)` resta prima del fallback su `index.html`.
- Per l'aspetto grafico c'è la skill `.claude/skills/frontend-design`.
