---
name: babysit
description: Come portare una PR di ScriptaScuola fino al merge, con CI verde e review automatica risolta. Da usare dopo aver aperto una PR verso develop o quando la CI o la review segnalano problemi.
---

# Portare una PR a verde

1. **Controlla la CI** della PR (`.github/workflows/ci.yml`): tre job, Frontend, Shared, Backend.
   - Job rosso: leggi il log del passo fallito, riproduci in locale lo stesso comando (vedi `docs/agents/comandi.md`), correggi, push.
   - Mai saltare o indebolire test, mai `any` o `eslint-disable` per far passare.
   - Test del backend non eseguibili in locale (niente Postgres): correggi leggendo il log della CI e dillo nella PR.
2. **Leggi la review automatica** (commento di `claude[bot]`, regole in `.github/review-rules.md`).
   - 🔴 Bloccante: correggi, oppure rispondi spiegando perché non è un problema, con il riferimento al codice.
   - 🟡 Da verificare: controlla e rispondi con cosa hai verificato.
   - 🔵 Suggerimento: facoltativo.
   - Dopo il push o la risposta la verifica riparte da sola (massimo 3 volte); `/review` in un commento la riavvia a mano.
3. **Conflitti con `develop`**: `git fetch origin develop && git merge origin/develop`. Per `package-lock.json` prendi la versione di `develop` e rilancia `npm install`. Mai force push su branch altrui.
4. **Prima di dire che è pronta**: CI verde, nessun 🔴 aperto, «Fatto quando» dell'issue soddisfatto, sezione «Come è stato verificato» del template compilata con cosa è stato provato e cosa no.
5. Il merge lo fa il proprietario del repo, non l'agente.
