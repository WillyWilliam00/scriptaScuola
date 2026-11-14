# Shared: regole locali

- `validation.ts` e `types.ts` sono usati sia dal backend sia dal frontend: una modifica qui va riportata su entrambi i lati **nella stessa PR** (rotte, hook, form, `lib/api.ts`).
- Se la modifica serve a più lavori in corso, falla in una PR piccola e separata, da unire per prima.
- Schemi di input stretti (`.strict()` o `z.strictObject`): i campi extra sono un errore.
- Messaggi di validazione in italiano: compaiono nei form.
- Ogni schema nuovo o cambiato ha i test in `validation.test.ts`; `npm test -w shared` non richiede database.
