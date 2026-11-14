# Todo – ScriptaScuola

Lista di attività da completare per il progetto. Usa questo file come roadmap.

---






## Architettura frontend



- [x] **Gestire ErrorBoundary localmente nelle tabelle**  
  Implementare ErrorBoundary specifici per ogni tabella (DataTable) per gestire errori di caricamento dati in modo isolato, senza bloccare l'intera applicazione.

---

## CRUD


- [x] **Controllare CRUD in tabelle**  
  Verificare e testare tutte le operazioni CRUD nelle tabelle esistenti (docenti, utenti) per assicurarsi che funzionino correttamente e gestiscano gli errori in modo appropriato.
  - [x] *Inserire filtri di ricerca in gestione utente e registrazioni*  
  - [x] *impaginazione lato backend*  






- [x] **Export tabelle**  
  Esportazione in CSV (apribile in Excel) delle tabelle docenti, utenti e registrazioni, con i filtri applicati. L'import dei docenti da Excel usa `exceljs`.



## UI/UX

- [x] **Inserire landing page**  
  Creare una pagina di benvenuto/landing page per l'applicazione con informazioni sul sistema e accesso rapido alle funzionalità principali.
  - [x] *rendere respnsive tabella + sidebar*  



---

## Autenticazione e sicurezza

- [ ] **Inserire recupero password per collaboratori**  
  Implementare funzionalità di recupero password per i collaboratori:
  - Form di richiesta reset password (email)
  - Invio email con link di reset
  - Pagina di reset password con token
  - Validazione token e aggiornamento password

- [ ] **Recupero password per admin**  
  Implementare funzionalità di recupero password per gli amministratori:
  - Form di richiesta reset password (email)
  - Invio email con link di reset
  - Pagina di reset password con token
  - Validazione token e aggiornamento password





---

*Ultimo aggiornamento: novembre 2026*
