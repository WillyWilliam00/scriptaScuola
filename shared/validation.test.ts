import { describe, it, expect } from 'vitest';
import { ZodError } from 'zod';
import {
  insertIstitutoSchema,
  loginSchema,
  registerSchema,
  docentiQuerySchema,
  createUtenteSchema,
  idParamSchema,
  uuidParamSchema,
  refreshTokenSchema,
  insertDocenteSchema,
  modifyUtenteSchema,
  richiestaResetSchema,
  modifyDocenteSchema,
  insertRegistrazioneSchema,
  registrazioniCopieQuerySchema,
  utentiQuerySchema,
  // ... altri che aggiungi
} from './validation.js';

describe('insertIstitutoSchema', () => {
  it('accetta input valido', () => {
    const input = { nome: 'Liceo Rossi', codiceIstituto: '1234567890' };
    expect(insertIstitutoSchema.parse(input)).toEqual(input);
  });

  it('rifiuta nome troppo corto', () => {
    const input = { nome: 'Ab', codiceIstituto: '1234567890' };
    expect(() => insertIstitutoSchema.parse(input)).toThrow(ZodError);
  });

  it('rifiuta codice istituto non di 10 caratteri', () => {
    const input = { nome: 'Liceo Rossi', codiceIstituto: '123' };
    expect(() => insertIstitutoSchema.parse(input)).toThrow(ZodError);
  });
});

describe('insertDocenteSchema', () => {
  it('accetta nome e cognome non vuoti', () => {
    const input = { nome: 'Mario', cognome: 'Rossi', limiteCopie: 10, istitutoId: 1 };
    expect(insertDocenteSchema.parse(input)).toEqual(input);
  });
  it('rifiuta nome troppo corto', () => {
    const input = { nome: 'A', cognome: 'Rossi', limiteCopie: 10, istitutoId: 1 };
    expect(() => insertDocenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta cognome troppo corto', () => {
    const input = { nome: 'Mario', cognome: 'R', limiteCopie: 10, istitutoId: 1 };
    expect(() => insertDocenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta limiteCopie minore di 0', () => {
    const input = { nome: 'Mario', cognome: 'Rossi', limiteCopie: -1, istitutoId: 1 };
    expect(() => insertDocenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta istitutoId non valido', () => {
    const input = { nome: 'Mario', cognome: 'Rossi', limiteCopie: 10, istitutoId: '1' };
    expect(() => insertDocenteSchema.parse(input)).toThrow(ZodError);
  })
})

describe('loginSchema', () => {
  it('accetta identifier e password non vuoti', () => {
    const input = { identifier: 'admin@scuola.it', password: 'password' };
    expect(loginSchema.parse(input)).toEqual(input);
  });

  it('rifiuta identifier vuoto', () => {
    expect(() => loginSchema.parse({ identifier: '', password: 'x' })).toThrow(ZodError);
  });

  it('rifiuta password vuota', () => {
    expect(() => loginSchema.parse({ identifier: 'x', password: '' })).toThrow(ZodError);
  });
  it('rifiuta password troppo corta', () => {
    const input = { identifier: 'admin@scuola.it', password: '1234567' };
    expect(() => loginSchema.parse(input)).toThrow(ZodError);
  })
});

describe('docentiQuerySchema', () => {
  it('applica default per page e pageSize', () => {
    const result = docentiQuerySchema.parse({});
    expect(result.page).toBe(1);
    expect(result.pageSize).toBe(20);
  });

  it('coerce stringhe numeriche da query', () => {
    const result = docentiQuerySchema.parse({ page: '2', pageSize: '10' });
    expect(result.page).toBe(2);
    expect(result.pageSize).toBe(10);
  });

  it('accetta sortField e sortOrder validi', () => {
    const result = docentiQuerySchema.parse({ sortField: 'cognome', sortOrder: 'desc' });
    expect(result.sortField).toBe('cognome');
    expect(result.sortOrder).toBe('desc');
  });
  it('rifiuta sortField non valido', () => {
    const input = { sortField: 'cognomeo', sortOrder: 'desc' };
    expect(() => docentiQuerySchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta sortOrder non valido', () => {
    const input = { sortField: 'cognome', sortOrder: 'desco' };
    expect(() => docentiQuerySchema.parse(input)).toThrow(ZodError);
  })
  it('corregge page minore di 1 usando catch', () => {
    const input = { page: 0, pageSize: 10, sortField: 'cognome', sortOrder: 'desc' };
    const result = docentiQuerySchema.parse(input);
    expect(result.page).toBe(1); // viene corretto a 1 invece di lanciare errore
  })
  it('corregge pageSize minore di 1 usando catch', () => {
    const input = { page: 1, pageSize: 0, sortField: 'cognome', sortOrder: 'desc' };
    const result = docentiQuerySchema.parse(input);
    expect(result.pageSize).toBe(20); // viene corretto a 20 invece di lanciare errore
  })
  it('corregge pageSize maggiore di 100 usando catch', () => {
    const input = { page: 1, pageSize: 101, sortField: 'cognome', sortOrder: 'desc' };
    const result = docentiQuerySchema.parse(input);
    expect(result.pageSize).toBe(20); // viene corretto a 20 invece di lanciare errore
  })
  it('accetta filtro nome', () => {
    const result = docentiQuerySchema.parse({ nome: 'Mario' });
    expect(result.nome).toBe('Mario');
  })
  it('accetta filtro cognome', () => {
    const result = docentiQuerySchema.parse({ cognome: 'Rossi' });
    expect(result.cognome).toBe('Rossi');
  })
  it('accetta filtri nome e cognome insieme', () => {
    const result = docentiQuerySchema.parse({ nome: 'Mario', cognome: 'Rossi' });
    expect(result.nome).toBe('Mario');
    expect(result.cognome).toBe('Rossi');
  })
});

describe('registrazioniCopieQuerySchema', () => {
  it('applica default per page e pageSize', () => {
    const result = registrazioniCopieQuerySchema.parse({});
    expect(result.page).toBe(1);
    expect(result.pageSize).toBe(20);
  });
  it('coerce stringhe numeriche da query', () => {
    const result = registrazioniCopieQuerySchema.parse({ page: '2', pageSize: '10' });
    expect(result.page).toBe(2);
    expect(result.pageSize).toBe(10);
  });
  it('accetta sortField e sortOrder validi', () => {
    const result = registrazioniCopieQuerySchema.parse({ sortField: 'docenteId', sortOrder: 'desc' });
    expect(result.sortField).toBe('docenteId');
    expect(result.sortOrder).toBe('desc');
  })
  it('rifiuta sortField non valido', () => {
    const input = { sortField: 'docenteIdo', sortOrder: 'desc' };
    expect(() => registrazioniCopieQuerySchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta sortOrder non valido', () => {
    const input = { sortField: 'docenteId', sortOrder: 'desco' };
    expect(() => registrazioniCopieQuerySchema.parse(input)).toThrow(ZodError);
  })
  it('corregge page minore di 1 usando catch', () => {
    const input = { page: 0, pageSize: 10, sortField: 'docenteId', sortOrder: 'desc' };
    const result = registrazioniCopieQuerySchema.parse(input);
    expect(result.page).toBe(1); // viene corretto a 1 invece di lanciare errore
  })
  it('corregge pageSize minore di 1 usando catch', () => {
    const input = { page: 1, pageSize: 0, sortField: 'docenteId', sortOrder: 'desc' };
    const result = registrazioniCopieQuerySchema.parse(input);
    expect(result.pageSize).toBe(20); // viene corretto a 20 invece di lanciare errore
  })
  it('corregge pageSize maggiore di 100 usando catch', () => {
    const input = { page: 1, pageSize: 101, sortField: 'docenteId', sortOrder: 'desc' };
    const result = registrazioniCopieQuerySchema.parse(input);
    expect(result.pageSize).toBe(20); // viene corretto a 20 invece di lanciare errore
  })
  it('accetta filtro docenteId', () => {
    const result = registrazioniCopieQuerySchema.parse({ docenteId: 5 });
    expect(result.docenteId).toBe(5);
  })
  it('accetta filtro utenteId', () => {
    const uuid = '123e4567-e89b-12d3-a456-426614174000';
    const result = registrazioniCopieQuerySchema.parse({ utenteId: uuid });
    expect(result.utenteId).toBe(uuid);
  })
  it('accetta filtri docenteId e utenteId insieme', () => {
    const uuid = '123e4567-e89b-12d3-a456-426614174000';
    const result = registrazioniCopieQuerySchema.parse({ docenteId: 5, utenteId: uuid });
    expect(result.docenteId).toBe(5);
    expect(result.utenteId).toBe(uuid);
  })
})
describe('utentiQuerySchema', () => {
  it('applica default per page e pageSize', () => {
    const result = utentiQuerySchema.parse({});
    expect(result.page).toBe(1);
    expect(result.pageSize).toBe(20);
  });
  it('coerce stringhe numeriche da query', () => {
    const result = utentiQuerySchema.parse({ page: '2', pageSize: '10' });
    expect(result.page).toBe(2);
    expect(result.pageSize).toBe(10);
  });
  it('accetta sortField e sortOrder validi', () => {
    const result = utentiQuerySchema.parse({ sortField: 'identificativo', sortOrder: 'desc' });
    expect(result.sortField).toBe('identificativo');
    expect(result.sortOrder).toBe('desc');
  })
  it('rifiuta sortField non valido', () => {
    const input = { sortField: 'username', sortOrder: 'desc' };
    expect(() => utentiQuerySchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta sortOrder non valido', () => {
    const input = { sortField: 'identificativo', sortOrder: 'desco' };
    expect(() => utentiQuerySchema.parse(input)).toThrow(ZodError);
  })
  it('corregge page minore di 1 usando catch', () => {
    const input = { page: 0, pageSize: 10, sortField: 'ruolo', sortOrder: 'desc' };
    const result = utentiQuerySchema.parse(input);
    expect(result.page).toBe(1); // viene corretto a 1 invece di lanciare errore
  })
  it('corregge pageSize minore di 1 usando catch', () => {
    const input = { page: 1, pageSize: 0, sortField: 'ruolo', sortOrder: 'desc' };
    const result = utentiQuerySchema.parse(input);
    expect(result.pageSize).toBe(20); // viene corretto a 20 invece di lanciare errore
  })
  it('corregge pageSize maggiore di 100 usando catch', () => {
    const input = { page: 1, pageSize: 101, sortField: 'ruolo', sortOrder: 'desc' };
    const result = utentiQuerySchema.parse(input);
    expect(result.pageSize).toBe(20); // viene corretto a 20 invece di lanciare errore
  })
  it('rifiuta identifier vuoto', () => {
    const input = { identifier: '', sortField: 'ruolo', sortOrder: 'desc' };
    expect(() => utentiQuerySchema.parse(input)).toThrow(ZodError);
  })
  it('accetta filtro ruolo admin', () => {
    const result = utentiQuerySchema.parse({ ruolo: 'admin' });
    expect(result.ruolo).toBe('admin');
  })
  it('accetta filtro ruolo collaboratore', () => {
    const result = utentiQuerySchema.parse({ ruolo: 'collaboratore' });
    expect(result.ruolo).toBe('collaboratore');
  })
})
describe('createUtenteSchema', () => {
  it('accetta utente admin con email', () => {
    const input = { ruolo: 'admin', email: 'a@b.it' };
    expect(createUtenteSchema.parse(input)).toEqual(input);
  });

  it('accetta collaboratore con username senza @', () => {
    const input = { ruolo: 'collaboratore', username: 'mariorossi' };
    expect(createUtenteSchema.parse(input)).toEqual(input);
  });

  it('rifiuta la password: la genera la piattaforma', () => {
    const input = { ruolo: 'collaboratore', username: 'mariorossi', password: 'password123' };
    expect(() => createUtenteSchema.parse(input)).toThrow(ZodError);
  });

  it('rifiuta collaboratore con username che contiene @', () => {
    const input = { ruolo: 'collaboratore', username: 'mario@rossi' };
    expect(() => createUtenteSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta email non valida per admin', () => {
    const input = { ruolo: 'admin', email: 'emailinvalida' };
    expect(() => createUtenteSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta username troppo corto per collaboratore', () => {
    const input = { ruolo: 'collaboratore', username: 'ab' };
    expect(() => createUtenteSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta ruolo non valido', () => {
    const input = { ruolo: 'superadmin', email: 'a@b.it' };
    expect(() => createUtenteSchema.parse(input)).toThrow(ZodError);
  });
});

describe('modifyUtenteSchema', () => {
  it('accetta utente admin con email', () => {
    const input = { ruolo: 'admin', email: 'a@b.it' };
    expect(modifyUtenteSchema.parse(input)).toEqual(input);
  });
  it('accetta collaboratore con username senza @', () => {
    const input = { ruolo: 'collaboratore', username: 'mariorossi' };
    expect(modifyUtenteSchema.parse(input)).toEqual(input);
  })
  it('rifiuta la password: si reimposta solo con la password temporanea', () => {
    const input = { ruolo: 'admin', email: 'a@b.it', password: 'password123' };
    expect(() => modifyUtenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta username troppo corto', () => {
    const input = { ruolo: 'collaboratore', username: 'm' };
    expect(() => modifyUtenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta ruolo non valido', () => {
    const input = { ruolo: 'admino', email: 'a@b.it' };
    expect(() => modifyUtenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta email non valida', () => {
    const input = { ruolo: 'admin', email: 'ab' };
    expect(() => modifyUtenteSchema.parse(input)).toThrow(ZodError);
  })
  // Issue #35
  it('rifiuta campi extra non permessi per admin (strict)', () => {
    const input = { ruolo: 'admin', email: 'a@b.it', campoExtra: 'x' };
    expect(() => modifyUtenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta campi extra non permessi per collaboratore (strict)', () => {
    const input = { ruolo: 'collaboratore', username: 'mariorossi', campoExtra: 'x' };
    expect(() => modifyUtenteSchema.parse(input)).toThrow(ZodError);
  })
})

describe('richiestaResetSchema', () => {
  it('accetta username con nota facoltativa', () => {
    expect(richiestaResetSchema.parse({ username: 'mariorossi' })).toEqual({ username: 'mariorossi' });
    expect(richiestaResetSchema.parse({ username: ' mariorossi ', nota: ' segreteria ' }))
      .toEqual({ username: 'mariorossi', nota: 'segreteria' });
  })
  it('rifiuta username vuoto e nota troppo lunga', () => {
    expect(() => richiestaResetSchema.parse({ username: '  ' })).toThrow(ZodError);
    expect(() => richiestaResetSchema.parse({ username: 'mario', nota: 'x'.repeat(201) })).toThrow(ZodError);
  })
})
describe('modifyDocenteSchema', () => {
  it('accetta nome e cognome non vuoti', () => {
    const input = { nome: 'Mario', cognome: 'Rossi', limiteCopie: 10 };
    expect(modifyDocenteSchema.parse(input)).toEqual(input);
  })
  it('rifiuta nome troppo corto', () => {
    const input = { nome: 'A', cognome: 'Rossi', limiteCopie: 10 };
    expect(() => modifyDocenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta cognome troppo corto', () => {
    const input = { nome: 'Mario', cognome: 'R', limiteCopie: 10 };
    expect(() => modifyDocenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta limiteCopie minore di 0', () => {
    const input = { nome: 'Mario', cognome: 'Rossi', limiteCopie: -1 };
    expect(() => modifyDocenteSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta campi extra non permessi (strict)', () => {
    const input = { nome: 'Mario', cognome: 'Rossi', limiteCopie: 10, campoExtra: 'x' };
    expect(() => modifyDocenteSchema.parse(input)).toThrow(ZodError);
  })
})
describe('insertRegistrazioneSchema', () => {
  it('accetta docenteId, copieEffettuate e utenteId non vuoti', () => {
    const input = { docenteId: 1, copieEffettuate: 10, utenteId: '123e4567-e89b-12d3-a456-426614174000' };
    expect(insertRegistrazioneSchema.parse(input)).toEqual(input);
  })
  it('rifiuta docenteId non valido', () => {
    const input = { docenteId: '1', copieEffettuate: 10, utenteId: '123e4567-e89b-12d3-a456-426614174000' };
    expect(() => insertRegistrazioneSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta copieEffettuate minore di 1', () => {
    const input = { docenteId: 1, copieEffettuate: 0, utenteId: '123e4567-e89b-12d3-a456-426614174000' };
    expect(() => insertRegistrazioneSchema.parse(input)).toThrow(ZodError);
  })
  it('rifiuta utenteId non valido', () => {
    const input = { docenteId: 1, copieEffettuate: 10, utenteId: '123e4567-e89b-12d3-a456-4266141740001' };
    expect(() => insertRegistrazioneSchema.parse(input)).toThrow(ZodError);
  })
  it('accetta note opzionale', () => {
    const input = { docenteId: 1, copieEffettuate: 10, utenteId: '123e4567-e89b-12d3-a456-426614174000', note: 'Test note' };
    const result = insertRegistrazioneSchema.parse(input);
    expect(result.note).toBe('Test note');
  })
  it('accetta note undefined', () => {
    const input = { docenteId: 1, copieEffettuate: 10, utenteId: '123e4567-e89b-12d3-a456-426614174000' };
    const result = insertRegistrazioneSchema.parse(input);
    expect(result.note).toBeUndefined();
  })
})

describe('registerSchema', () => {
  it('accetta email valida e password di almeno 8 caratteri', () => {
    const input = { email: 'admin@scuola.it', password: 'password123' };
    expect(registerSchema.parse(input)).toEqual(input);
  });
  it('rifiuta email non valida', () => {
    const input = { email: 'emailinvalida', password: 'password123' };
    expect(() => registerSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta password troppo corta', () => {
    const input = { email: 'admin@scuola.it', password: '1234567' };
    expect(() => registerSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta password vuota', () => {
    const input = { email: 'admin@scuola.it', password: '' };
    expect(() => registerSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta email vuota', () => {
    const input = { email: '', password: 'password123' };
    expect(() => registerSchema.parse(input)).toThrow(ZodError);
  });
});

describe('idParamSchema', () => {
  it('accetta numero positivo', () => {
    const input = { id: 5 };
    expect(idParamSchema.parse(input)).toEqual({ id: 5 });
  });
  it('coerce stringa numerica a numero', () => {
    const input = { id: '5' };
    const result = idParamSchema.parse(input);
    expect(result.id).toBe(5);
  });
  it('rifiuta numero negativo', () => {
    const input = { id: -1 };
    expect(() => idParamSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta zero', () => {
    const input = { id: 0 };
    expect(() => idParamSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta stringa non numerica', () => {
    const input = { id: 'abc' };
    expect(() => idParamSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta numero decimale', () => {
    const input = { id: 5.5 };
    expect(() => idParamSchema.parse(input)).toThrow(ZodError);
  });
});

describe('uuidParamSchema', () => {
  it('accetta UUID valido', () => {
    const uuid = '123e4567-e89b-12d3-a456-426614174000';
    const input = { id: uuid };
    expect(uuidParamSchema.parse(input)).toEqual({ id: uuid });
  });
  it('rifiuta stringa non UUID', () => {
    const input = { id: 'not-a-uuid' };
    expect(() => uuidParamSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta UUID malformato (troppo corto)', () => {
    const input = { id: '123e4567-e89b-12d3-a456' };
    expect(() => uuidParamSchema.parse(input)).toThrow(ZodError);
  });
  it('rifiuta UUID con caratteri invalidi', () => {
    const input = { id: '123e4567-e89b-12d3-a456-42661417400g' };
    expect(() => uuidParamSchema.parse(input)).toThrow(ZodError);
  });
});
