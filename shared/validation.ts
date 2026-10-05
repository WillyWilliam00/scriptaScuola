import { z } from 'zod';

// --- Istituto ---
export const insertIstitutoSchema = z.object({
  nome: z.string().min(3, 'Il nome deve essere lungo almeno 3 caratteri'),
  codiceIstituto: z.string().length(10, 'Il codice istituto deve essere lungo 10 caratteri'),
});

// --- Docente ---
export const insertDocenteSchema = z.object({
  nome: z.string().min(2, 'Inserisci un nome valido'),
  cognome: z.string().min(2, 'Inserisci un cognome valido'),
  limiteCopie: z.number().min(0, 'Il limite di copie deve essere almeno 0'),
  istitutoId: z.number(),
});
export const modifyDocenteSchema = z.object({
  nome: z.string().min(2, 'Inserisci un nome valido').optional(),
  
  cognome: z.string().min(2, 'Inserisci un cognome valido').optional(),
  limiteCopie: z.number().min(0, 'Il limite di copie deve essere almeno 0').optional(),
}).strict();

// --- Registrazione (setup istituto + admin, email sempre obbligatoria) ---
export const registerSchema = z.object({
  email: z.email('Formato email non valido'),
  password: z.string().min(8, 'La password deve essere di almeno 8 caratteri'),
});

// --- Creazione utente (admin crea un nuovo utente dal pannello) ---
// Se ruolo è 'admin' → serve email
// Se ruolo è 'collaboratore' → serve username (senza @)


export const createUtenteSchema = z.discriminatedUnion('ruolo', [
  z.object({
    ruolo: z.literal('admin'),
    email: z.email('Formato email non valido'),
    password: z.string().min(8, 'La password deve essere di almeno 8 caratteri'),
  }),
  z.object({
    ruolo: z.literal('collaboratore'),
    username: z.string()
      .min(3, 'Il username deve essere lungo almeno 3 caratteri')
      .refine((val) => !val.includes('@'), { message: 'Il username non può contenere @' }),
    password: z.string().min(8, 'La password deve essere di almeno 8 caratteri'),
  }),
]);

export const modifyUtenteSchema = z.discriminatedUnion('ruolo', [
  z.object({
    ruolo: z.literal('admin'),
    email: z.email('Formato email non valido'),
    password: z.string().min(8, 'La password deve essere di almeno 8 caratteri').or(z.literal('')).optional(),
  }),
  z.object({
    ruolo: z.literal('collaboratore'),
    username: z.string()
      .min(3, 'Il username deve essere lungo almeno 3 caratteri')
      .refine((val) => !val.includes('@'), { message: 'Il username non può contenere @' }),
    password: z.string().min(8, 'La password deve essere di almeno 8 caratteri').or(z.literal('')).optional(),
  })
]);

const passwordMin8 = z.string().min(8, 'La password deve essere di almeno 8 caratteri');

/**
 * Schema del form di modifica utente: email e username restano entrambi nel state.
 * Si valida solo il campo del ruolo attivo, così un campo vuoto dell'altro ruolo
 * non tiene canSubmit a false.
 */
export const modifyUtenteFormSchema = z.object({
  ruolo: z.enum(['admin', 'collaboratore']),
  email: z.string(),
  username: z.string(),
  password: passwordMin8.or(z.literal('')).optional(),
}).superRefine((data, ctx) => {
  if (data.ruolo === 'admin') {
    const emailResult = z.email('Formato email non valido').safeParse(data.email);
    if (!emailResult.success) {
      ctx.addIssue({
        code: 'custom',
        message: emailResult.error.issues[0]?.message ?? 'Formato email non valido',
        path: ['email'],
      });
    }
    return;
  }

  if (data.username.length < 3) {
    ctx.addIssue({
      code: 'custom',
      message: 'Il username deve essere lungo almeno 3 caratteri',
      path: ['username'],
    });
  }
  if (data.username.includes('@')) {
    ctx.addIssue({
      code: 'custom',
      message: 'Il username non può contenere @',
      path: ['username'],
    });
  }
});

export const changePasswordSchema = z.object({
  currentPassword: passwordMin8,
  newPassword: passwordMin8,
});

export const changePasswordFormSchema = changePasswordSchema.extend({
  confirmPassword: passwordMin8,
}).superRefine((data, ctx) => {
  if (data.newPassword === data.currentPassword) {
    ctx.addIssue({
      code: 'custom',
      message: 'La nuova password deve essere diversa da quella attuale',
      path: ['newPassword'],
    });
  }
  if (data.newPassword !== data.confirmPassword) {
    ctx.addIssue({
      code: 'custom',
      message: 'Le password non coincidono',
      path: ['confirmPassword'],
    });
  }
});

// --- Login (un campo identificativo: email o username) ---
export const loginSchema = z.object({
  identifier: z.string()
    .min(1, 'Inserisci email o username')
    .refine((val) => {
      // Se NON c'è una @, la validazione passa (è un username)
      if (!val.includes('@')) return true;
      
      // Se C'È una @, usiamo lo schema email di Zod per validarlo
      return z.email().safeParse(val).success;
    }, {
      message: 'Inserisci un indirizzo email valido',
    }),
  password: z.string()
  .min(8, 'La password deve essere di almeno 8 caratteri'),
});

// --- Registrazione copie ---
export const insertRegistrazioneSchema = z.object({
  docenteId: z.number(),
  copieEffettuate: z.number().positive('Le copie devono essere almeno 1'),
  utenteId: z.uuid(),
  note: z.string().optional(),
});

export const modifyRegistrazioneSchema = z.object({
  copieEffettuate: z.number().positive('Le copie devono essere almeno 1'),
  note: z.string().optional(),
}).strict();

/**
 * Schema per il form di registrazione copie (solo i campi inseriti dall'utente)
 * Usato con TanStack Form nel componente frontend
 * I valori sono stringhe nel form, quindi validiamo come stringhe e convertiamo dopo
 */
export const createRegistrazioneFormSchema = (limiteCopie: number, copieRimanenti: number) => {
  return z.object({
    copie: z.string()
      .min(1, 'Inserisci un numero di copie')
      .refine(
        (val) => {
          const num = Number(val);
          return !Number.isNaN(num) && Number.isInteger(num);
        },
        { message: 'Inserisci un numero intero valido' }
      )
      .refine(
        (val) => {
          const num = Number(val);
          return num >= 1;
        },
        { message: 'Le copie devono essere almeno 1' }
      )
      .refine(
        (val) => {
          const num = Number(val);
          return num <= limiteCopie;
        },
        { message: `Le copie non possono superare il limite (${limiteCopie})` }
      )
      .refine(
        (val) => {
          const num = Number(val);
          return num <= copieRimanenti;
        },
        { message: `Le copie non possono superare il numero di copie rimaste (${copieRimanenti})` }
      ),
    note: z.string(), // Sempre stringa nel form (può essere vuota)
  });
};

/**
 * Schema per il form di modifica registrazione copie
 * Usato con TanStack Form nel componente frontend per modificare una registrazione esistente
 * I valori sono stringhe nel form, quindi validiamo come stringhe e convertiamo dopo
 */
export const modifyRegistrazioneFormSchema = z.object({
  copieEffettuate: z.string()
    .min(1, 'Inserisci un numero di copie')
    .refine(
      (val) => {
        const num = Number(val);
        return !Number.isNaN(num) && Number.isInteger(num) && num > 0;
      },
      { message: 'Le copie devono essere un numero intero positivo' }
    ),
  note: z.string(), // Sempre stringa nel form (può essere vuota)
});

/**
 * Schema per il form di inserimento docente (solo i campi inseriti dall'utente)
 * Usato con TanStack Form nel componente frontend
 * NOTA: istitutoId NON è incluso perché viene preso automaticamente dal token JWT nel backend
 */
export const insertDocenteFormSchema = z.object({
  nome: z.string().min(2, 'Inserisci un nome valido'),
  cognome: z.string().min(2, 'Inserisci un cognome valido'),
  limiteCopie: z.string()
    .min(1, 'Inserisci il limite di copie')
    .refine(
      (val) => {
        const num = Number(val);
        return !Number.isNaN(num) && Number.isInteger(num);
      },
      { message: 'Inserisci un numero intero valido' }
    )
    .refine(
      (val) => {
        const num = Number(val);
        return num >= 0;
      },
      { message: 'Il limite di copie deve essere almeno 0' }
    ),
});

/**
 * Schema per import bulk di docenti con copie già effettuate
 * Usato per importare docenti da file Excel
 * NOTA: istitutoId NON è incluso perché viene preso automaticamente dal token JWT nel backend
 */
export const bulkImportDocentiSchema = z.object({
  docenti: z.array(
    z.object({
      nome: z.string().min(2, 'Inserisci un nome valido'),
      cognome: z.string().min(2, 'Inserisci un cognome valido'),
      limiteCopie: z.number().min(0, 'Il limite di copie deve essere almeno 0'),
      copieEffettuate: z.number().min(0, 'Le copie effettuate devono essere almeno 0').default(0),
      note: z.string().optional(),
    }).refine(
      (data) => data.copieEffettuate <= data.limiteCopie,
      {
        message: 'Le copie effettuate non possono superare il limite di copie',
        path: ['copieEffettuate'],
      }
    )
  ).min(1, 'Devi inserire almeno un docente'),
});

export const docentiQuerySchema = z.object({
  page: z.coerce.number().int().positive().min(1).catch(1).default(1),
  pageSize: z.coerce.number().int().positive().max(100).catch(20).default(20),
  nome: z.string().optional(),
  cognome: z.string().optional(),
  sortField: z
    .enum([
      'nome',
      'cognome',
      'limiteCopie',
      'copieEffettuate',
      'copieRimanenti',
      'createdAt',
      'updatedAt',
    ])
    .default('createdAt'),
  sortOrder: z.enum(['asc', 'desc']).optional().default('asc'),
});
export const registrazioniCopieQuerySchema = z.object({
  page: z.coerce.number().int().positive().catch(1).default(1),
  pageSize: z.coerce.number().int().positive().max(100).catch(20).default(20),
  docenteId: z.number().optional(),
  utenteId: z.uuid().optional(),
  docenteNome: z.string().optional(),
  docenteCognome: z.string().optional(),
  copieEffettuate: z.coerce.number().int().nonnegative().optional(),
  utenteIdentifier: z.string().optional(),
  sortField: z.enum(['docenteId', 'utenteId', 'createdAt', 'updatedAt', 'docenteNome', 'docenteCognome', 'copieEffettuate', 'utente']).default('createdAt'),
  sortOrder: z.enum(['asc', 'desc']).optional().default('asc'),
})
export const utentiQuerySchema = z.object({
  page: z.coerce.number().int().positive().catch(1).default(1),
  pageSize: z.coerce.number().int().positive().max(100).catch(20).default(20),
  identifier: z.string().min(1, 'Inserisci email o username').optional(),
  ruolo: z.enum(['admin', 'collaboratore']).optional(),
  sortField: z.enum(['ruolo', 'identificativo']).default('ruolo'),
  sortOrder: z.enum(['asc', 'desc']).optional().default('asc'),
  
})

// --- Validazione parametri route ---
// Schema per validare ID numerici dai parametri della route (es. docenti)
export const idParamSchema = z.object({
  id: z.coerce.number().int().positive('L\'ID deve essere un numero intero positivo'),
});

// Schema per validare ID UUID dai parametri della route (es. utenti)
export const uuidParamSchema = z.object({
  id: z.uuid('L\'ID deve essere un UUID valido'),
});

export const refreshTokenSchema = z.object({
 refreshToken: z.string().length(64, 'Il refresh token deve essere lungo 64 caratteri').regex(/^[a-f0-9]+$/, 'Il refresh token deve essere un valore esadecimale'),
});



// --- Tipi derivati dagli schema ---
export type RefreshToken = z.infer<typeof refreshTokenSchema>;
export type InsertIstituto = z.infer<typeof insertIstitutoSchema>;
export type InsertDocente = z.infer<typeof insertDocenteSchema>;
export type RegisterData = z.infer<typeof registerSchema>;
export type CreateUtente = z.infer<typeof createUtenteSchema>;
export type ModifyUtente = z.infer<typeof modifyUtenteSchema>
export type ModifyUtenteForm = z.infer<typeof modifyUtenteFormSchema>;
export type ChangePassword = z.infer<typeof changePasswordSchema>;
export type ChangePasswordForm = z.infer<typeof changePasswordFormSchema>;
export type ModifyDocente = z.infer<typeof modifyDocenteSchema>;
export type LoginData = z.infer<typeof loginSchema>;
export type InsertRegistrazione = z.infer<typeof insertRegistrazioneSchema>;
export type ModifyRegistrazione = z.infer<typeof modifyRegistrazioneSchema>;
export type DocentiQuery = z.infer<typeof docentiQuerySchema>;
export type RegistrazioniCopieQuery = z.infer<typeof registrazioniCopieQuerySchema>;
export type UtentiQuery = z.infer<typeof utentiQuerySchema>;
export type IdParam = z.infer<typeof idParamSchema>;
export type UuidParam = z.infer<typeof uuidParamSchema>;
export type BulkImportDocenti = z.infer<typeof bulkImportDocentiSchema>;