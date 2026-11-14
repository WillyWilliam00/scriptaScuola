import { check } from 'drizzle-orm/pg-core';
import { integer, pgTable, text, timestamp, varchar, uuid, pgEnum, index, boolean } from 'drizzle-orm/pg-core';
import { sql } from 'drizzle-orm';

export const istituti = pgTable('istituti', {
    id: integer('id').generatedAlwaysAsIdentity().primaryKey(),
    nome: text('nome').notNull(),
    codiceIstituto: varchar('codice_istituto', { length: 10 }).notNull().unique(),
    createdAt: timestamp('created_at').notNull().defaultNow(),
    updatedAt: timestamp('updated_at').notNull().defaultNow(),
})

export const docenti = pgTable('docenti', {
    id: integer('id').generatedAlwaysAsIdentity().primaryKey(),
    createdAt: timestamp('created_at').notNull().defaultNow(),
    updatedAt: timestamp('updated_at').notNull().defaultNow(),
    nome: text('nome').notNull(),
    cognome: text('cognome').notNull(),
    limiteCopie: integer('limite_copie').notNull().default(0),
    istitutoId: integer('istituto_id').notNull().references(() => istituti.id, { onDelete: 'cascade' }),
}, (table) => [
    index('idx_docenti_istituto_id').on(table.istitutoId),
]);
export const ruoloEnum = pgEnum('ruolo', ['admin', 'collaboratore']);
export const utenti = pgTable('utenti', {
    id: uuid('id').primaryKey().defaultRandom(), // Aggiunto defaultRandom() per generare ID automaticamente
    createdAt: timestamp('created_at').notNull().defaultNow(),
    updatedAt: timestamp('updated_at').notNull().defaultNow(),
    username: varchar('username', { length: 255 }).unique(),
    email: varchar('email', { length: 255 }).unique(),
    istitutoId: integer('istituto_id').notNull().references(() => istituti.id, { onDelete: 'cascade' }),
    ruolo: ruoloEnum('ruolo').notNull().default('collaboratore'),
    passwordHash: varchar('password_hash', { length: 255 }).notNull(),
    // Password temporanea generata dalla piattaforma (nuova utenza o reset):
    // l'utente deve cambiarla al primo accesso, e scade dopo 48 ore.
    deveCambiarePassword: boolean('deve_cambiare_password').notNull().default(false),
    passwordTemporaneaScadeAt: timestamp('password_temporanea_scade_at'),
    // Quando l'utente ha confermato l'email cliccando il link ricevuto (null = non verificata).
    emailVerificataAt: timestamp('email_verificata_at'),
}, (table) => [
    check('utenti_at_least_one_identifier', sql`${table.username} IS NOT NULL OR ${table.email} IS NOT NULL`),
    index('idx_utenti_istituto_id').on(table.istitutoId),
]);



// 4. Registrazioni Copie
export const registrazioniCopie = pgTable('registrazioni_copie', {
    id: integer('id').generatedAlwaysAsIdentity().primaryKey(),
    createdAt: timestamp('created_at').notNull().defaultNow(),
    updatedAt: timestamp('updated_at').notNull().defaultNow(),
    docenteId: integer('docente_id').notNull().references(() => docenti.id, { onDelete: 'cascade' }),
    copieEffettuate: integer('copie_effettuate').notNull().default(0),
    istitutoId: integer('istituto_id').notNull().references(() => istituti.id, { onDelete: 'cascade' }),
    utenteId: uuid('utente_id').references(() => utenti.id, { onDelete: 'set null' }),
    note: text('note').default(''),
}, (table) => [
    index('idx_registrazioni_copie_docente_id').on(table.docenteId),
    index('idx_registrazioni_copie_utente_id').on(table.utenteId),
    index('idx_registrazioni_copie_istituto_id').on(table.istitutoId),
]);


export const refreshTokens = pgTable('refresh_tokens', {
    id: uuid('id').primaryKey().defaultRandom(),
    // Hash SHA-256 del refresh token (mai il token in chiaro).
    // La colonna nel database si chiama ancora `token` per non richiedere una migrazione.
    tokenHash: varchar('token', { length: 255 }).notNull().unique(),
    userId: uuid('user_id').notNull().references(() => utenti.id, { onDelete: 'cascade' }),
    expiresAt: timestamp('expires_at').notNull(),
    createdAt: timestamp('created_at').notNull().defaultNow(),
    revokedAt: timestamp('revoked_at'),
}, (table) => [

    index('idx_refresh_tokens_user_id').on(table.userId),
    index('idx_refresh_tokens_token').on(table.tokenHash),
])

// Richieste di reset password inviate dai collaboratori dalla pagina "Password dimenticata?".
// Le vedono gli admin dell'istituto in Gestione utenze.
export const esitoRichiestaResetEnum = pgEnum('esito_richiesta_reset', ['password_generata', 'ignorata']);
export const richiesteResetPassword = pgTable('richieste_reset_password', {
    id: uuid('id').primaryKey().defaultRandom(),
    utenteId: uuid('utente_id').notNull().references(() => utenti.id, { onDelete: 'cascade' }),
    istitutoId: integer('istituto_id').notNull().references(() => istituti.id, { onDelete: 'cascade' }),
    nota: varchar('nota', { length: 200 }),
    createdAt: timestamp('created_at').notNull().defaultNow(),
    gestitaAt: timestamp('gestita_at'),
    gestitaDa: uuid('gestita_da').references(() => utenti.id, { onDelete: 'set null' }),
    esito: esitoRichiestaResetEnum('esito'),
}, (table) => [
    index('idx_richieste_reset_istituto_id').on(table.istitutoId),
    index('idx_richieste_reset_utente_id').on(table.utenteId),
]);


// Token monouso inviati via email: verifica dell'indirizzo e recupero password degli admin.
// Nel database c'è solo l'hash SHA-256 del token, come per i refresh token.
export const tipoTokenEmailEnum = pgEnum('tipo_token_email', ['verifica_email', 'reset_password']);
export const tokenEmail = pgTable('token_email', {
    id: uuid('id').primaryKey().defaultRandom(),
    utenteId: uuid('utente_id').notNull().references(() => utenti.id, { onDelete: 'cascade' }),
    tipo: tipoTokenEmailEnum('tipo').notNull(),
    tokenHash: varchar('token_hash', { length: 64 }).notNull().unique(),
    // Indirizzo a cui è stato inviato: se nel frattempo l'email cambia, il token non vale più
    email: varchar('email', { length: 255 }).notNull(),
    expiresAt: timestamp('expires_at').notNull(),
    usatoAt: timestamp('usato_at'),
    createdAt: timestamp('created_at').notNull().defaultNow(),
}, (table) => [
    index('idx_token_email_utente_id').on(table.utenteId),
]);
