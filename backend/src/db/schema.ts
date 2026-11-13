import { check } from 'drizzle-orm/pg-core';
import { integer, pgTable, text, timestamp, varchar, uuid, pgEnum, index } from 'drizzle-orm/pg-core';
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