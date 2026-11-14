import crypto from 'crypto';
import { and, eq, gt, isNull } from 'drizzle-orm';
import type { NodePgDatabase } from 'drizzle-orm/node-postgres';
import * as schema from '../schema.js';
import { tokenEmail } from '../schema.js';

type DbInstance = NodePgDatabase<typeof schema>;
type Transaction = Parameters<Parameters<DbInstance['transaction']>[0]>[0];
type Db = DbInstance | Transaction;

export type TipoTokenEmail = 'verifica_email' | 'reset_password';

/** Validità del link di verifica dell'email. */
export const VERIFICA_EMAIL_ORE = 48;
/** Validità del link per reimpostare la password: breve, perché dà accesso all'account. */
export const RESET_PASSWORD_MINUTI = 60;

export const FORMATO_TOKEN_EMAIL = /^[a-f0-9]{64}$/;

function hashToken(token: string): string {
  return crypto.createHash('sha256').update(token).digest('hex');
}

/**
 * Crea un token monouso e lo restituisce in chiaro (va solo nell'email).
 * Invalida i token dello stesso tipo ancora aperti per quell'utente:
 * vale sempre e solo l'ultimo link inviato.
 */
export async function creaTokenEmail(db: Db, utenteId: string, tipo: TipoTokenEmail, email: string): Promise<string> {
  const token = crypto.randomBytes(32).toString('hex');
  const durataMs = tipo === 'verifica_email'
    ? VERIFICA_EMAIL_ORE * 60 * 60 * 1000
    : RESET_PASSWORD_MINUTI * 60 * 1000;

  await db.update(tokenEmail)
    .set({ usatoAt: new Date() })
    .where(and(eq(tokenEmail.utenteId, utenteId), eq(tokenEmail.tipo, tipo), isNull(tokenEmail.usatoAt)));

  await db.insert(tokenEmail).values({
    utenteId,
    tipo,
    tokenHash: hashToken(token),
    email,
    expiresAt: new Date(Date.now() + durataMs),
  });
  return token;
}

/**
 * Consuma un token: se è valido (giusto tipo, non usato, non scaduto) lo segna come usato
 * e restituisce utente ed email a cui era stato inviato; altrimenti null.
 * L'aggiornamento è una sola istruzione, quindi due richieste parallele non possono usarlo entrambe.
 */
export async function consumaTokenEmail(db: Db, token: string, tipo: TipoTokenEmail): Promise<{ utenteId: string; email: string } | null> {
  if (!FORMATO_TOKEN_EMAIL.test(token)) return null;
  const [riga] = await db.update(tokenEmail)
    .set({ usatoAt: new Date() })
    .where(and(
      eq(tokenEmail.tokenHash, hashToken(token)),
      eq(tokenEmail.tipo, tipo),
      isNull(tokenEmail.usatoAt),
      gt(tokenEmail.expiresAt, new Date()),
    ))
    .returning({ utenteId: tokenEmail.utenteId, email: tokenEmail.email });
  return riga ?? null;
}
