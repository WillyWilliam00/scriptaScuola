import crypto from 'crypto';
import { and, eq, isNull } from 'drizzle-orm';
import type { NodePgDatabase } from 'drizzle-orm/node-postgres';
import * as schema from '../schema.js';
import { refreshTokens } from '../schema.js';

type DbInstance = NodePgDatabase<typeof schema>;
type Transaction = Parameters<Parameters<DbInstance['transaction']>[0]>[0];

export const REFRESH_TOKEN_DAYS = 7;

/**
 * Hash SHA-256 del refresh token.
 * Il token ha 256 bit di entropia, quindi basta un hash veloce (niente bcrypt):
 * nel database resta solo l'hash, mai il token in chiaro.
 */
export function hashRefreshToken(token: string): string {
  return crypto.createHash('sha256').update(token).digest('hex');
}

/** Genera un refresh token casuale (64 caratteri esadecimali). */
export function generateRefreshToken(): string {
  return crypto.randomBytes(32).toString('hex');
}

/**
 * Crea un refresh token per l'utente: salva l'hash e restituisce il token in chiaro,
 * che va consegnato al client e non è più recuperabile dal database.
 */
export async function issueRefreshToken(db: DbInstance | Transaction, userId: string): Promise<string> {
  const token = generateRefreshToken();
  const expiresAt = new Date();
  expiresAt.setDate(expiresAt.getDate() + REFRESH_TOKEN_DAYS);
  await db.insert(refreshTokens).values({
    tokenHash: hashRefreshToken(token),
    userId,
    expiresAt,
  });
  return token;
}

/** Revoca tutti i refresh token ancora attivi dell'utente (tutte le sue sessioni). */
export async function revokeAllRefreshTokens(db: DbInstance | Transaction, userId: string): Promise<void> {
  await db.update(refreshTokens)
    .set({ revokedAt: new Date() })
    .where(and(eq(refreshTokens.userId, userId), isNull(refreshTokens.revokedAt)));
}
