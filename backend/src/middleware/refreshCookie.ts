import type { CookieOptions, Request, Response } from 'express';
import { REFRESH_TOKEN_DAYS } from '../db/utils/refreshTokens.js';

/**
 * Il refresh token viaggia solo in un cookie httpOnly: JavaScript nella pagina non può leggerlo,
 * quindi uno script malevolo non può rubarlo. Lo imposta e lo legge solo il server.
 */
export const REFRESH_COOKIE_NAME = 'refreshToken';

/** I refresh token sono 64 caratteri esadecimali (32 byte casuali). */
const REFRESH_TOKEN_FORMAT = /^[a-f0-9]{64}$/;

/**
 * Il cookie è Secure (solo HTTPS) sempre, tranne nei test e quando COOKIE_SECURE=false.
 * Chrome e Firefox accettano cookie Secure anche su http://localhost; per Safari in locale
 * impostare COOKIE_SECURE=false.
 */
function isSecureCookie(): boolean {
  if (process.env.COOKIE_SECURE !== undefined) return process.env.COOKIE_SECURE === 'true';
  return process.env.NODE_ENV !== 'test';
}

/**
 * Opzioni del cookie. Devono essere identiche quando lo si imposta e quando lo si cancella,
 * altrimenti il browser non riconosce il cookie da eliminare.
 *
 * - httpOnly: non leggibile da JavaScript
 * - secure: solo HTTPS (vedi isSecureCookie)
 * - sameSite strict: il browser lo invia solo per richieste partite dal nostro sito
 * - path /api/auth: viene inviato solo alle rotte di autenticazione, non a tutta l'API
 */
function cookieOptions(): CookieOptions {
  return {
    httpOnly: true,
    secure: isSecureCookie(),
    sameSite: 'strict',
    path: '/api/auth',
  };
}

export function setRefreshCookie(res: Response, token: string): void {
  res.cookie(REFRESH_COOKIE_NAME, token, {
    ...cookieOptions(),
    maxAge: REFRESH_TOKEN_DAYS * 24 * 60 * 60 * 1000,
  });
}

export function clearRefreshCookie(res: Response): void {
  res.clearCookie(REFRESH_COOKIE_NAME, cookieOptions());
}

/** Legge il refresh token dal cookie. Restituisce null se manca o non ha il formato atteso. */
export function readRefreshCookie(req: Request): string | null {
  const value: unknown = req.cookies?.[REFRESH_COOKIE_NAME];
  return typeof value === 'string' && REFRESH_TOKEN_FORMAT.test(value) ? value : null;
}
