import jwt from 'jsonwebtoken';
import type { Request, Response, NextFunction } from 'express';
import type { JwtPayload } from '../../../shared/types.js';
import { ZodError } from 'zod';
import { extractBearerToken } from '../db/utils/extractBearerToken.js';
import { db } from '../db/index.js';
import { utenti } from '../db/schema.js';
import { eq } from 'drizzle-orm';
import { CAMBIO_PASSWORD_OBBLIGATORIO } from '../../../shared/types.js';


export const asyncHandler = (fn: any) => (req: Request, res: Response, next: NextFunction) => {
    Promise.resolve(fn(req, res, next)).catch(next);
}



declare global {
    namespace Express {
        interface Request {
            user: JwtPayload;
        }
    }
}

export interface ErrorWithStatus extends Error {
    status?: number;
}

// Fallisce all'avvio invece che al primo login se la variabile manca
if (!process.env.JWT_SECRET) {
    throw new Error('JWT_SECRET mancante: impostala nelle variabili d\'ambiente del backend');
}
export const JWT_SECRET: string = process.env.JWT_SECRET;

/**
 * Middleware di autenticazione che verifica:
 * 1. Presenza e validità del token JWT
 * 2. Scadenza del token
 * 3. Esistenza dell'utente nel database
 * 
 * Spiegazione:
 * - Verifica che il token sia valido e non scaduto
 * - Verifica che l'utente esista ancora nel database (sicurezza)
 * - Se l'utente è stato eliminato, blocca immediatamente la richiesta
 * - Imposta req.user con i dati del payload JWT se tutto è valido
 * 
 * Nota: Express gestisce automaticamente i Promise reject nei middleware async,
 * quindi gli errori non gestiti verranno passati all'errorHandler globale
 */
export function createAuthMiddleware(opzioni: { consentiCambioObbligato?: boolean } = {}) {
    return async function authMiddleware(req: Request, res: Response, next: NextFunction) {
        const token = extractBearerToken(req.headers.authorization);
        if(!token) {
            res.status(401).json({ error: 'Token mancante' });
            return
        }
        try {
            // Verifica validità e scadenza del token
            const payload = jwt.verify(token, JWT_SECRET) as JwtPayload;
            
            // Verifica che l'utente esista ancora nel database
            // Questo previene l'uso di token validi per utenti eliminati
            const [utente] = await db.select().from(utenti)
                .where(eq(utenti.id, payload.userId));
            
            if(!utente) {
                res.status(401).json({ error: 'Utente non trovato o account eliminato' });
                return
            }

            // Con una password temporanea si può solo cambiarla: tutto il resto è bloccato
            if(utente.deveCambiarePassword && !opzioni.consentiCambioObbligato) {
                res.status(403).json({
                    error: 'Devi cambiare la password temporanea prima di continuare',
                    code: CAMBIO_PASSWORD_OBBLIGATORIO,
                });
                return
            }
            
            req.user = payload;
            next()
        } catch(error) {
            // jwt.verify lancia un errore se il token è invalido o scaduto
            if(error instanceof jwt.JsonWebTokenError || error instanceof jwt.TokenExpiredError) {
                res.status(401).json({ error: 'Token non valido o scaduto' });
                return
            }
            // Se è un errore del database o altro, passalo all'errorHandler globale
            next(error);
        }
    }
}

/** Middleware di autenticazione per tutte le rotte protette. */
export const authMiddleware = createAuthMiddleware();

/** Come authMiddleware, ma lascia passare chi deve ancora cambiare la password temporanea. */
export const authMiddlewareCambioPassword = createAuthMiddleware({ consentiCambioObbligato: true });

export function requireRole(...roles: Array<'admin' | 'collaboratore'>) {
    return (req: Request, res: Response, next: NextFunction) => {
        if(!req.user) {
            res.status(401).json({ error: 'Non autenticato' });
            return
        }
        if(!roles.includes(req.user.ruolo)) {
            res.status(403).json({ error: 'Non hai i permessi per questa operazione' });
            return // ⚠️ IMPORTANTE: ferma l'esecuzione se non ha i permessi
        }
        next()
    }
}

/** Codici errore Node/Postgres per connessione DB non disponibile */
const DB_UNAVAILABLE_CODES = ['ECONNREFUSED', 'ETIMEDOUT', 'ENOTFOUND', 'ECONNRESET', 'ENETUNREACH'];

/** Messaggio generico quando il servizio (es. database) non è raggiungibile. Non esponiamo query o dettagli tecnici. */
const SERVICE_UNAVAILABLE_MESSAGE = 'Servizio temporaneamente non disponibile. Riprova tra qualche minuto.';

function isDatabaseUnavailableError(err: unknown): boolean {
    if (!err || typeof err !== 'object') return false;
    const e = err as { code?: string; message?: string; cause?: unknown };
    if (typeof e.code === 'string' && DB_UNAVAILABLE_CODES.includes(e.code)) return true;
    if (typeof e.message === 'string' && e.message.includes('Failed query')) return true;
    if (e.cause && isDatabaseUnavailableError(e.cause)) return true;
    return false;
}

export function errorHandler(err: ErrorWithStatus, req: Request, res: Response, next: NextFunction) {
    console.error(err);
    if(err instanceof ZodError) {
        return res.status(400).json({
            error: 'Dati di input non validi',
            details: err.issues
        })
    }
    // Database non raggiungibile: risposta generica 503 senza dettagli (query, host, ecc.)
    if (isDatabaseUnavailableError(err)) {
        return res.status(503).json({ error: SERVICE_UNAVAILABLE_MESSAGE });
    }
    const statusCode = err.status || 500;
    const message = err.message || 'Qualcosa è andato storto!';
    return res.status(statusCode).json({ error: message });
}
