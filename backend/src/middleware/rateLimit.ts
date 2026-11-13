import rateLimit, { type Options } from 'express-rate-limit';

const QUINDICI_MINUTI = 15 * 60_000;

/**
 * Limite ai tentativi sulle rotte pubbliche di autenticazione (login, registrazione scuola).
 * Default: 10 richieste ogni 15 minuti per indirizzo IP, poi 429.
 *
 * Nei test di integrazione il limite è disattivato (tutte le richieste arrivano dallo stesso IP);
 * il comportamento è verificato a parte in tests/middleware/rateLimit.test.ts.
 */
export function createAuthLimiter(options: Partial<Options> = {}) {
  return rateLimit({
    windowMs: QUINDICI_MINUTI,
    limit: 10,
    standardHeaders: 'draft-7',
    legacyHeaders: false,
    message: { error: 'Troppi tentativi. Riprova tra qualche minuto.' },
    skip: () => process.env.NODE_ENV === 'test',
    ...options,
  });
}

/**
 * Numero di proxy fidati davanti al backend.
 * Serve a Express per leggere l'IP reale del client da X-Forwarded-For:
 * se il numero è troppo basso, tutti gli utenti sembrano arrivare dallo stesso indirizzo
 * (quello del proxy) e condividono un unico limite di tentativi.
 *
 * Default 2: il browser chiama il sito su Vercel, che inoltra /api a Render (vercel.json),
 * e Render ha a sua volta un proxy davanti al backend.
 * Dopo un deploy si verifica con GET /api/health/ip: deve mostrare il proprio IP pubblico.
 */
export function getTrustProxy(): number {
  const value = Number(process.env.TRUST_PROXY ?? 2);
  return Number.isInteger(value) && value >= 0 ? value : 2;
}
