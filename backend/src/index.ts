import express, { type NextFunction, type Request, type Response } from 'express';
import 'dotenv/config';
import cors from 'cors';
import { authMiddleware, authMiddlewareCambioPassword, requireRole, errorHandler } from './middleware/auth.js';
import { tenantStoreMiddleware } from './middleware/tenantStore.js';
import authenticationRoutes, { changePasswordRouter } from './route/authentication.js';
import docentiRoutes from './route/docenti.js';
import utentiRoutes from './route/utenti.js';
import registrazioniCopieRoutes from './route/registrazioniCopie.js';
import istitutiRoutes from './route/istituti.js';
import helmet from 'helmet';
import cookieParser from 'cookie-parser';
import healthRoutes from './route/health.js';
import { getTrustProxy } from './middleware/rateLimit.js';
const app = express();
// Dietro il proxy di Render: necessario perché il rate limit veda l'IP reale del client
app.set('trust proxy', getTrustProxy());
// Porta del backend: default 3001 per evitare conflitto con frontend Vite su 3000
const PORT = process.env.PORT || 3001;
// Origini ammesse dal CORS: FRONTEND_ORIGIN (produzione) più ALLOWED_ORIGINS (lista separata da virgole,
// es. per provare da un altro dispositivo in rete locale). In sviluppo resta ammesso http://localhost:3000.
export function getAllowedOrigins(env: NodeJS.ProcessEnv = process.env): string[] {
    const origini = [
        env.FRONTEND_ORIGIN || 'http://localhost:3000',
        ...(env.ALLOWED_ORIGINS ?? '').split(','),
    ];
    if (env.NODE_ENV !== 'production') origini.push('http://localhost:3000');
    return [...new Set(origini.map((o) => o.trim()).filter(Boolean))];
}

const allowedOrigins = getAllowedOrigins();

app.use(
    helmet({
        contentSecurityPolicy: false,
        crossOriginEmbedderPolicy: false,
        crossOriginResourcePolicy: false,
        referrerPolicy: {
            policy: 'strict-origin-when-cross-origin',
        },
        hsts: {
            maxAge: 31536000,
            includeSubDomains: true,
            preload: true,
        }
    })
)

app.use(cors({
    origin: (origin, callback) => {
        if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
        // 403 esplicito: vale anche come controllo dell'origine sulle rotte che usano il cookie di sessione
        const error = new Error('Origine non consentita') as Error & { status?: number };
        error.status = 403;
        return callback(error, false);
    },
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
}))
app.use(express.json());
// Necessario per leggere il cookie httpOnly del refresh token
app.use(cookieParser());

// Rotte di health check
app.use('/api/health', healthRoutes);

// Rotte pubbliche (autenticazione)
app.use('/api/auth', authenticationRoutes);

// Cambio password: autenticato, disponibile a tutti i ruoli
app.use('/api/auth', authMiddlewareCambioPassword, changePasswordRouter);

// Rotte protette - richiedono autenticazione e tenantStore
// Docenti: accessibili a admin e collaboratori
app.use('/api/docenti', authMiddleware, tenantStoreMiddleware, docentiRoutes);

// Utenti: accessibili solo agli admin
app.use('/api/utenti', authMiddleware, tenantStoreMiddleware, requireRole('admin'), utentiRoutes);

// Registrazioni Copie: accessibili a admin e collaboratori
app.use('/api/registrazioni-copie', authMiddleware, tenantStoreMiddleware, registrazioniCopieRoutes);

// Istituti: accessibili a admin
app.use('/api/istituti', authMiddleware, tenantStoreMiddleware, requireRole('admin'), istitutiRoutes);
app.use(errorHandler);

// Esporta l'app per i test di integrazione
export { app };

// Avvia il server solo se non siamo in ambiente di test
if (process.env.NODE_ENV !== 'test') {
    app.listen(PORT, () => {
        console.log(`Server is running on port ${PORT}`);
    });
}
