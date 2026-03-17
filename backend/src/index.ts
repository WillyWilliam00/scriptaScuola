import express, { type NextFunction, type Request, type Response } from 'express';
import 'dotenv/config';
import cors from 'cors';
import { authMiddleware, requireRole, errorHandler } from './middleware/auth.js';
import { tenantStoreMiddleware } from './middleware/tenantStore.js';
import authenticationRoutes from './route/authentication.js';
import docentiRoutes from './route/docenti.js';
import utentiRoutes from './route/utenti.js';
import registrazioniCopieRoutes from './route/registrazioniCopie.js';
import istitutiRoutes from './route/istituti.js';
import helmet from 'helmet';
import healthRoutes from './route/health.js';
const app = express();
// Porta del backend: default 3001 per evitare conflitto con frontend Vite su 3000
const PORT = process.env.PORT || 3001;
const origin = process.env.FRONTEND_ORIGIN || 'http://localhost:3000';

const allowedOrigins = [
    origin,
    'http://localhost:3000',
    'http://192.168.1.50:3000',

];

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
        const msg = 'The CORS policy for this site does not allow access from the specified Origin.';
        return callback(new Error(msg), false);
    },
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
}))
app.use(express.json());

// Rotte di health check
app.use('/api/health', healthRoutes);

// Rotte pubbliche (autenticazione)
app.use('/api/auth', authenticationRoutes);

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
