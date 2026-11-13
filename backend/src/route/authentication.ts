import express from 'express';
import { asyncHandler, type ErrorWithStatus } from '../middleware/auth.js';
import type { Request, Response } from 'express';
import { insertIstitutoSchema, loginSchema, registerSchema, createUtenteSchema, changePasswordSchema, type InsertIstituto, type LoginData, type RegisterData, type CreateUtente } from '../../../shared/validation.js';
import { istituti, refreshTokens, utenti } from '../db/schema.js';
import { db } from '../db/index.js';
import { and, eq, gt, isNull } from 'drizzle-orm';
import bcrypt from 'bcrypt';
import jwt from 'jsonwebtoken';
import { JWT_SECRET } from '../middleware/auth.js';
import { createTenantStore } from '../db/tenantStore.js';
import type { LoginResponse, RefreshTokenResponse } from '../../../shared/types.js';
import { isEmailIdentifier } from '../db/utils/auth.js';
import { createAuthLimiter } from '../middleware/rateLimit.js';
import { clearRefreshCookie, readRefreshCookie, setRefreshCookie } from '../middleware/refreshCookie.js';
import { hashRefreshToken, issueRefreshToken, revokeAllRefreshTokens } from '../db/utils/refreshTokens.js';
const router = express.Router();

// Limite ai tentativi sulle rotte pubbliche (vedi middleware/rateLimit.ts)
const authLimiter = createAuthLimiter();

const ACCESS_TOKEN_EXPIRATION = '15m';

router.post('/setup-scuola', authLimiter, asyncHandler(async (req: Request, res: Response) => {
    //ricevo i dati dal body 
    const dataIstituto: InsertIstituto = insertIstitutoSchema.parse(req.body.istituto)
    const dataUtente : RegisterData = registerSchema.parse(req.body.utente)
    const utenteEsistente = await db.select().from(utenti)
        .where(eq(utenti.email, dataUtente.email))
    const istitutoEsistente = await db.select().from(istituti)
        .where(eq(istituti.codiceIstituto, dataIstituto.codiceIstituto))
    if (utenteEsistente.length > 0 || istitutoEsistente.length > 0) {
        const error = new Error('Utente o istituto già esistente') as ErrorWithStatus;
        error.status = 400;
        throw error;

    }
    const result = await db.transaction(async (tx) => {
        const [istituto] = await tx.insert(istituti)
            .values(dataIstituto)
            .returning()
        if (!istituto) {
            const error = new Error('Errore nella creazione dell\'istituto') as ErrorWithStatus;
            error.status = 500;
            throw error;
        }
        
        // Creiamo lo store usando l'istitutoId appena creato
        // Passiamo la transazione tx per garantire atomicità delle operazioni
        const tenantStore = createTenantStore(istituto.id, tx);
        
        // Usiamo lo store per creare l'utente admin
        // Lo store gestisce già hash della password e validazione
        // Type assertion necessaria perché TypeScript non deduce correttamente la union type discriminata
        const utenteAdminData: Extract<CreateUtente, { ruolo: 'admin' }> = {
            ruolo: 'admin',
            email: dataUtente.email,
            password: dataUtente.password,
        };
        const utente = await tenantStore.utenti.create(utenteAdminData);
        
        return { istituto, utente }
    })
    res.status(201).json({
        message: 'Istituto e utente creati con successo',
        istituto: result.istituto,
        utente: result.utente,
    })
}))

router.post('/login', authLimiter, asyncHandler(async (req: Request, res: Response) => {
    const { identifier, password }: LoginData = loginSchema.parse(req.body)
    const isEmail = isEmailIdentifier(identifier)
    const [utente] = await db.select().from(utenti)
        .where(
            isEmail ?
                eq(utenti.email, identifier) :
                eq(utenti.username, identifier)
        )
    if (!utente) {
        const error = new Error('Credenziali non valide') as ErrorWithStatus;
        error.status = 401;
        throw error;
    }

    //verifica password
    const passwordValida = await bcrypt.compare(password, utente.passwordHash)
    if (!passwordValida) {
        const error = new Error('Credenziali non valide') as ErrorWithStatus;
        error.status = 401;
        throw error;
    }
    
    // Recupera dati istituto
    const [istituto] = await db.select().from(istituti).where(eq(istituti.id, utente.istitutoId))
    if (!istituto) {
        const error = new Error('Istituto non trovato') as ErrorWithStatus;
        error.status = 500;
        throw error;
    }
    
    const token = jwt.sign(
        {userId: utente.id, ruolo: utente.ruolo, istitutoId: utente.istitutoId},
        JWT_SECRET,
        {expiresIn: ACCESS_TOKEN_EXPIRATION}

    )
    // Il refresh token va solo nel cookie httpOnly, mai nel corpo della risposta
    const refreshToken = await issueRefreshToken(db, utente.id);
    setRefreshCookie(res, refreshToken);
    const response: LoginResponse = {
        token,
        utente: {
            id: utente.id,
            email: utente.email,
            username: utente.username,
            ruolo: utente.ruolo,
        },
        istituto: {
            id: istituto.id,
            nome: istituto.nome,
            codiceIstituto: istituto.codiceIstituto,
        }
    }
    res.status(200).json(response)
}))
router.post('/logout', asyncHandler(async (req: Request, res: Response) => {
    const refreshToken = readRefreshCookie(req)

    // Idempotente: senza cookie (o con un cookie non valido) non c'è nulla da revocare
    if (refreshToken) {
        await db.update(refreshTokens).set(
            {
                revokedAt: new Date(),
            }).where(eq(refreshTokens.tokenHash, hashRefreshToken(refreshToken)))
    }
    clearRefreshCookie(res)
    res.status(200).json({
        message: 'Logout effettuato con successo',
    })
}))

router.post('/refresh', asyncHandler(async (req: Request, res: Response) => {
    const refreshToken = readRefreshCookie(req)
    if (!refreshToken) {
        const error = new Error('Refresh token non valido o scaduto') as ErrorWithStatus;
        error.status = 401;
        throw error;
    }

    const {  utente, newRefreshToken} = await db.transaction(async (tx) => {
        const [row] = await tx.select().from(refreshTokens)
        .where(
            and(
                eq(refreshTokens.tokenHash, hashRefreshToken(refreshToken)),
                isNull(refreshTokens.revokedAt),
                gt(refreshTokens.expiresAt, new Date())
            )
        )
        if(!row) {
            const error = new Error('Refresh token non valido o scaduto') as ErrorWithStatus;
            error.status = 401;
            throw error;
        }
        const [utente] = await tx.select().from(utenti)
        .where(eq(utenti.id, row.userId))
        if(!utente) {
            const error = new Error('Utente non trovato') as ErrorWithStatus;
            error.status = 404;
            throw error;
        }
        await tx.update(refreshTokens).set({
            revokedAt: new Date(),
        }).where(eq(refreshTokens.id, row.id))
        
        const newRefreshToken = await issueRefreshToken(tx, utente.id);
        return { utente, newRefreshToken }
    })
    const newAccessToken = jwt.sign(
        {userId: utente.id, ruolo: utente.ruolo, istitutoId: utente.istitutoId},
        JWT_SECRET,
        {expiresIn: ACCESS_TOKEN_EXPIRATION}
    )
    setRefreshCookie(res, newRefreshToken)
    const response: RefreshTokenResponse = {
        token: newAccessToken,
    }
    res.status(200).json(response)
}))

const changePasswordRouter = express.Router();

changePasswordRouter.post('/change-password', asyncHandler(async (req: Request, res: Response) => {
    const { currentPassword, newPassword } = changePasswordSchema.parse(req.body);

    const [utente] = await db.select().from(utenti).where(eq(utenti.id, req.user.userId));
    if (!utente) {
        const error = new Error('Utente non trovato') as ErrorWithStatus;
        error.status = 404;
        throw error;
    }

    const passwordAttualeValida = await bcrypt.compare(currentPassword, utente.passwordHash);
    if (!passwordAttualeValida) {
        const error = new Error('Password attuale non corretta') as ErrorWithStatus;
        error.status = 400;
        throw error;
    }

    const stessaPassword = await bcrypt.compare(newPassword, utente.passwordHash);
    if (stessaPassword) {
        const error = new Error('La nuova password deve essere diversa da quella attuale') as ErrorWithStatus;
        error.status = 400;
        throw error;
    }

    const passwordHash = await bcrypt.hash(newPassword, 10);
    // Cambio password, revoca di tutte le sessioni e nuova sessione per chi ha fatto la richiesta:
    // tutto nella stessa transazione, così non restano sessioni vecchie valide.
    const refreshToken = await db.transaction(async (tx) => {
        await tx.update(utenti)
            .set({
                passwordHash,
                updatedAt: new Date(),
            })
            .where(eq(utenti.id, utente.id));
        await revokeAllRefreshTokens(tx, utente.id);
        return issueRefreshToken(tx, utente.id);
    });

    setRefreshCookie(res, refreshToken);
    const response = {
        message: 'Password aggiornata con successo',
    };
    res.status(200).json(response);
}));

export { changePasswordRouter };
export default router;