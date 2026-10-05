import express from 'express';
import { asyncHandler, type ErrorWithStatus } from '../middleware/auth.js';
import type { Request, Response } from 'express';
import { insertIstitutoSchema, loginSchema, registerSchema, createUtenteSchema, changePasswordSchema, type InsertIstituto, type LoginData, type RegisterData, type CreateUtente, type RefreshToken, refreshTokenSchema } from '../../../shared/validation.js';
import { istituti, refreshTokens, utenti } from '../db/schema.js';
import { db } from '../db/index.js';
import { and, eq, gt, isNull, lt } from 'drizzle-orm';
import bcrypt from 'bcrypt';
import jwt from 'jsonwebtoken';
import { JWT_SECRET } from '../middleware/auth.js';
import { createTenantStore } from '../db/tenantStore.js';
import crypto from 'crypto';
import type { LoginResponse, RefreshTokenResponse } from '../../../shared/types.js';
import { isEmailIdentifier } from '../db/utils/auth.js';
const router = express.Router();

const ACCESS_TOKEN_EXPIRATION = '15m';
const REFRESH_TOKEN_DAYS = 7;

router.post('/setup-scuola', asyncHandler(async (req: Request, res: Response) => {
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

router.post('/login', asyncHandler(async (req: Request, res: Response) => {
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
    const refreshToken = crypto.randomBytes(32).toString('hex');
    const expires_at = new Date()
    expires_at.setDate(expires_at.getDate() + REFRESH_TOKEN_DAYS);
    await db.insert(refreshTokens).values({
        token: refreshToken,
        userId: utente.id,
        expiresAt: expires_at,
    })
    const response: LoginResponse = {
        token,
        refreshToken,
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
    const { refreshToken} : RefreshToken = refreshTokenSchema.parse(req.body)

    await db.update(refreshTokens).set(
        {
            revokedAt: new Date(),
        }).where(eq(refreshTokens.token, refreshToken))
    res.status(200).json({
        message: 'Logout effettuato con successo',
    })
}))

router.post('/refresh', asyncHandler(async (req: Request, res: Response) => {
    const { refreshToken} : RefreshToken = refreshTokenSchema.parse(req.body)

    const {  utente, newRefreshToken} = await db.transaction(async (tx) => {
        const [row] = await tx.select().from(refreshTokens)
        .where(
            and(
                eq(refreshTokens.token, refreshToken),
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
        
        const newRefreshToken = crypto.randomBytes(32).toString('hex');
        const expires_at = new Date()
        expires_at.setDate(expires_at.getDate() + REFRESH_TOKEN_DAYS);
        await tx.insert(refreshTokens).values({
            token: newRefreshToken,
            userId: utente.id,
            expiresAt: expires_at,
        })
        return { utente, newRefreshToken }
    })
    const newAccessToken = jwt.sign(
        {userId: utente.id, ruolo: utente.ruolo, istitutoId: utente.istitutoId},
        JWT_SECRET,
        {expiresIn: ACCESS_TOKEN_EXPIRATION}
    )
    const response: RefreshTokenResponse = {
        token: newAccessToken,
        refreshToken: newRefreshToken,
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
    await db.update(utenti)
        .set({
            passwordHash,
            updatedAt: new Date(),
        })
        .where(eq(utenti.id, utente.id));

    res.status(200).json({
        message: 'Password aggiornata con successo',
    });
}));

export { changePasswordRouter };
export default router;