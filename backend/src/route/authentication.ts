import express from 'express';
import { asyncHandler, type ErrorWithStatus } from '../middleware/auth.js';
import type { Request, Response } from 'express';
import { insertIstitutoSchema, loginSchema, registerSchema, changePasswordSchema, richiestaResetSchema, passwordDimenticataSchema, reimpostaPasswordSchema, verificaEmailSchema, type InsertIstituto, type LoginData, type RegisterData, type CreateUtente } from '../../../shared/validation.js';
import { istituti, refreshTokens, richiesteResetPassword, utenti } from '../db/schema.js';
import { inizioFinestraRichieste } from '../db/utils/passwordTemporanea.js';
import { db } from '../db/index.js';
import { and, eq, gt, isNotNull, isNull } from 'drizzle-orm';
import bcrypt from 'bcrypt';
import jwt from 'jsonwebtoken';
import { JWT_SECRET } from '../middleware/auth.js';
import { createTenantStore } from '../db/tenantStore.js';
import type { LoginResponse, RefreshTokenResponse } from '../../../shared/types.js';
import { isEmailIdentifier } from '../db/utils/auth.js';
import { createAuthLimiter } from '../middleware/rateLimit.js';
import { clearRefreshCookie, readRefreshCookie, setRefreshCookie } from '../middleware/refreshCookie.js';
import { hashRefreshToken, issueRefreshToken, revokeAllRefreshTokens } from '../db/utils/refreshTokens.js';
import { consumaTokenEmail, creaTokenEmail, RESET_PASSWORD_MINUTI, VERIFICA_EMAIL_ORE } from '../db/utils/tokenEmail.js';
import { inviaEmail, inviaEmailInBackground } from '../email/index.js';
import { emailReimpostaPassword, emailRichiestaResetCollaboratore, emailVerificaIndirizzo } from '../email/modelli.js';
const router = express.Router();

// Limite ai tentativi sulle rotte pubbliche (vedi middleware/rateLimit.ts)
const authLimiter = createAuthLimiter();
// Più stretto: ogni richiesta finisce davanti agli admin dell'istituto
const richiestaResetLimiter = createAuthLimiter({ limit: 5, windowMs: 60 * 60 * 1000 });
// Ogni richiesta manda un'email: stesso limite stretto (un contatore per rotta)
const passwordDimenticataLimiter = createAuthLimiter({ limit: 5, windowMs: 60 * 60 * 1000 });
const inviaVerificaLimiter = createAuthLimiter({ limit: 5, windowMs: 60 * 60 * 1000 });

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
        };
        // Alla registrazione l'admin sceglie la sua password: niente password temporanea
        const { utente } = await tenantStore.utenti.create(utenteAdminData, { password: dataUtente.password });
        
        return { istituto, utente }
    })

    // Email di verifica all'admin appena registrato. Se l'invio fallisce la registrazione
    // resta valida: l'email si può chiedere di nuovo dal Profilo.
    if (result.utente.email) {
        const token = await creaTokenEmail(db, result.utente.id, 'verifica_email', result.utente.email);
        inviaEmailInBackground(emailVerificaIndirizzo(result.utente.email, token, VERIFICA_EMAIL_ORE));
    }

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

    // Password temporanea non cambiata in tempo: serve un nuovo reset da parte degli admin
    if (utente.deveCambiarePassword && utente.passwordTemporaneaScadeAt && utente.passwordTemporaneaScadeAt <= new Date()) {
        const error = new Error('La password temporanea è scaduta. Chiedi agli amministratori del tuo istituto di generarne una nuova.') as ErrorWithStatus;
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
            deveCambiarePassword: utente.deveCambiarePassword,
            emailVerificata: utente.emailVerificataAt !== null,
            haPassword: !!utente.passwordHash,
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

const MESSAGGIO_RICHIESTA_RESET = 'Se lo username esiste, gli amministratori del tuo istituto hanno ricevuto la richiesta. Ti daranno una password temporanea.';

/**
 * POST /api/auth/richiesta-reset
 * Un collaboratore che ha dimenticato la password chiede agli admin del suo istituto
 * una password temporanea. Risponde sempre allo stesso modo, così non si può scoprire
 * quali username esistono. Al massimo una richiesta aperta per utenza.
 */
router.post('/richiesta-reset', richiestaResetLimiter, asyncHandler(async (req: Request, res: Response) => {
    const { username, nota } = richiestaResetSchema.parse(req.body);

    const [utente] = await db.select({ id: utenti.id, istitutoId: utenti.istitutoId })
        .from(utenti)
        .where(and(eq(utenti.username, username), eq(utenti.ruolo, 'collaboratore')));

    if (utente) {
        const [richiestaAperta] = await db.select({ id: richiesteResetPassword.id })
            .from(richiesteResetPassword)
            .where(and(
                eq(richiesteResetPassword.utenteId, utente.id),
                isNull(richiesteResetPassword.gestitaAt),
                gt(richiesteResetPassword.createdAt, inizioFinestraRichieste()),
            ));
        if (!richiestaAperta) {
            await db.insert(richiesteResetPassword).values({
                utenteId: utente.id,
                istitutoId: utente.istitutoId,
                nota: nota || null,
            });
            // Avviso via email agli admin dell'istituto, solo a quelli con l'email verificata
            const admin = await db.select({ email: utenti.email })
                .from(utenti)
                .where(and(
                    eq(utenti.istitutoId, utente.istitutoId),
                    eq(utenti.ruolo, 'admin'),
                    isNotNull(utenti.emailVerificataAt),
                ));
            for (const { email } of admin) {
                if (email) inviaEmailInBackground(emailRichiestaResetCollaboratore(email, username, nota || null));
            }
        }
    }

    res.status(200).json({ message: MESSAGGIO_RICHIESTA_RESET });
}));

const MESSAGGIO_PASSWORD_DIMENTICATA = 'Se l\'indirizzo appartiene a un amministratore, riceverai un\'email con il link per scegliere una nuova password. Controlla anche la cartella spam.';

/**
 * POST /api/auth/password-dimenticata
 * Un admin che ha dimenticato la password riceve via email un link monouso (valido 60 minuti).
 * Risponde sempre allo stesso modo, anche se l'email non esiste o è di un collaboratore,
 * e l'email parte in background, così né il testo né i tempi di risposta rivelano chi è registrato.
 */
router.post('/password-dimenticata', passwordDimenticataLimiter, asyncHandler(async (req: Request, res: Response) => {
    const { email } = passwordDimenticataSchema.parse(req.body);

    const [utente] = await db.select({ id: utenti.id, email: utenti.email })
        .from(utenti)
        .where(and(eq(utenti.email, email), eq(utenti.ruolo, 'admin')));

    if (utente?.email) {
        const token = await creaTokenEmail(db, utente.id, 'reset_password', utente.email);
        inviaEmailInBackground(emailReimpostaPassword(utente.email, token, RESET_PASSWORD_MINUTI));
    }

    res.status(200).json({ message: MESSAGGIO_PASSWORD_DIMENTICATA });
}));

/**
 * POST /api/auth/reimposta-password
 * Imposta la nuova password usando il link ricevuto via email.
 * Il link vale una volta sola; tutte le sessioni aperte vengono chiuse.
 * Aver aperto il link dimostra anche che l'email è dell'utente: la segniamo come verificata.
 */
router.post('/reimposta-password', authLimiter, asyncHandler(async (req: Request, res: Response) => {
    const { token, newPassword } = reimpostaPasswordSchema.parse(req.body);
    const passwordHash = await bcrypt.hash(newPassword, 10);

    await db.transaction(async (tx) => {
        const consumato = await consumaTokenEmail(tx, token, 'reset_password');
        const [utente] = consumato
            ? await tx.select().from(utenti).where(eq(utenti.id, consumato.utenteId))
            : [];
        // Se nel frattempo l'email dell'account è cambiata, il link non vale più
        if (!consumato || !utente || utente.email !== consumato.email) {
            const error = new Error('Il link non è valido o è scaduto. Chiedine uno nuovo da "Password dimenticata?".') as ErrorWithStatus;
            error.status = 400;
            throw error;
        }
        await tx.update(utenti).set({
            passwordHash,
            deveCambiarePassword: false,
            passwordTemporaneaScadeAt: null,
            emailVerificataAt: utente.emailVerificataAt ?? new Date(),
            updatedAt: new Date(),
        }).where(eq(utenti.id, utente.id));
        await revokeAllRefreshTokens(tx, utente.id);
    });

    res.status(200).json({ message: 'Password aggiornata. Ora puoi accedere con la nuova password.' });
}));

/**
 * POST /api/auth/verifica-email
 * Conferma l'indirizzo email dal link ricevuto. Non serve essere collegati.
 */
router.post('/verifica-email', authLimiter, asyncHandler(async (req: Request, res: Response) => {
    const { token } = verificaEmailSchema.parse(req.body);

    const verificato = await db.transaction(async (tx) => {
        const consumato = await consumaTokenEmail(tx, token, 'verifica_email');
        if (!consumato) return false;
        const [aggiornato] = await tx.update(utenti)
            .set({ emailVerificataAt: new Date(), updatedAt: new Date() })
            .where(and(eq(utenti.id, consumato.utenteId), eq(utenti.email, consumato.email)))
            .returning({ id: utenti.id });
        return !!aggiornato;
    });

    if (!verificato) {
        const error = new Error('Il link non è valido o è scaduto. Puoi chiederne uno nuovo dal tuo Profilo.') as ErrorWithStatus;
        error.status = 400;
        throw error;
    }
    res.status(200).json({ message: 'Email confermata.' });
}));

const changePasswordRouter = express.Router();

/**
 * GET /api/auth/me
 * Dati aggiornati dell'utente collegato (per esempio dopo aver confermato l'email in un'altra scheda).
 */
changePasswordRouter.get('/me', asyncHandler(async (req: Request, res: Response) => {
    const [utente] = await db.select().from(utenti).where(eq(utenti.id, req.user.userId));
    if (!utente) {
        const error = new Error('Utente non trovato') as ErrorWithStatus;
        error.status = 404;
        throw error;
    }
    const response: LoginResponse['utente'] = {
        id: utente.id,
        email: utente.email,
        username: utente.username,
        ruolo: utente.ruolo,
        deveCambiarePassword: utente.deveCambiarePassword,
        emailVerificata: utente.emailVerificataAt !== null,
        haPassword: !!utente.passwordHash,
    };
    res.status(200).json(response);
}));

/**
 * POST /api/auth/invia-verifica-email
 * Rimanda l'email di verifica all'utente collegato (solo chi ha un'email, cioè gli admin).
 */
changePasswordRouter.post('/invia-verifica-email', inviaVerificaLimiter, asyncHandler(async (req: Request, res: Response) => {
    const [utente] = await db.select().from(utenti).where(eq(utenti.id, req.user.userId));
    if (!utente?.email) {
        const error = new Error('Questa utenza non ha un indirizzo email') as ErrorWithStatus;
        error.status = 400;
        throw error;
    }
    if (utente.emailVerificataAt) {
        const error = new Error('L\'email è già confermata') as ErrorWithStatus;
        error.status = 400;
        throw error;
    }
    const token = await creaTokenEmail(db, utente.id, 'verifica_email', utente.email);
    // Qui l'utente aspetta l'esito: se Resend rifiuta, glielo diciamo
    try {
        await inviaEmail(emailVerificaIndirizzo(utente.email, token, VERIFICA_EMAIL_ORE));
    } catch (err) {
        console.error(err);
        const error = new Error('Non è stato possibile inviare l\'email. Riprova più tardi.') as ErrorWithStatus;
        error.status = 502;
        throw error;
    }
    res.status(200).json({ message: `Ti abbiamo inviato un'email a ${utente.email}.` });
}));

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
                // La password scelta dall'utente sostituisce quella temporanea
                deveCambiarePassword: false,
                passwordTemporaneaScadeAt: null,
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