import express from 'express';
import bcrypt from 'bcrypt';
import { z } from 'zod';
import { and, eq } from 'drizzle-orm';
import { asyncHandler } from '../middleware/auth.js';
import type { Request, Response } from 'express';
import { db } from '../db/index.js';
import { istituti, utenti } from '../db/schema.js';

const router = express.Router();

const confermaEliminazioneSchema = z.object({
    // Nome dell'istituto digitato dall'admin: vale per tutti, anche per chi accede solo con Google/Microsoft
    nomeIstituto: z.string().trim().min(1, "Scrivi il nome dell'istituto per confermare"),
    // Password dell'admin: obbligatoria se l'account ne ha una
    password: z.string().optional(),
});

// Elimina l'istituto dell'utente autenticato (istitutoId dal token, via tenantStore).
// Cancella a cascata docenti, utenti e registrazioni: serve una conferma esplicita.
router.delete('/delete-istituto', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' })
    }
    const { nomeIstituto, password } = confermaEliminazioneSchema.parse(req.body ?? {});

    const [utente] = await db.select({ passwordHash: utenti.passwordHash }).from(utenti)
        .where(and(eq(utenti.id, req.user.userId), eq(utenti.istitutoId, req.user.istitutoId)));
    const [istituto] = await db.select({ nome: istituti.nome }).from(istituti)
        .where(eq(istituti.id, req.user.istitutoId));
    if (!utente || !istituto) {
        return res.status(404).json({ error: 'Istituto non trovato' });
    }

    if (nomeIstituto !== istituto.nome.trim()) {
        return res.status(400).json({ error: "Il nome dell'istituto non corrisponde" });
    }

    // La colonna diventerà facoltativa per gli admin che accedono solo con Google/Microsoft
    const passwordHash: string | null = utente.passwordHash;
    if (passwordHash) {
        if (!password) {
            return res.status(401).json({ error: 'Inserisci la password per confermare' });
        }
        if (!(await bcrypt.compare(password, passwordHash))) {
            return res.status(401).json({ error: 'Password errata' });
        }
    }

    const result = await req.tenantStore.istituti.delete();
    res.status(200).json(result)
}))

export default router;
