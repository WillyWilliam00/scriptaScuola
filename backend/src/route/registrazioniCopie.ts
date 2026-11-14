import express from 'express';
import { asyncHandler, requireRole } from '../middleware/auth.js';
import type { Request, Response } from 'express';
import { idParamSchema, insertRegistrazioneSchema, modifyRegistrazioneSchema, registrazioniCopieQuerySchema, type InsertRegistrazione, type ModifyRegistrazione } from '../../../shared/validation.js';
import type { RegistrazioniCopiePaginatedResponse } from '../../../shared/types.js';
import { toCsvValue } from '../db/utils/pagination.js';
const router = express.Router();




router.get('/', requireRole('admin'), asyncHandler(async (req: Request, res: Response) => {
    if(!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    const query = registrazioniCopieQuerySchema.parse(req.query);
    const result: RegistrazioniCopiePaginatedResponse = await req.tenantStore.registrazioniCopie.getPaginated(query);
    res.status(200).json(result);
}))

/**
 * GET /api/registrazioni-copie/export
 * Esporta in CSV tutte le registrazioni che corrispondono ai filtri (senza paginazione)
 */
router.get('/export', requireRole('admin'), asyncHandler(async (req: Request, res: Response) => {
    if(!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    const query = registrazioniCopieQuerySchema.parse(req.query);
    const rows = await req.tenantStore.registrazioniCopie.getAllForExport(query);

    const header = [
        'ID',
        'Docente nome',
        'Docente cognome',
        'Copie effettuate',
        'Utente',
        'Note',
        'Creato il',
        'Aggiornato il',
    ];

    const lines = rows.map((row) => {
        const utente =
            row.utenteUsername ??
            row.utenteEmail ??
            '';
        const createdAt = (row.createdAt instanceof Date) ? row.createdAt.toISOString() : String(row.createdAt);
        const updatedAt = (row.updatedAt instanceof Date) ? row.updatedAt.toISOString() : String(row.updatedAt);

        return [
            row.id,
            row.docenteNome ?? '',
            row.docenteCognome ?? '',
            row.copieEffettuate,
            utente,
            row.note ?? '',
            createdAt,
            updatedAt,
        ].map(toCsvValue).join(';');
    });

    const csv = [header.map(toCsvValue).join(';'), ...lines].join('\n');

    const today = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="registrazioni-${today}.csv"`);
    res.status(200).send('\uFEFF' + csv);
}))

router.post('/new-registrazione', asyncHandler(async (req: Request, res: Response) => {
    if(!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    if(!req.user) {
        return res.status(401).json({ error: 'Utente non autenticato' });
    }
    // L'autore della registrazione è sempre l'utente del token: un eventuale utenteId nel body viene ignorato
    const data = insertRegistrazioneSchema.omit({ utenteId: true }).parse(req.body);
    const result: InsertRegistrazione = await req.tenantStore.registrazioniCopie.create({ ...data, utenteId: req.user.userId });
    res.status(201).json(result);
}))

router.put('/update-registrazione/:id', requireRole('admin'), asyncHandler(async (req: Request, res: Response) => {
    if(!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    const { id } = idParamSchema.parse(req.params);
    const data = modifyRegistrazioneSchema.parse(req.body);
    const result = await req.tenantStore.registrazioniCopie.update(id, data);
    res.status(200).json({
        message: 'Registrazione aggiornata con successo',
        data: result,
    });
}))

router.delete('/delete-registrazione/:id', requireRole('admin'), asyncHandler(async (req: Request, res: Response) => {
    if(!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    const { id } = idParamSchema.parse(req.params);
    const result = await req.tenantStore.registrazioniCopie.delete(id);
    res.status(200).json(result);
}))

router.delete('/delete-all', requireRole('admin'), asyncHandler(async (req: Request, res: Response) => {
    if(!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    await req.tenantStore.registrazioniCopie.deleteAll();
    res.status(200).json({ message: 'Tutte le registrazioni copie dell\'istituto sono state eliminate.' });
}))

export default router;