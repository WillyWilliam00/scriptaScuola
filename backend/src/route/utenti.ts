import express from 'express';
import { asyncHandler, type ErrorWithStatus } from '../middleware/auth.js';
import type { Request, Response } from 'express';
import { utentiQuerySchema, createUtenteSchema, modifyUtenteSchema, uuidParamSchema, type UtentiQuery, type CreateUtente, type ModifyUtente } from '../../../shared/validation.js';
import { utenti } from '../db/schema.js';
import { db } from '../db/index.js';
import { and, eq, sql } from 'drizzle-orm';
import { toCsvValue } from '../db/utils/pagination.js';
const router = express.Router();



/**
 * GET /api/utenti
 * Ottiene la lista paginata degli utenti con filtri e ordinamento
 */
router.get('/', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }

    const query = utentiQuerySchema.parse(req.query);
    const myUserId = req.user.userId;
    const result = await req.tenantStore.utenti.getPaginated({
        ...query
    }, myUserId);
    
    res.status(200).json(result);
}));

/**
 * GET /api/utenti/export
 * Esporta in CSV tutti gli utenti che corrispondono ai filtri (senza paginazione)
 */
router.get('/export', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }

    const query = utentiQuerySchema.parse(req.query);
    const myUserId = req.user.userId;
    const rows = await req.tenantStore.utenti.getAllForExport(query, myUserId);

    const header = [
        'Identificativo',
        'Ruolo',
        'Creato il',
        'Aggiornato il',
    ];

    const lines = rows.map((row) => {
        const identificativo = row.ruolo === 'admin' ? row.email : row.username;
        const createdAt = (row.createdAt instanceof Date) ? row.createdAt.toISOString() : String(row.createdAt);
        const updatedAt = (row.updatedAt instanceof Date) ? row.updatedAt.toISOString() : String(row.updatedAt);

        return [
            identificativo ?? '',
            row.ruolo,
            createdAt,
            updatedAt,
        ].map(toCsvValue).join(';');
    });

    const csv = [header.map(toCsvValue).join(';'), ...lines].join('\n');

    const today = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="utenti-${today}.csv"`);
    res.status(200).send('\uFEFF' + csv);
}));

/**
 * POST /api/utenti
 * Crea un nuovo utente (solo admin)
 */
router.post('/new-utente', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }

    const data = createUtenteSchema.parse(req.body);
    // La password la genera la piattaforma: torna una sola volta, qui, da consegnare all'utente
    const { utente, passwordTemporanea, scadeAt } = await req.tenantStore.utenti.create(data);
    
    res.status(201).json({
        message: 'Utente creato con successo',
        data: utente,
        passwordTemporanea,
        scadeAt,
    });
}));

/**
 * PUT /api/utenti/:id
 * Aggiorna un utente esistente (solo admin)
 */
router.put('/update-utente/:id', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }

    // Valida l'ID come UUID (stringa)
    const { id } = uuidParamSchema.parse(req.params);

    const data = modifyUtenteSchema.parse(req.body);
    const utenteAggiornato = await req.tenantStore.utenti.update(id, data as ModifyUtente);
    
    res.status(200).json({
        message: 'Utente aggiornato con successo',
        data: utenteAggiornato,
    });
}));

/**
 * GET /api/utenti/richieste-reset
 * Richieste di reset password aperte dei collaboratori dell'istituto (solo admin)
 */
router.get('/richieste-reset', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    const richieste = await req.tenantStore.utenti.getRichiesteReset();
    res.status(200).json({ data: richieste });
}));

/**
 * POST /api/utenti/richieste-reset/:id/ignora
 * Chiude una richiesta di reset senza generare una password (solo admin)
 */
router.post('/richieste-reset/:id/ignora', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    const { id } = uuidParamSchema.parse(req.params);
    await req.tenantStore.utenti.ignoraRichiestaReset(id, req.user.userId);
    res.status(200).json({ message: 'Richiesta ignorata', id });
}));

/**
 * POST /api/utenti/:id/password-temporanea
 * Genera una nuova password temporanea per un'utenza dell'istituto (solo admin).
 * La password torna una sola volta; l'utente dovrà cambiarla al primo accesso.
 */
router.post('/:id/password-temporanea', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }
    const { id } = uuidParamSchema.parse(req.params);
    const { passwordTemporanea, scadeAt } = await req.tenantStore.utenti.generaPasswordTemporanea(id, req.user.userId);
    res.status(200).json({ passwordTemporanea, scadeAt });
}));

/**
 * DELETE /api/utenti/:id
 * Elimina un utente (solo admin)
 * Non permette di eliminare l'unico admin dell'istituto
 */
router.delete('/delete-utente/:id', asyncHandler(async (req: Request, res: Response) => {
    if (!req.tenantStore) {
        return res.status(500).json({ error: 'Store non inizializzato' });
    }

    // Valida l'ID come UUID (stringa)
    const { id } = uuidParamSchema.parse(req.params);
    const istitutoId = req.user.istitutoId;

    // Recupera l'utente per verificare se è admin
    const [utenteDaEliminare] = await db.select().from(utenti)
        .where(and(eq(utenti.id, id), eq(utenti.istitutoId, istitutoId)));
    
    if (!utenteDaEliminare) {
        const error = new Error('Utente non trovato') as ErrorWithStatus;
        error.status = 404;
        throw error;
    }

    // Se l'utente è admin, verifica che non sia l'unico admin dell'istituto
    if (utenteDaEliminare.ruolo === 'admin') {
        const [adminCount] = await db.select({ valore: sql<number>`count(*)::int` })
            .from(utenti)
            .where(and(
                eq(utenti.istitutoId, istitutoId),
                eq(utenti.ruolo, 'admin')
            ));
        
        const numeroAdmin = Number(adminCount?.valore ?? 0);
        
        if (numeroAdmin === 1) {
            const error = new Error('Non è possibile eliminare l\'unico admin dell\'istituto') as ErrorWithStatus;
            error.status = 400;
            throw error;
        }
    }

    const idEliminato = await req.tenantStore.utenti.delete(id);
    
    res.status(200).json({
        message: 'Utente eliminato con successo',
        id: idEliminato,
    });
}));

export default router;
