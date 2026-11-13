import { and, asc, desc, eq, gt, ilike, isNull, ne, or, type AnyColumn } from "drizzle-orm";
import { bulkImportDocentiSchema, createUtenteSchema, docentiQuerySchema, insertDocenteSchema, insertRegistrazioneSchema, modifyDocenteSchema, modifyRegistrazioneSchema, modifyUtenteSchema, registrazioniCopieQuerySchema, utentiQuerySchema, type BulkImportDocenti, type CreateUtente, type DocentiQuery, type InsertDocente, type InsertRegistrazione, type ModifyDocente, type ModifyRegistrazione, type ModifyUtente, type RegistrazioniCopieQuery, type UtentiQuery } from "../../../shared/validation.js"
import { docenti, istituti, registrazioniCopie, richiesteResetPassword, utenti } from "./schema.js";
import type { DocentiPaginatedResponse, DocentiSort, UtentiPaginatedResponse, UtentiSort, Utente, RegistrazioniCopieSort, RegistrazioniCopiePaginatedResponse } from "../../../shared/types.js";
import { SQL, sql } from "drizzle-orm";
import type { ErrorWithStatus } from "../middleware/auth.js";
import bcrypt from 'bcrypt';
import type { NodePgDatabase } from "drizzle-orm/node-postgres";
import * as schema from "./schema.js";
import { buildPagination } from "./utils/pagination.js";
import { wouldExceedLimit } from "./utils/limiti.js";
import { stripUndefined } from "./utils/object.js";
import { getSort } from "./utils/sort.js";
import { isEmailIdentifier } from "./utils/auth.js";
import { revokeAllRefreshTokens } from "./utils/refreshTokens.js";
import { inizioFinestraRichieste, valoriPasswordTemporanea } from "./utils/passwordTemporanea.js";

// Tipo helper per estrarre il tipo della transazione dal database
type DbInstance = NodePgDatabase<typeof schema>;
type Transaction = Parameters<Parameters<DbInstance['transaction']>[0]>[0];
export const createTenantStore = (istitutoId: number, db: DbInstance | Transaction) => {

    const docentiStore = {
    
            getPaginated: async (query: DocentiQuery) => {

                const { page, pageSize, nome, cognome, sortField, sortOrder } = docentiQuerySchema.parse(query);
                const offset = (page - 1) * pageSize;



                const conditions = [eq(docenti.istitutoId, istitutoId)];
                if (nome) conditions.push(ilike(docenti.nome, `%${nome}%`));
                if (cognome) conditions.push(ilike(docenti.cognome, `%${cognome}%`));

                const copieSubquery = db
                    .select({
                        docenteId: registrazioniCopie.docenteId,
                        totaleCopie: sql<number>`SUM(${registrazioniCopie.copieEffettuate})::int`.as('totale_copie')
                    })
                    .from(registrazioniCopie)
                    .where(eq(registrazioniCopie.istitutoId, istitutoId))
                    .groupBy(registrazioniCopie.docenteId)
                    .as('copie_totali');

                const sortMap: Record<DocentiSort['field'], AnyColumn | SQL> = {
                    nome: docenti.nome,
                    cognome: docenti.cognome,
                    limiteCopie: docenti.limiteCopie,
                    createdAt: docenti.createdAt,
                    updatedAt: docenti.updatedAt,
                    // Aliased fields (totaleCopie) non sono visti da TS come SQL,
                    // quindi li castiamo esplicitamente a SQL per soddisfare il tipo.
                    copieEffettuate: copieSubquery.totaleCopie as unknown as SQL,
                    copieRimanenti: sql<number>`GREATEST(0, ${docenti.limiteCopie} - COALESCE(${copieSubquery.totaleCopie}, 0))::int` as unknown as SQL,
                }
                const sortKey = (sortField || 'nome') as DocentiSort['field'];
                const orderByColumn = sortMap[sortKey];
                const result = await db.select({
                    id: docenti.id,
                    nome: docenti.nome,
                    cognome: docenti.cognome,
                    limiteCopie: docenti.limiteCopie,
                    istitutoId: docenti.istitutoId,
                    createdAt: docenti.createdAt,
                    updatedAt: docenti.updatedAt,
                    copieEffettuate: sql<number>`COALESCE(${copieSubquery.totaleCopie}, 0)::int`.as('copie_effettuate'),
                    copieRimanenti: sql<number>`GREATEST(0, ${docenti.limiteCopie} - COALESCE(${copieSubquery.totaleCopie}, 0))::int`.as('copie_rimanenti')
                })
                    .from(docenti)
                    .leftJoin(copieSubquery, eq(docenti.id, copieSubquery.docenteId))
                    .where(and(...conditions))
                    .limit(pageSize)
                    .offset(offset)
                    .orderBy(...getSort(orderByColumn, sortOrder, docenti.id))
                
                const [totalCount] = await db.select({valore:sql`count(*)::int`}).from(docenti).where(and(...conditions));

                const totalItems = Number(totalCount?.valore ?? 0);
                const response: DocentiPaginatedResponse = {
                    data: result,
                    pagination: buildPagination(page, pageSize, totalItems),
                }

                return response;
            },
            getAllForExport: async (query: DocentiQuery) => {
                const { nome, cognome, sortField, sortOrder } = docentiQuerySchema.parse(query);

                const conditions = [eq(docenti.istitutoId, istitutoId)];
                if (nome) conditions.push(ilike(docenti.nome, `%${nome}%`));
                if (cognome) conditions.push(ilike(docenti.cognome, `%${cognome}%`));

                const copieSubquery = db
                    .select({
                        docenteId: registrazioniCopie.docenteId,
                        totaleCopie: sql<number>`SUM(${registrazioniCopie.copieEffettuate})::int`.as('totale_copie')
                    })
                    .from(registrazioniCopie)
                    .where(eq(registrazioniCopie.istitutoId, istitutoId))
                    .groupBy(registrazioniCopie.docenteId)
                    .as('copie_totali');

                const sortMap: Record<DocentiSort['field'], AnyColumn | SQL> = {
                    nome: docenti.nome,
                    cognome: docenti.cognome,
                    limiteCopie: docenti.limiteCopie,
                    createdAt: docenti.createdAt,
                    updatedAt: docenti.updatedAt,
                    copieEffettuate: copieSubquery.totaleCopie as unknown as SQL,
                    copieRimanenti: sql<number>`GREATEST(0, ${docenti.limiteCopie} - COALESCE(${copieSubquery.totaleCopie}, 0))::int` as unknown as SQL,
                };
                const sortKey = (sortField || 'nome') as DocentiSort['field'];
                const orderByColumn = sortMap[sortKey];

                const result = await db.select({
                    id: docenti.id,
                    nome: docenti.nome,
                    cognome: docenti.cognome,
                    limiteCopie: docenti.limiteCopie,
                    istitutoId: docenti.istitutoId,
                    createdAt: docenti.createdAt,
                    updatedAt: docenti.updatedAt,
                    copieEffettuate: sql<number>`COALESCE(${copieSubquery.totaleCopie}, 0)::int`.as('copie_effettuate'),
                    copieRimanenti: sql<number>`GREATEST(0, ${docenti.limiteCopie} - COALESCE(${copieSubquery.totaleCopie}, 0))::int`.as('copie_rimanenti')
                })
                    .from(docenti)
                    .leftJoin(copieSubquery, eq(docenti.id, copieSubquery.docenteId))
                    .where(and(...conditions))
                    .orderBy(...getSort(orderByColumn, sortOrder, docenti.id));

                return result;
            },
            create: async (data: Omit<InsertDocente, 'istitutoId'>) => {
                const { nome, cognome, limiteCopie } = insertDocenteSchema.parse({...data, istitutoId});
                const docenteEsistente = await db.select().from(docenti).where(and(eq(docenti.nome, nome), eq(docenti.cognome, cognome), eq(docenti.istitutoId, istitutoId)));
                if (docenteEsistente.length > 0) {
                    const error = new Error('Docente già esistente') as ErrorWithStatus;
                    error.status = 400;
                    throw error;
                }
                
                const [nuovoDocente] = await db
                .insert(docenti)
                .values({nome, cognome, limiteCopie, istitutoId})
                .returning()

                if (!nuovoDocente) {
                    const error = new Error('Errore nella creazione del docente') as ErrorWithStatus;
                    error.status = 500;
                    throw error;
                }
                return nuovoDocente
            },
            update: async(id: number, data: Omit<ModifyDocente, 'istitutoId'>) => {
                const validateData = modifyDocenteSchema.parse({...data});
                const updateValues: Partial<typeof docenti.$inferInsert> = {
                    updatedAt: new Date(),
                };
                Object.assign(updateValues, validateData);
                const cleanUpdateValues = stripUndefined(updateValues);
                const [docenteAggiornato] = await db
                .update(docenti)
                .set(cleanUpdateValues)
                .where(and(eq(docenti.id, id), eq(docenti.istitutoId, istitutoId)))
                .returning()
                if (!docenteAggiornato) {
                    const error = new Error('Docente non trovato') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                return docenteAggiornato
                
            },
            delete: async(id: number) => {
                const [docenteEliminato] = await db.delete(docenti)
                .where(and(eq(docenti.id, id), eq(docenti.istitutoId, istitutoId)))
                .returning({id: docenti.id})
                if (!docenteEliminato) {
                    const error = new Error('Docente non trovato') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                return docenteEliminato.id
            },
            deleteAll: async () => {
                const eliminati = await db.delete(docenti)
                    .where(eq(docenti.istitutoId, istitutoId))
                    .returning({ id: docenti.id });
                return eliminati.length;
            },
            checkLimiteCopie: async(docenteId: number, nuoveCopie: number, tx?: typeof db) => {
                const dbInstance = tx || db;
                const [stats] = await dbInstance
                .select({
                    limite: docenti.limiteCopie,
                    copieEffettuate: sql<number>`COALESCE(SUM(${registrazioniCopie.copieEffettuate}), 0)::int`
                })
                .from(docenti)
                .leftJoin(registrazioniCopie, and(eq(docenti.id, registrazioniCopie.docenteId), eq(registrazioniCopie.istitutoId, istitutoId)))
                .where(
                    and(
                        eq(docenti.id, docenteId),
                        eq(docenti.istitutoId, istitutoId),
                    )
                ).groupBy(docenti.id)
                if(!stats) {
                    const error = new Error('Docente non trovato') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                if (wouldExceedLimit(stats.copieEffettuate, nuoveCopie, stats.limite)) {
                    const error = new Error('Limite di copie superato') as ErrorWithStatus;
                    error.status = 400;
                    throw error;
                }
                return true;
            },
            bulkImportDocenti: async(data: BulkImportDocenti) => {
                const validateData = bulkImportDocentiSchema.parse(data);
                const { docenti: docentiDaImportare } = validateData;

                return await db.transaction(async (tx: any) => {
                    // 1. Recupera tutti i nomi/cognomi per un controllo massivo (evita N SELECT nel loop)
                    const docentiEsistenti = await tx
                        .select({ nome: docenti.nome, cognome: docenti.cognome })
                        .from(docenti)
                        .where(eq(docenti.istitutoId, istitutoId));

                    const existingSet = new Set(
                        docentiEsistenti.map((d: { nome: string; cognome: string }) => `${d.nome.toLowerCase()}-${d.cognome.toLowerCase()}`)
                    );

                    // 2. Valida e prepara i dati per l'inserimento batch
                    const docentiToInsert = [];
                    for (const d of docentiDaImportare) {
                        const key = `${d.nome.toLowerCase()}-${d.cognome.toLowerCase()}`;
                        if (existingSet.has(key)) {
                            const error = new Error(`Docente ${d.nome} ${d.cognome} già esistente`) as ErrorWithStatus;
                            error.status = 400;
                            throw error;
                        }
                        if (d.copieEffettuate > d.limiteCopie) {
                            const error = new Error(`Limiti superati per ${d.nome} ${d.cognome}`) as ErrorWithStatus;
                            error.status = 400;
                            throw error;
                        }
                        docentiToInsert.push({
                            nome: d.nome,
                            cognome: d.cognome,
                            limiteCopie: d.limiteCopie,
                            istitutoId,
                        });
                    }

                    // 3. Inserimento massivo docenti
                    const nuoviDocenti = await tx.insert(docenti).values(docentiToInsert).returning();

                    // 4. Registrazioni copie (ordine preservato: nuoviDocenti[i] corrisponde a docentiDaImportare[i])
                    const registrazioniToInsert: { docenteId: number; copieEffettuate: number; utenteId: null; note: string | null; istitutoId: number }[] = [];
                    nuoviDocenti.forEach((nuovo: typeof docenti.$inferSelect, index: number) => {
                        const dataOriginale = docentiDaImportare[index];
                        if (dataOriginale && dataOriginale.copieEffettuate > 0) {
                            registrazioniToInsert.push({
                                docenteId: nuovo.id,
                                copieEffettuate: dataOriginale.copieEffettuate,
                                utenteId: null,
                                note: dataOriginale.note ?? null,
                                istitutoId,
                            });
                        }
                    });

                    if (registrazioniToInsert.length > 0) {
                        await tx.insert(registrazioniCopie).values(registrazioniToInsert);
                    }

                    return {
                        docenti: nuoviDocenti,
                        totaleCreati: nuoviDocenti.length,
                        totaleConCopie: registrazioniToInsert.length,
                    };
                });
            }
        
    }

    return {
        docenti: docentiStore,
        utenti: {
            getPaginated: async(query: UtentiQuery, myUserId: string) => {
                const { page, pageSize, identifier, ruolo, sortField, sortOrder } = utentiQuerySchema.parse(query);
                const offset = (page - 1 ) * pageSize;
                const isEmail = identifier ? isEmailIdentifier(identifier) : false;

                const condition = [eq(utenti.istitutoId, istitutoId), ne(utenti.id, myUserId)];
                if (identifier) {
                    if (isEmail) condition.push(ilike(utenti.email, `%${identifier}%`));
                    else condition.push(ilike(utenti.username, `%${identifier}%`));
                }
                if (ruolo) condition.push(eq(utenti.ruolo, ruolo));

                const sortMap: Record<UtentiSort['field'], AnyColumn | SQL> = {
                    ruolo: utenti.ruolo,
                    identificativo: sql`COALESCE(${utenti.username}, ${utenti.email})`,
                }
                const sortKey = (sortField || 'ruolo') as UtentiSort['field'];
                const orderByColumn = sortMap[sortKey];
                const result = await db.select({
                    id: utenti.id,
                    username: utenti.username, 
                    email: utenti.email,
                    ruolo: utenti.ruolo,
                    createdAt: utenti.createdAt,
                    updatedAt: utenti.updatedAt,
                    istitutoId: utenti.istitutoId,
                })
                .from(utenti)
                .where(and(...condition))
                .limit(pageSize)
                .offset(offset)
                .orderBy(...getSort(orderByColumn, sortOrder, utenti.id))

                const [totalCount] = await db.select({valore: sql`count(*)::int`}).from(utenti).where(and(...condition));

                const totalItems = Number(totalCount?.valore ?? 0);
                const response: UtentiPaginatedResponse = {
                    // Type assertion sicuro: il constraint del database garantisce che
                    // admin ha email non-null e collaboratore ha username non-null
                    data: result as Utente[],
                    pagination: buildPagination(page, pageSize, totalItems),
                }
                return response;
            },
            getAllForExport: async(query: UtentiQuery, myUserId: string) => {
                const { identifier, ruolo, sortField, sortOrder } = utentiQuerySchema.parse(query);
                const isEmail = identifier ? isEmailIdentifier(identifier) : false;

                const condition = [eq(utenti.istitutoId, istitutoId), ne(utenti.id, myUserId)];
                if (identifier) {
                    if (isEmail) condition.push(ilike(utenti.email, `%${identifier}%`));
                    else condition.push(ilike(utenti.username, `%${identifier}%`));
                }
                if (ruolo) condition.push(eq(utenti.ruolo, ruolo));

                const sortMap: Record<UtentiSort['field'], AnyColumn | SQL> = {
                    ruolo: utenti.ruolo,
                    identificativo: sql`COALESCE(${utenti.username}, ${utenti.email})`,
                };
                const sortKey = (sortField || 'ruolo') as UtentiSort['field'];
                const orderByColumn = sortMap[sortKey];

                const result = await db.select({
                    id: utenti.id,
                    username: utenti.username, 
                    email: utenti.email,
                    ruolo: utenti.ruolo,
                    createdAt: utenti.createdAt,
                    updatedAt: utenti.updatedAt,
                    istitutoId: utenti.istitutoId,
                })
                .from(utenti)
                .where(and(...condition))
                .orderBy(...getSort(orderByColumn, sortOrder, utenti.id));

                return result as Utente[];
            },
            /**
             * Crea un'utenza. Senza `opzioni.password` la piattaforma genera una password temporanea
             * (da cambiare al primo accesso) e la restituisce una sola volta in `passwordTemporanea`.
             * La password scelta si usa solo alla registrazione dell'istituto.
             */
            create: async(data: Omit<CreateUtente, 'istitutoId'>, opzioni: { password?: string } = {}) => {
                const validateData = createUtenteSchema.parse(data);
                const isAdmin = validateData.ruolo === 'admin';
                const utenteEsistente = await db.select().from(utenti)
                .where(
                    isAdmin ?
                        eq(utenti.email, validateData.email) :
                        eq(utenti.username, validateData.username)
                )
                if(utenteEsistente.length > 0) {
                    const error = new Error('Utente già esistente') as ErrorWithStatus;
                    error.status = 400;
                    throw error;
                }
                const temporanea = opzioni.password ? null : await valoriPasswordTemporanea();
                const valoriPassword = temporanea
                    ? temporanea.valori
                    : { passwordHash: await bcrypt.hash(opzioni.password!, 10) };
                const [nuovoUtente] = await db.insert(utenti)
                .values({
                    email: isAdmin ? validateData.email : null,
                    username: isAdmin ? null : validateData.username,
                    ...valoriPassword,
                    ruolo: validateData.ruolo,
                    istitutoId,
                })
                .returning({
                    id: utenti.id,
                    email: utenti.email,
                    username: utenti.username,
                    ruolo: utenti.ruolo,
                    istitutoId: utenti.istitutoId,
                    createdAt: utenti.createdAt,
                    updatedAt: utenti.updatedAt,
                })
                if(!nuovoUtente) {
                    const error = new Error('Errore nella creazione dell\'utente') as ErrorWithStatus;
                    error.status = 500;
                    throw error;
                }
                return {
                    utente: nuovoUtente,
                    passwordTemporanea: temporanea?.passwordTemporanea ?? null,
                    scadeAt: temporanea?.valori.passwordTemporaneaScadeAt ?? null,
                };
            },
            update: async(id: string, data: Omit<ModifyUtente, 'istitutoId'>) => {
                const validateData = modifyUtenteSchema.parse(data);
                const updateValues: Partial<typeof utenti.$inferSelect> = {
                    updatedAt: new Date(),
                    ruolo: validateData.ruolo,
                }
                
                if(validateData.ruolo === 'admin') {
                    if(!validateData.email) {
                        const error = new Error('Email è obbligatoria per admin') as ErrorWithStatus;
                        error.status = 400;
                        throw error;
                    }
                    updateValues.username = null;
                    const [utenteEsistente] = await db.select().from(utenti).where(and(eq(utenti.email, validateData.email), ne(utenti.id, id)));
                    if(utenteEsistente) {
                        const error = new Error('Email già in uso') as ErrorWithStatus;
                        error.status = 400;
                        throw error;
                    }
                    updateValues.email = validateData.email;
                
                } else {
                    if(!validateData.username) {
                        const error = new Error('Username è obbligatorio per collaboratore') as ErrorWithStatus;
                        error.status = 400;
                        throw error;
                    }
                    updateValues.email = null;
                    const [utenteEsistente] = await db.select().from(utenti).where(and(eq(utenti.username, validateData.username), ne(utenti.id, id)));
                    if(utenteEsistente) {
                        const error = new Error('Username già in uso') as ErrorWithStatus;
                        error.status = 400;
                        throw error;
                    }
                    updateValues.username = validateData.username;
                }
                const {ruolo, ...rest} = validateData;
                Object.assign(updateValues, rest);
                const cleanUpdateValues = stripUndefined(updateValues);
                
                const [utenteAggiornato] = await db.update(utenti).set(
                    cleanUpdateValues
                )
                .where(and(eq(utenti.id, id), eq(utenti.istitutoId, istitutoId)))
                // Colonne esplicite: l'hash della password non deve mai uscire dallo store
                .returning({
                    id: utenti.id,
                    email: utenti.email,
                    username: utenti.username,
                    ruolo: utenti.ruolo,
                    istitutoId: utenti.istitutoId,
                    createdAt: utenti.createdAt,
                    updatedAt: utenti.updatedAt,
                })
                if(!utenteAggiornato) {
                    const error = new Error('Utente non trovato') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                return utenteAggiornato;
            },
            /**
             * Genera una nuova password temporanea per un'utenza dell'istituto (reset fatto da un admin).
             * Revoca le sessioni aperte e chiude le richieste di reset ancora aperte per quell'utenza.
             */
            generaPasswordTemporanea: async(id: string, adminId: string) => {
                if (id === adminId) {
                    const error = new Error('Per cambiare la tua password usa la pagina Profilo') as ErrorWithStatus;
                    error.status = 400;
                    throw error;
                }
                const { passwordTemporanea, valori } = await valoriPasswordTemporanea();
                const eseguiReset = async (tx: Transaction) => {
                    const [aggiornato] = await tx.update(utenti)
                        .set({ ...valori, updatedAt: new Date() })
                        .where(and(eq(utenti.id, id), eq(utenti.istitutoId, istitutoId)))
                        .returning({ id: utenti.id });
                    if (!aggiornato) {
                        const error = new Error('Utente non trovato') as ErrorWithStatus;
                        error.status = 404;
                        throw error;
                    }
                    await revokeAllRefreshTokens(tx, id);
                    await tx.update(richiesteResetPassword)
                        .set({ gestitaAt: new Date(), gestitaDa: adminId, esito: 'password_generata' })
                        .where(and(
                            eq(richiesteResetPassword.utenteId, id),
                            eq(richiesteResetPassword.istitutoId, istitutoId),
                            isNull(richiesteResetPassword.gestitaAt),
                        ));
                };
                // Dentro una transazione esistente Drizzle usa un savepoint
                await db.transaction(eseguiReset);
                return {
                    passwordTemporanea,
                    scadeAt: valori.passwordTemporaneaScadeAt,
                };
            },
            /** Richieste di reset aperte (non gestite, degli ultimi 7 giorni), dalla più recente. */
            getRichiesteReset: async() => {
                return db.select({
                    id: richiesteResetPassword.id,
                    utenteId: richiesteResetPassword.utenteId,
                    username: utenti.username,
                    email: utenti.email,
                    ruolo: utenti.ruolo,
                    nota: richiesteResetPassword.nota,
                    createdAt: richiesteResetPassword.createdAt,
                })
                .from(richiesteResetPassword)
                .innerJoin(utenti, eq(utenti.id, richiesteResetPassword.utenteId))
                .where(and(
                    eq(richiesteResetPassword.istitutoId, istitutoId),
                    isNull(richiesteResetPassword.gestitaAt),
                    gt(richiesteResetPassword.createdAt, inizioFinestraRichieste()),
                ))
                .orderBy(desc(richiesteResetPassword.createdAt));
            },
            /** Segna una richiesta di reset come ignorata. */
            ignoraRichiestaReset: async(richiestaId: string, adminId: string) => {
                const [richiesta] = await db.update(richiesteResetPassword)
                    .set({ gestitaAt: new Date(), gestitaDa: adminId, esito: 'ignorata' })
                    .where(and(
                        eq(richiesteResetPassword.id, richiestaId),
                        eq(richiesteResetPassword.istitutoId, istitutoId),
                        isNull(richiesteResetPassword.gestitaAt),
                    ))
                    .returning({ id: richiesteResetPassword.id });
                if (!richiesta) {
                    const error = new Error('Richiesta non trovata o già gestita') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                return richiesta.id;
            },
            delete: async(id: string) => {
                const [utenteEliminato] = await db.delete(utenti)
                .where(and(eq(utenti.id, id), eq(utenti.istitutoId, istitutoId)))
                .returning({id: utenti.id})
                if(!utenteEliminato) {
                    const error = new Error('Utente non trovato') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                return utenteEliminato.id;
            }
        },
        registrazioniCopie: {
            getPaginated: async(query: RegistrazioniCopieQuery) => {
                const { page, pageSize, docenteId, utenteId, docenteNome, docenteCognome, copieEffettuate: copieFilter, utenteIdentifier, sortField, sortOrder } = registrazioniCopieQuerySchema.parse(query);
                const offset = (page - 1) * pageSize;
                const condition = [eq(registrazioniCopie.istitutoId, istitutoId)];
                if(docenteId) condition.push(eq(registrazioniCopie.docenteId, docenteId));
                if(utenteId) condition.push(eq(registrazioniCopie.utenteId, utenteId));
                if(docenteNome) condition.push(ilike(docenti.nome, `%${docenteNome}%`));
                if(docenteCognome) condition.push(ilike(docenti.cognome, `%${docenteCognome}%`));
                if(copieFilter !== undefined) condition.push(eq(registrazioniCopie.copieEffettuate, copieFilter));
                if(utenteIdentifier) {
                    const toPush = or(ilike(utenti.username, `%${utenteIdentifier}%`), ilike(utenti.email, `%${utenteIdentifier}%`));
                    if(toPush) condition.push(toPush);
                }
                const sortMap: Record<RegistrazioniCopieSort['field'], AnyColumn | SQL> = {
                    docenteId: registrazioniCopie.docenteId,
                    utenteId: registrazioniCopie.utenteId,
                    createdAt: registrazioniCopie.createdAt,
                    updatedAt: registrazioniCopie.updatedAt,
                    docenteNome: docenti.nome,
                    docenteCognome: docenti.cognome,
                    copieEffettuate: registrazioniCopie.copieEffettuate,
                    utente: sql`COALESCE(${utenti.username}, ${utenti.email})`,
                }
                const sortKey = (sortField || 'createdAt') as RegistrazioniCopieSort['field'];
                const orderByColumn = sortMap[sortKey]; 
                
                const registrazioniPaginated = await db.select({
                    id: registrazioniCopie.id,
                    docenteId: registrazioniCopie.docenteId,
                    copieEffettuate: registrazioniCopie.copieEffettuate,
                    istitutoId: registrazioniCopie.istitutoId,
                    utenteId: registrazioniCopie.utenteId,
                    note: registrazioniCopie.note,
                    createdAt: registrazioniCopie.createdAt,
                    updatedAt: registrazioniCopie.updatedAt,
                    docenteNome: docenti.nome,
                    docenteCognome: docenti.cognome,
                    utenteUsername: utenti.username,
                    utenteEmail: utenti.email,
                })
                .from(registrazioniCopie)
                .leftJoin(docenti, eq(registrazioniCopie.docenteId, docenti.id))
                .leftJoin(utenti, eq(registrazioniCopie.utenteId, utenti.id))
                .where(and(...condition))
                .limit(pageSize)
                .offset(offset)
                .orderBy(...getSort(orderByColumn, sortOrder, registrazioniCopie.id))
             

                const [totalCount] = await db.select({valore: sql`count(*)::int`}).from(registrazioniCopie).leftJoin(docenti, eq(registrazioniCopie.docenteId, docenti.id)).leftJoin(utenti, eq(registrazioniCopie.utenteId, utenti.id)).where(and(...condition));

                const totalItems = Number(totalCount?.valore ?? 0);
                const response: RegistrazioniCopiePaginatedResponse = {
                    data: registrazioniPaginated,
                    pagination: buildPagination(page, pageSize, totalItems),
                }
                return response;

            },
            getAllForExport: async(query: RegistrazioniCopieQuery) => {
                const { docenteId, utenteId, docenteNome, docenteCognome, copieEffettuate: copieFilter, utenteIdentifier, sortField, sortOrder } = registrazioniCopieQuerySchema.parse(query);
                const condition = [eq(registrazioniCopie.istitutoId, istitutoId)];
                if(docenteId) condition.push(eq(registrazioniCopie.docenteId, docenteId));
                if(utenteId) condition.push(eq(registrazioniCopie.utenteId, utenteId));
                if(docenteNome) condition.push(ilike(docenti.nome, `%${docenteNome}%`));
                if(docenteCognome) condition.push(ilike(docenti.cognome, `%${docenteCognome}%`));
                if(copieFilter !== undefined) condition.push(eq(registrazioniCopie.copieEffettuate, copieFilter));
                if(utenteIdentifier) {
                    const toPush = or(ilike(utenti.username, `%${utenteIdentifier}%`), ilike(utenti.email, `%${utenteIdentifier}%`));
                    if(toPush) condition.push(toPush);
                }
                const sortMap: Record<RegistrazioniCopieSort['field'], AnyColumn | SQL> = {
                    docenteId: registrazioniCopie.docenteId,
                    utenteId: registrazioniCopie.utenteId,
                    createdAt: registrazioniCopie.createdAt,
                    updatedAt: registrazioniCopie.updatedAt,
                    docenteNome: docenti.nome,
                    docenteCognome: docenti.cognome,
                    copieEffettuate: registrazioniCopie.copieEffettuate,
                    utente: sql`COALESCE(${utenti.username}, ${utenti.email})`,
                }
                const sortKey = (sortField || 'createdAt') as RegistrazioniCopieSort['field'];
                const orderByColumn = sortMap[sortKey]; 
                
                const registrazioni = await db.select({
                    id: registrazioniCopie.id,
                    docenteId: registrazioniCopie.docenteId,
                    copieEffettuate: registrazioniCopie.copieEffettuate,
                    istitutoId: registrazioniCopie.istitutoId,
                    utenteId: registrazioniCopie.utenteId,
                    note: registrazioniCopie.note,
                    createdAt: registrazioniCopie.createdAt,
                    updatedAt: registrazioniCopie.updatedAt,
                    docenteNome: docenti.nome,
                    docenteCognome: docenti.cognome,
                    utenteUsername: utenti.username,
                    utenteEmail: utenti.email,
                })
                .from(registrazioniCopie)
                .leftJoin(docenti, eq(registrazioniCopie.docenteId, docenti.id))
                .leftJoin(utenti, eq(registrazioniCopie.utenteId, utenti.id))
                .where(and(...condition))
                .orderBy(...getSort(orderByColumn, sortOrder, registrazioniCopie.id));

                return registrazioni;
            },
            create: async(data: Omit<InsertRegistrazione, 'istitutoId'>) => {
                const validateData = insertRegistrazioneSchema.parse(data);
                const result = await db.transaction(async (tx: any) => {
                    const [docente] = await tx.select().from(docenti).where(and(eq(docenti.id, validateData.docenteId), eq(docenti.istitutoId, istitutoId)));
                    const [utente] = await tx.select().from(utenti).where(and(eq(utenti.id, validateData.utenteId), eq(utenti.istitutoId, istitutoId)));
                    if(!docente) {
                        const error = new Error('Docente non trovato') as ErrorWithStatus;
                        error.status = 404;
                        throw error;
                    }
                    if(!utente) {
                        const error = new Error('Utente non trovato') as ErrorWithStatus;
                        error.status = 404;
                        throw error;
                    }
                    await docentiStore.checkLimiteCopie(validateData.docenteId, validateData.copieEffettuate, tx);
                    
                    const [registrazione] = await tx.insert(registrazioniCopie).values({
                        docenteId: validateData.docenteId,
                        copieEffettuate: validateData.copieEffettuate,
                        utenteId: validateData.utenteId,
                        note: validateData.note,
                        istitutoId,
                    })
                    .returning();
                    return registrazione;
                });
                return result;
            },
            update: async(id: number, data: ModifyRegistrazione) => {
                const validateData = modifyRegistrazioneSchema.parse(data);
                const result = await db.transaction(async (tx: any) => {
                    // Verifica che la registrazione esista e appartenga all'istituto
                    const [registrazioneEsistente] = await tx.select()
                        .from(registrazioniCopie)
                        .where(and(
                            eq(registrazioniCopie.id, id),
                            eq(registrazioniCopie.istitutoId, istitutoId)
                        ));
                    
                    if(!registrazioneEsistente) {
                        const error = new Error('Registrazione non trovata') as ErrorWithStatus;
                        error.status = 404;
                        throw error;
                    }

                    // Ottieni il docente per verificare il limite
                    const [docente] = await tx.select()
                        .from(docenti)
                        .where(and(
                            eq(docenti.id, registrazioneEsistente.docenteId),
                            eq(docenti.istitutoId, istitutoId)
                        ));

                    if(!docente) {
                        const error = new Error('Docente non trovato') as ErrorWithStatus;
                        error.status = 404;
                        throw error;
                    }

                    // Calcola le copie totali del docente escludendo questa registrazione
                    const [totaleCopie] = await tx.select({
                        totale: sql<number>`COALESCE(SUM(${registrazioniCopie.copieEffettuate}), 0)::int`
                    })
                    .from(registrazioniCopie)
                    .where(and(
                        eq(registrazioniCopie.docenteId, docente.id),
                        eq(registrazioniCopie.istitutoId, istitutoId),
                    ));

                    const nuoveCopieTotali = totaleCopie + validateData.copieEffettuate;

                    // Verifica che non superi il limite
                    if(nuoveCopieTotali > docente.limiteCopie) {
                        const error = new Error(`Limite di copie superato. Il docente ha un limite di ${docente.limiteCopie} copie e ne ha già utilizzate ${totaleCopie}. Le nuove copie (${validateData.copieEffettuate}) porterebbero il totale a ${nuoveCopieTotali}.`) as ErrorWithStatus;
                        error.status = 400;
                        throw error;
                    }

                    // Aggiorna la registrazione
                    const [registrazioneAggiornata] = await tx.update(registrazioniCopie)
                        .set({
                            copieEffettuate: validateData.copieEffettuate,
                            note: validateData.note ?? null,
                            updatedAt: new Date(),
                        })
                        .where(and(
                            eq(registrazioniCopie.id, id),
                            eq(registrazioniCopie.istitutoId, istitutoId)
                        ))
                        .returning();

                    return registrazioneAggiornata;
                });
                return result;
            },
            delete: async(id: number) => {
                const [registrazioneEliminata] = await db.delete(registrazioniCopie)
                .where(and(eq(registrazioniCopie.id, id), eq(registrazioniCopie.istitutoId, istitutoId)))
                .returning({id: registrazioniCopie.id})
                if(!registrazioneEliminata) {
                    const error = new Error('Registrazione non trovata') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                return registrazioneEliminata;
            },
            deleteAll: async() => {
                await db.delete(registrazioniCopie).where(eq(registrazioniCopie.istitutoId, istitutoId));
            }
        },
        istituti: {
            // Elimina l'istituto dell'utente autenticato (istitutoId dalla closure = token)
            delete: async() => {
                const [istitutoEliminato] = await db.delete(istituti).where(eq(istituti.id, istitutoId)).returning({id: istituti.id});
                if(!istitutoEliminato) {
                    const error = new Error('Istituto non trovato') as ErrorWithStatus;
                    error.status = 404;
                    throw error;
                }
                return {
                    message: 'Istituto eliminato con successo',
                    id: istitutoEliminato.id,
                }
            }
        }
    }
}
