import 'dotenv/config';
import { drizzle } from 'drizzle-orm/node-postgres';
import { Pool } from 'pg';
import { sql } from 'drizzle-orm';
import * as schema from '../../src/db/schema.js';

/**
 * Crea una connessione al database di test
 * Usa DATABASE_URL_TEST se disponibile, altrimenti DATABASE_URL
 * 
 * Creiamo il Pool manualmente così possiamo chiuderlo quando necessario
 */
const testDbUrl = process.env.DATABASE_URL_TEST || process.env.DATABASE_URL!;
const pool = new Pool({ connectionString: testDbUrl });
export const testDb = drizzle(pool, { schema });

/**
 * Pulisce tutte le tabelle del database di test
 * Utile per resettare il DB prima/dopo i test
 * 
 * Usa TRUNCATE CASCADE per eliminare tutti i dati rispettando le foreign key:
 * - TRUNCATE elimina tutti i dati dalle tabelle
 * - CASCADE elimina automaticamente i dati dalle tabelle dipendenti
 * - RESTART IDENTITY resetta i contatori degli ID auto-incrementali
 */
export async function cleanTestDb() {
  // TRUNCATE CASCADE elimina tutti i dati rispettando le foreign key
  // L'ordine delle tabelle non è critico con CASCADE, ma le elenchiamo per chiarezza
  await testDb.execute(sql`
    TRUNCATE TABLE 
      refresh_tokens,
      richieste_reset_password,
      token_email,
      registrazioni_copie,
      docenti,
      utenti,
      istituti
    RESTART IDENTITY CASCADE
  `);
}

/**
 * Chiude la connessione al database di test
 * 
 * IMPORTANTE: Chiama questa funzione dopo tutti i test per evitare che il processo Node.js
 * rimanga in attesa di connessioni aperte. Senza questo, i test potrebbero non terminare.
 */
export async function closeTestDb() {
  await pool.end();
}

