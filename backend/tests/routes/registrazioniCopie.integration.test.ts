import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import request from 'supertest';
import { app } from '../../src/index.js';
import { cleanTestDb } from '../db/test-setup.js';
import { seedTestData } from '../db/test-seed.js';
import { loginAsAdmin, loginAsCollaboratore } from '../utils/test-helpers.js';
import { testDb } from '../db/test-setup.js';
import { docenti, istituti } from '../../src/db/schema.js';
import { eq } from 'drizzle-orm';

/** Crea un docente nell'istituto di test con il limite indicato, per test indipendenti dai dati del seed. */
async function creaDocente(limiteCopie: number) {
  const [istituto] = await testDb.select().from(istituti).where(eq(istituti.codiceIstituto, 'TEST123456'));
  const [docente] = await testDb.insert(docenti).values({
    nome: 'Limite',
    cognome: `Test${Math.random().toString(36).slice(2, 8)}`,
    limiteCopie,
    istitutoId: istituto!.id,
  }).returning();
  return docente!;
}

async function registra(token: string, docenteId: number, copieEffettuate: number) {
  return request(app)
    .post('/api/registrazioni-copie/new-registrazione')
    .set('Authorization', `Bearer ${token}`)
    .send({ docenteId, copieEffettuate });
}

async function modifica(token: string, id: number, copieEffettuate: number) {
  return request(app)
    .put(`/api/registrazioni-copie/update-registrazione/${id}`)
    .set('Authorization', `Bearer ${token}`)
    .send({ copieEffettuate });
}

describe('Rotte /api/registrazioni-copie', () => {
  beforeAll(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('GET /api/registrazioni-copie restituisce lista paginata (admin)', async () => {
    const { token } = await loginAsAdmin();

    const res = await request(app)
      .get('/api/registrazioni-copie')
      .set('Authorization', `Bearer ${token}`);

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('data');
    expect(Array.isArray(res.body.data)).toBe(true);
    expect(res.body.pagination).toHaveProperty('page');
    expect(res.body.pagination).toHaveProperty('totalItems');
  });

  it('GET /api/registrazioni-copie con collaboratore restituisce 403 (se route è admin-only)', async () => {
    const { token } = await loginAsCollaboratore();

    const res = await request(app)
      .get('/api/registrazioni-copie')
      .set('Authorization', `Bearer ${token}`);

    expect(res.status).toBe(403);
  });

  it('GET /api/registrazioni-copie senza token restituisce 401', async () => {
    const res = await request(app).get('/api/registrazioni-copie');
    expect(res.status).toBe(401);
  });

  it('POST /api/registrazioni-copie/new-registrazione crea registrazione con admin', async () => {
    const { token: adminToken, utente: adminUtente } = await loginAsAdmin();

    // Ottieni un docente esistente dalla lista
    const docentiRes = await request(app)
      .get('/api/docenti')
      .set('Authorization', `Bearer ${adminToken}`);
    expect(docentiRes.status).toBe(200);
    expect(docentiRes.body.data.length).toBeGreaterThan(0);
    const docenteId = docentiRes.body.data[0].id;

    const res = await request(app)
      .post('/api/registrazioni-copie/new-registrazione')
      .set('Authorization', `Bearer ${adminToken}`)
      .send({
        docenteId,
        copieEffettuate: 15,
        utenteId: adminUtente.id,
        note: 'Test registrazione admin',
      });

    expect(res.status).toBe(201);
    expect(res.body).toHaveProperty('id');
    expect(res.body).toHaveProperty('docenteId', docenteId);
    expect(res.body).toHaveProperty('copieEffettuate', 15);
    expect(res.body).toHaveProperty('utenteId', adminUtente.id);
  });

  it('POST /api/registrazioni-copie/new-registrazione crea registrazione con collaboratore', async () => {
    const { token: collaboratoreToken, utente: collaboratoreUtente } = await loginAsCollaboratore();

    // Ottieni un docente esistente dalla lista (usando token admin per vedere la lista)
    const { token: adminToken } = await loginAsAdmin();
    const docentiRes = await request(app)
      .get('/api/docenti')
      .set('Authorization', `Bearer ${adminToken}`);
    expect(docentiRes.status).toBe(200);
    expect(docentiRes.body.data.length).toBeGreaterThan(0);
    const docenteId = docentiRes.body.data[0].id;

    const res = await request(app)
      .post('/api/registrazioni-copie/new-registrazione')
      .set('Authorization', `Bearer ${collaboratoreToken}`)
      .send({
        docenteId,
        copieEffettuate: 20,
        utenteId: collaboratoreUtente.id,
        note: 'Test registrazione collaboratore',
      });
      
    expect(res.status).toBe(201);
    expect(res.body).toHaveProperty('id');
    expect(res.body).toHaveProperty('docenteId', docenteId);
    expect(res.body).toHaveProperty('copieEffettuate', 20);
    expect(res.body).toHaveProperty('utenteId', collaboratoreUtente.id);
  });

  it('POST /api/registrazioni-copie/new-registrazione senza token restituisce 401', async () => {
    const res = await request(app)
      .post('/api/registrazioni-copie/new-registrazione')
      .send({
        docenteId: 1,
        copieEffettuate: 10,
        utenteId: '123e4567-e89b-12d3-a456-426614174000',
      });

    expect(res.status).toBe(401);
  });

  it('DELETE /api/registrazioni-copie/delete-registrazione/:id elimina registrazione con admin', async () => {
    const { token: adminToken, utente: adminUtente } = await loginAsAdmin();

    // Ottieni un docente esistente per creare una registrazione valida
    const docentiRes = await request(app)
      .get('/api/docenti')
      .set('Authorization', `Bearer ${adminToken}`);
    expect(docentiRes.status).toBe(200);
    expect(docentiRes.body.data.length).toBeGreaterThan(0);
    const docenteId = docentiRes.body.data[0].id;

    const registrazioniRes = await request(app)
      .post('/api/registrazioni-copie/new-registrazione')
      .set('Authorization', `Bearer ${adminToken}`)
      .send({
        docenteId,
        copieEffettuate: 10,
        utenteId: adminUtente.id,
        note: 'Test registrazione admin',
      });

    expect(registrazioniRes.status).toBe(201);
    const registrazioneId = registrazioniRes.body.id;

    const deleteRes = await request(app)
      .delete(`/api/registrazioni-copie/delete-registrazione/${registrazioneId}`)
      .set('Authorization', `Bearer ${adminToken}`);
    expect(deleteRes.status).toBe(200);
    expect(deleteRes.body).toHaveProperty('id');
    expect(deleteRes.body.id).toBe(registrazioneId);
  });

  describe('attribuzione della registrazione (#25)', () => {
    it("salva sempre l'utente del token, anche se il body indica un altro utenteId", async () => {
      const { token: collaboratoreToken, utente: collaboratore } = await loginAsCollaboratore();
      const { utente: admin } = await loginAsAdmin();
      const docente = await creaDocente(100);

      const res = await request(app)
        .post('/api/registrazioni-copie/new-registrazione')
        .set('Authorization', `Bearer ${collaboratoreToken}`)
        .send({ docenteId: docente.id, copieEffettuate: 5, utenteId: admin.id });

      expect(res.status).toBe(201);
      expect(res.body).toHaveProperty('utenteId', collaboratore.id);
    });

    it('accetta una registrazione senza utenteId nel body', async () => {
      const { token, utente } = await loginAsCollaboratore();
      const docente = await creaDocente(100);

      const res = await registra(token, docente.id, 5);

      expect(res.status).toBe(201);
      expect(res.body).toHaveProperty('utenteId', utente.id);
    });
  });

  describe('limite copie in modifica (#21)', () => {
    it('rifiuta la modifica che supera il limite e accetta quelle che restano nel limite', async () => {
      const { token } = await loginAsAdmin();
      const docente = await creaDocente(10);
      const a = await registra(token, docente.id, 6);
      const b = await registra(token, docente.id, 3);
      expect(a.status).toBe(201);
      expect(b.status).toBe(201);

      // 6 + 5 = 11 > 10
      const oltre = await modifica(token, b.body.id, 5);
      expect(oltre.status).toBe(400);
      expect(oltre.body.error).toContain('ne ha già utilizzate 6');

      // 6 + 4 = 10: nel limite (la registrazione modificata non viene contata due volte)
      const alLimite = await modifica(token, b.body.id, 4);
      expect(alLimite.status).toBe(200);
      expect(alLimite.body.data).toHaveProperty('copieEffettuate', 4);

      // Riduzione: sempre accettata
      const riduzione = await modifica(token, a.body.id, 1);
      expect(riduzione.status).toBe(200);
      expect(riduzione.body.data).toHaveProperty('copieEffettuate', 1);
    });
  });

  describe('registrazioni simultanee (#22)', () => {
    it('registrazioni in parallelo sullo stesso docente non superano il limite', async () => {
      const { token: adminToken } = await loginAsAdmin();
      const { token: collaboratoreToken } = await loginAsCollaboratore();
      const docente = await creaDocente(10);

      // 6 richieste da 3 copie con limite 10: ne possono passare solo 3 (totale 9)
      const risposte = await Promise.all(
        Array.from({ length: 6 }, (_, i) => registra(i % 2 ? adminToken : collaboratoreToken, docente.id, 3)),
      );

      const accettate = risposte.filter((r) => r.status === 201).length;
      const rifiutate = risposte.filter((r) => r.status === 400).length;
      expect(accettate).toBe(3);
      expect(rifiutate).toBe(3);
    });
  });
});
