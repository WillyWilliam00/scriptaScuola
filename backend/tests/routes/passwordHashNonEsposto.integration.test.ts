import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import request from 'supertest';
import { app } from '../../src/index.js';
import { cleanTestDb } from '../db/test-setup.js';
import { seedTestData, testCredentials } from '../db/test-seed.js';
import { loginAsAdmin } from '../utils/test-helpers.js';

/**
 * Nessuna risposta dell'API deve contenere l'hash della password,
 * né come chiave (`passwordHash`, `password_hash`) né come valore (hash bcrypt `$2b$...`).
 */
function expectNessunHash(body: unknown) {
  const json = JSON.stringify(body);
  expect(json).not.toMatch(/password_?hash/i);
  expect(json).not.toMatch(/\$2[aby]\$\d{2}\$/);
}

describe('L\'hash della password non esce mai dalle risposte', () => {
  let token: string;
  let collaboratoreId: string;

  beforeAll(async () => {
    await cleanTestDb();
    const seed = await seedTestData();
    collaboratoreId = seed.collaboratore.id;
    token = (await loginAsAdmin()).token;
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('login', async () => {
    const res = await request(app).post('/api/auth/login').send({
      identifier: testCredentials.admin.email,
      password: testCredentials.admin.password,
    });
    expect(res.status).toBe(200);
    expectNessunHash(res.body);
  });

  it('registrazione della scuola', async () => {
    const res = await request(app).post('/api/auth/setup-scuola').send({
      istituto: { nome: 'Altra Scuola', codiceIstituto: 'ALTRA00001' },
      utente: { email: 'altro@admin.com', password: 'password123' },
    });
    expect(res.status).toBe(201);
    expectNessunHash(res.body);
  });

  it('elenco utenti', async () => {
    const res = await request(app).get('/api/utenti').set('Authorization', `Bearer ${token}`);
    expect(res.status).toBe(200);
    expectNessunHash(res.body);
  });

  it('creazione utente', async () => {
    const res = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({ ruolo: 'collaboratore', username: 'nuovo_collab', password: 'password123' });
    expect(res.status).toBe(201);
    expectNessunHash(res.body);
  });

  it('modifica utente senza nuova password', async () => {
    const res = await request(app)
      .put(`/api/utenti/update-utente/${collaboratoreId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({ ruolo: 'collaboratore', username: testCredentials.collaboratore.username });
    expect(res.status).toBe(200);
    expect(res.body.data).toMatchObject({
      id: collaboratoreId,
      username: testCredentials.collaboratore.username,
      ruolo: 'collaboratore',
    });
    expectNessunHash(res.body);
  });

  it('modifica utente con nuova password', async () => {
    const res = await request(app)
      .put(`/api/utenti/update-utente/${collaboratoreId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({ ruolo: 'collaboratore', username: testCredentials.collaboratore.username, password: 'passwordNuova1' });
    expect(res.status).toBe(200);
    expectNessunHash(res.body);
  });

  it('elenco registrazioni (contiene i dati dell\'utente che ha registrato)', async () => {
    const res = await request(app).get('/api/registrazioni-copie').set('Authorization', `Bearer ${token}`);
    expect(res.status).toBe(200);
    expectNessunHash(res.body);
  });

  it('eliminazione utente', async () => {
    const res = await request(app)
      .delete(`/api/utenti/delete-utente/${collaboratoreId}`)
      .set('Authorization', `Bearer ${token}`);
    expect(res.status).toBe(200);
    expectNessunHash(res.body);
  });
});
