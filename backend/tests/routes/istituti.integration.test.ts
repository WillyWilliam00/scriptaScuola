import { describe, it, expect, beforeEach, afterAll } from 'vitest';
import request from 'supertest';
import { eq } from 'drizzle-orm';
import { app } from '../../src/index.js';
import { cleanTestDb, testDb } from '../db/test-setup.js';
import { seedTestData, testCredentials } from '../db/test-seed.js';
import { loginAsAdmin, loginAsCollaboratore } from '../utils/test-helpers.js';
import { istituti } from '../../src/db/schema.js';

const NOME_ISTITUTO = 'Scuola Test';

async function istitutoEsiste(id: number) {
  const righe = await testDb.select({ id: istituti.id }).from(istituti).where(eq(istituti.id, id));
  return righe.length === 1;
}

describe('DELETE /api/istituti/delete-istituto', () => {
  let istitutoId: number;

  beforeEach(async () => {
    await cleanTestDb();
    const seed = await seedTestData();
    istitutoId = seed.istituto.id;
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  const elimina = (token: string, body?: object) => {
    const req = request(app).delete('/api/istituti/delete-istituto').set('Authorization', `Bearer ${token}`);
    return body ? req.send(body) : req;
  };

  it('senza conferma restituisce 400 e non elimina', async () => {
    const { token } = await loginAsAdmin();
    const res = await elimina(token);
    expect(res.status).toBe(400);
    expect(await istitutoEsiste(istitutoId)).toBe(true);
  });

  it('con nome istituto sbagliato restituisce 400 e non elimina', async () => {
    const { token } = await loginAsAdmin();
    const res = await elimina(token, { nomeIstituto: 'Altra Scuola', password: testCredentials.admin.password });
    expect(res.status).toBe(400);
    expect(await istitutoEsiste(istitutoId)).toBe(true);
  });

  it('senza password restituisce 401 e non elimina', async () => {
    const { token } = await loginAsAdmin();
    const res = await elimina(token, { nomeIstituto: NOME_ISTITUTO });
    expect(res.status).toBe(401);
    expect(await istitutoEsiste(istitutoId)).toBe(true);
  });

  it('con password errata restituisce 401 e non elimina', async () => {
    const { token } = await loginAsAdmin();
    const res = await elimina(token, { nomeIstituto: NOME_ISTITUTO, password: 'sbagliata' });
    expect(res.status).toBe(401);
    expect(await istitutoEsiste(istitutoId)).toBe(true);
  });

  it('un collaboratore riceve 403', async () => {
    const { token } = await loginAsCollaboratore();
    const res = await elimina(token, { nomeIstituto: NOME_ISTITUTO, password: testCredentials.collaboratore.password });
    expect(res.status).toBe(403);
    expect(await istitutoEsiste(istitutoId)).toBe(true);
  });

  it('con nome e password corretti elimina l\'istituto', async () => {
    const { token } = await loginAsAdmin();
    const res = await elimina(token, { nomeIstituto: `  ${NOME_ISTITUTO} `, password: testCredentials.admin.password });
    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('id', istitutoId);
    expect(await istitutoEsiste(istitutoId)).toBe(false);
  });
});
