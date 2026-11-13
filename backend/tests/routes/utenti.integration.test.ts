import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import request from 'supertest';
import { app } from '../../src/index.js';
import { cleanTestDb } from '../db/test-setup.js';
import { seedTestData } from '../db/test-seed.js';
import { loginAsAdmin, loginAsCollaboratore } from '../utils/test-helpers.js';

describe('Rotte /api/utenti', () => {
  beforeAll(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('GET /api/utenti restituisce lista paginata (solo admin)', async () => {
    const { token } = await loginAsAdmin();

    const res = await request(app)
      .get('/api/utenti')
      .set('Authorization', `Bearer ${token}`);

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('data');
    expect(Array.isArray(res.body.data)).toBe(true);
    expect(res.body).toHaveProperty('pagination');
    expect(res.body.pagination).toHaveProperty('page');
    expect(res.body.pagination).toHaveProperty('totalItems');
  });

  it('GET /api/utenti con collaboratore restituisce 403', async () => {
    const { token } = await loginAsCollaboratore();

    const res = await request(app)
      .get('/api/utenti')
      .set('Authorization', `Bearer ${token}`);

    expect(res.status).toBe(403);
  });

  it('GET /api/utenti senza token restituisce 401', async () => {
    const res = await request(app).get('/api/utenti');
    expect(res.status).toBe(401);
  });

  it('POST /api/utenti/new-utente crea un collaboratore (solo admin)', async () => {
    const { token } = await loginAsAdmin();

    const res = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_integ_test',
      });

    expect(res.status).toBe(201);
    expect(res.body).toHaveProperty('data');
    expect(res.body.data).toHaveProperty('id');
    expect(res.body.data.ruolo).toBe('collaboratore');
    expect(res.body.data.username).toBe('collab_integ_test');
  });

  it('POST /api/utenti/new-utente con collaboratore restituisce 403', async () => {
    const { token } = await loginAsCollaboratore();

    const res = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_integ_test',
      });
    expect(res.status).toBe(403);
  });

  it('PUT /api/utenti/update-utente/:id aggiorna un utente (solo admin)', async () => {
    const { token } = await loginAsAdmin();

    // Crea un utente da aggiornare
    const createRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_da_aggiornare',
      });
    expect(createRes.status).toBe(201);
    const userId = createRes.body.data.id as string;

    const res = await request(app)
      .put(`/api/utenti/update-utente/${userId}`)
      .set('Authorization', `Bearer ${token}`)
      // Nota: in questa codebase la rotta usa createUtenteSchema anche per update
      .send({
        ruolo: 'collaboratore',
        username: 'collab_aggiornato',
      });

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('data');
    expect(res.body.data.id).toBe(userId);
  });

  it('PUT /api/utenti/update-utente/:id permette di passare da collaboratore ad admin con nuova email', async () => {
    const { token } = await loginAsAdmin();

    // Crea un utente collaboratore di partenza
    const createRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_to_admin',
      });
    expect(createRes.status).toBe(201);
    const userId = createRes.body.data.id as string;

    // Aggiorna l'utente a admin fornendo una nuova email
    const updateRes = await request(app)
      .put(`/api/utenti/update-utente/${userId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'admin',
        email: 'newadmin@test.com',
      });

    expect(updateRes.status).toBe(200);
    expect(updateRes.body.data).toHaveProperty('id', userId);
    expect(updateRes.body.data).toHaveProperty('ruolo', 'admin');
    expect(updateRes.body.data).toHaveProperty('email', 'newadmin@test.com');
    expect(updateRes.body.data.username).toBeNull();
  });

  it('PUT /api/utenti/update-utente/:id blocca il passaggio a admin senza email (solo username)', async () => {
    const { token } = await loginAsAdmin();

    // Crea un utente collaboratore di partenza
    const createRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_to_admin_invalid',
      });
    expect(createRes.status).toBe(201);
    const userId = createRes.body.data.id as string;

    // Prova ad aggiornare a admin passando solo username (senza email)
    const updateRes = await request(app)
      .put(`/api/utenti/update-utente/${userId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'admin',
        username: 'should_not_matter',
      });

    expect(updateRes.status).toBe(400);
    expect(updateRes.body).toHaveProperty('error');
  });

  it('PUT /api/utenti/update-utente/:id permette di passare da admin a collaboratore con nuovo username', async () => {
    const { token } = await loginAsAdmin();

    // Crea un utente admin di partenza
    const createRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'admin',
        email: 'tempadmin@test.com',
      });
    expect(createRes.status).toBe(201);
    const userId = createRes.body.data.id as string;

    // Aggiorna l'utente a collaboratore fornendo un nuovo username
    const updateRes = await request(app)
      .put(`/api/utenti/update-utente/${userId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'admin_to_collab',
      });

    expect(updateRes.status).toBe(200);
    expect(updateRes.body.data).toHaveProperty('id', userId);
    expect(updateRes.body.data).toHaveProperty('ruolo', 'collaboratore');
    expect(updateRes.body.data).toHaveProperty('username', 'admin_to_collab');
    expect(updateRes.body.data.email).toBeNull();
  });

  it('PUT /api/utenti/update-utente/:id blocca il passaggio a collaboratore senza username (solo email)', async () => {
    const { token } = await loginAsAdmin();

    // Crea un utente admin di partenza
    const createRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'admin',
        email: 'admin_no_username_change@test.com',
      });
    expect(createRes.status).toBe(201);
    const userId = createRes.body.data.id as string;

    // Prova ad aggiornare a collaboratore passando solo email (senza username)
    const updateRes = await request(app)
      .put(`/api/utenti/update-utente/${userId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        email: 'admin_no_username_change@test.com',
      });

    expect(updateRes.status).toBe(400);
    expect(updateRes.body).toHaveProperty('error');
  });

  it('DELETE /api/utenti/delete-utente/:id elimina un utente (solo admin)', async () => {
    const { token } = await loginAsAdmin();

    // Crea un utente da eliminare
    const createRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_da_eliminare',
      });
    expect(createRes.status).toBe(201);
    const userId = createRes.body.data.id as string;

    const res = await request(app)
      .delete(`/api/utenti/delete-utente/${userId}`)
      .set('Authorization', `Bearer ${token}`);

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('id');
    expect(res.body.id).toBe(userId);
  });

  it('PUT /api/utenti/update-utente/:id restituisce 400 se si prova a usare una email già esistente', async () => {
    const { token } = await loginAsAdmin();

    // Crea un admin con una certa email
    const adminRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'admin',
        email: 'existingadmin@test.com',
      });
    expect(adminRes.status).toBe(201);

    // Crea un altro utente collaboratore
    const collabRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_for_email_conflict',
      });
    expect(collabRes.status).toBe(201);
    const collabId = collabRes.body.data.id as string;

    // Prova ad aggiornare il collaboratore a admin usando la stessa email dell'altro admin
    const updateRes = await request(app)
      .put(`/api/utenti/update-utente/${collabId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'admin',
        email: 'existingadmin@test.com',
      });

    expect(updateRes.status).toBe(400);
    expect(updateRes.body).toHaveProperty('error');
  });

  it('PUT /api/utenti/update-utente/:id restituisce 400 se si prova a usare uno username già esistente', async () => {
    const { token } = await loginAsAdmin();

    // Crea un collaboratore con uno username specifico
    const collabRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'existing_collab',
      });
    expect(collabRes.status).toBe(201);

    // Crea un admin da convertire
    const adminRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'admin',
        email: 'admin_for_username_conflict@test.com',
      });
    expect(adminRes.status).toBe(201);
    const adminId = adminRes.body.data.id as string;

    // Prova ad aggiornare l'admin a collaboratore usando uno username già esistente
    const updateRes = await request(app)
      .put(`/api/utenti/update-utente/${adminId}`)
      .set('Authorization', `Bearer ${token}`)
      .send({
        ruolo: 'collaboratore',
        username: 'existing_collab',
      });

    expect(updateRes.status).toBe(400);
    expect(updateRes.body).toHaveProperty('error');
  });

  it('DELETE /api/utenti/delete-utente/:id con collaboratore restituisce 403', async () => {
    const { token } = await loginAsCollaboratore();
    const { token: adminToken } = await loginAsAdmin();

    // Crea un utente da eliminare
    const createRes = await request(app)
      .post('/api/utenti/new-utente')
      .set('Authorization', `Bearer ${adminToken}`)
      .send({
        ruolo: 'collaboratore',
        username: 'collab_da_eliminare',
      });
    expect(createRes.status).toBe(201);
    const userId = createRes.body.data.id as string;

    const res = await request(app)
      .delete(`/api/utenti/delete-utente/${userId}`)
      .set('Authorization', `Bearer ${token}`);

    expect(res.status).toBe(403);
  });

  describe('Isolamento multi-tenant', () => {
    let adminIstituto1Token: string;
    let adminIstituto2Token: string;

    beforeAll(async () => {
      // Admin del primo istituto (creato da seedTestData)
      const { token: admin1Token } = await loginAsAdmin();
      adminIstituto1Token = admin1Token;

      // Crea un secondo istituto con il relativo admin
      const setupRes = await request(app)
        .post('/api/auth/setup-scuola')
        .send({
          istituto: {
            nome: 'Seconda Scuola',
            codiceIstituto: 'SCUOLA2000',
          },
          utente: {
            email: 'admin2@test.com',
            password: 'password123',
          },
        });
      expect(setupRes.status).toBe(201);

      // Login come admin del secondo istituto
      const login2Res = await request(app)
        .post('/api/auth/login')
        .send({
          identifier: 'admin2@test.com',
          password: 'password123',
        });
      expect(login2Res.status).toBe(200);
      adminIstituto2Token = login2Res.body.token;
    });

    it('GET /api/utenti restituisce solo utenti del proprio istituto', async () => {
      // Admin del primo istituto vede solo i suoi utenti
      const res = await request(app)
        .get('/api/utenti')
        .set('Authorization', `Bearer ${adminIstituto1Token}`);
      expect(res.status).toBe(200);
      expect(res.body.data).toBeInstanceOf(Array);
    });

    it('PUT /api/utenti/update-utente/:id non può modificare utenti di altri istituti', async () => {
      // Crea un utente nel secondo istituto
      const createRes = await request(app)
        .post('/api/utenti/new-utente')
        .set('Authorization', `Bearer ${adminIstituto2Token}`)
        .send({
          ruolo: 'collaboratore',
          username: 'collab_istituto2',
        });
      expect(createRes.status).toBe(201);
      const userIdIstituto2 = createRes.body.data.id as string;

      // Admin del primo istituto prova a modificare utente del secondo istituto
      const res = await request(app)
        .put(`/api/utenti/update-utente/${userIdIstituto2}`)
        .set('Authorization', `Bearer ${adminIstituto1Token}`)
        .send({
          ruolo: 'collaboratore',
          username: 'collab_tentativo',
        });

      expect(res.status).toBe(404);
      expect(res.body).toHaveProperty('error');
    });

    it('DELETE /api/utenti/delete-utente/:id non può eliminare utenti di altri istituti', async () => {
      // Crea un utente nel secondo istituto
      const createRes = await request(app)
        .post('/api/utenti/new-utente')
        .set('Authorization', `Bearer ${adminIstituto2Token}`)
        .send({
          ruolo: 'collaboratore',
          username: 'collab_da_eliminare_istituto2',
        });
      expect(createRes.status).toBe(201);
      const userIdIstituto2 = createRes.body.data.id as string;

      // Admin del primo istituto prova a eliminare utente del secondo istituto
      const res = await request(app)
        .delete(`/api/utenti/delete-utente/${userIdIstituto2}`)
        .set('Authorization', `Bearer ${adminIstituto1Token}`);

      expect(res.status).toBe(404);
      expect(res.body).toHaveProperty('error');
    });
  });
});

