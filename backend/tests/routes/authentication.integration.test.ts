import { describe, it, expect, beforeAll, afterAll, beforeEach } from 'vitest';
import request from 'supertest';
import { app } from '../../src/index.js';
import { cleanTestDb, testDb } from '../db/test-setup.js';
import { seedTestData, testCredentials } from '../db/test-seed.js';
import { refreshTokens } from '../../src/db/schema.js';
import { eq } from 'drizzle-orm';
import { loginAsAdmin, loginAsCollaboratore } from '../utils/test-helpers.js';

describe('POST /api/auth/login', () => {
  beforeAll(async () => {
    // Pulisci il DB prima di iniziare i test
    await cleanTestDb();
    // Popola con dati di test
    await seedTestData();
  });

  afterAll(async () => {
    // Pulisci il DB dopo tutti i test
    await cleanTestDb();
  });

  it('restituisce 200 e token con credenziali admin valide', async () => {
    const res = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        password: testCredentials.admin.password,
      });

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('token');
    expect(res.body).toHaveProperty('refreshToken');
    expect(res.body).toHaveProperty('utente');
    expect(res.body.utente.email).toBe(testCredentials.admin.email);
    expect(res.body.utente.ruolo).toBe('admin');
  });

  it('restituisce 200 e token con credenziali collaboratore valide', async () => {
    const res = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.collaboratore.username,
        password: testCredentials.collaboratore.password,
      });

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('token');
    expect(res.body).toHaveProperty('refreshToken');
    expect(res.body).toHaveProperty('utente');
    expect(res.body.utente.username).toBe(testCredentials.collaboratore.username);
    expect(res.body.utente.ruolo).toBe('collaboratore');
  });

  it('restituisce 401 con credenziali non valide', async () => {
    const res = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        password: 'password_sbagliata',
      });

    expect(res.status).toBe(401);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 401 con credenziali non valide', async () => {
    const res = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: 'inesistente@test.com',
        password: 'password123',
      });

    expect(res.status).toBe(401);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 400 con dati mancanti', async () => {
    const res = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        // password mancante
      });

    expect(res.status).toBe(400);
  });
});

describe('POST /api/auth/setup-scuola', () => {
  beforeEach(async () => {
    // Pulisci il DB prima di ogni test (setup-scuola richiede DB vuoto)
    await cleanTestDb();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('crea istituto e utente admin con successo', async () => {
    const res = await request(app)
      .post('/api/auth/setup-scuola')
      .send({
        istituto: {
          nome: 'Nuova Scuola',
          // deve essere lungo ESATTAMENTE 10 caratteri (schema Zod)
          codiceIstituto: 'NUOVO12345',
        },
        utente: {
          email: 'nuovo@admin.com',
          password: 'password123',
        },
      });

    expect(res.status).toBe(201);
    expect(res.body).toHaveProperty('message');
    expect(res.body).toHaveProperty('istituto');
    expect(res.body).toHaveProperty('utente');
    expect(res.body.istituto.nome).toBe('Nuova Scuola');
    expect(res.body.istituto.codiceIstituto).toBe('NUOVO12345');
    expect(res.body.utente.email).toBe('nuovo@admin.com');
    expect(res.body.utente.ruolo).toBe('admin');
  });

  it('restituisce 400 se utente già esistente', async () => {
    // Prima crea un utente
    await seedTestData();

    // Poi prova a crearne uno con la stessa email
    const res = await request(app)
      .post('/api/auth/setup-scuola')
      .send({
        istituto: {
          nome: 'Altra Scuola',
          // usa un valore valido per non fallire prima in validazione
          codiceIstituto: 'ALTRA45678',
        },
        utente: {
          email: testCredentials.admin.email, // Email già esistente
          password: 'password123',
        },
      });

    expect(res.status).toBe(400);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 400 se istituto già esistente', async () => {
    // Prima crea un istituto
    await seedTestData();

    // Poi prova a crearne uno con lo stesso codice
    const res = await request(app)
      .post('/api/auth/setup-scuola')
      .send({
        istituto: {
          nome: 'Altra Scuola',
          codiceIstituto: 'TEST123456', // Codice già esistente
        },
        utente: {
          email: 'altro@admin.com',
          password: 'password123',
        },
      });

    expect(res.status).toBe(400);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 400 con dati mancanti', async () => {
    const res = await request(app)
      .post('/api/auth/setup-scuola')
      .send({
        istituto: {
          nome: 'Scuola',
          // codiceIstituto mancante
        },
        utente: {
          email: 'test@test.com',
          password: 'password123',
        },
      });

    expect(res.status).toBe(400);
  });
});

describe('POST /api/auth/logout', () => {
  beforeAll(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('revoca il refresh token con successo', async () => {
    // Ottieni un refresh token fresco facendo login
    const loginRes = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        password: testCredentials.admin.password,
      });
    
    const refreshToken = loginRes.body.refreshToken;

    // Verifica che il token esista e non sia ancora revocato
    const [tokenBefore] = await testDb
      .select()
      .from(refreshTokens)
      .where(eq(refreshTokens.token, refreshToken));
    
    expect(tokenBefore).toBeDefined();
    expect(tokenBefore?.revokedAt).toBeNull(); // Non ancora revocato

    // Fai logout
    const res = await request(app)
      .post('/api/auth/logout')
      .send({
        refreshToken,
      });

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('message');
    expect(res.body.message).toBe('Logout effettuato con successo');

    // Verifica che il token sia stato effettivamente revocato nel DB
    const [tokenAfter] = await testDb
      .select()
      .from(refreshTokens)
      .where(eq(refreshTokens.token, refreshToken));
    
    expect(tokenAfter).toBeDefined();
    expect(tokenAfter?.revokedAt).not.toBeNull(); // Ora è revocato!
    expect(tokenAfter?.revokedAt).toBeInstanceOf(Date);
    
    // Verifica che la data di revoca sia recente (impostata durante il test)
    const revokedAtDate = tokenAfter!.revokedAt!;
    const now = new Date();
    const timeDiff = now.getTime() - revokedAtDate.getTime();
    const secondsDiff = timeDiff / 1000;
    
    // La data deve essere stata impostata \"da poco\".
    // Nota: teniamo una finestra più larga per evitare flakiness su macchine lente/CI.
    expect(secondsDiff).toBeGreaterThanOrEqual(0);
    expect(secondsDiff).toBeLessThan(30);
  });

  it('restituisce 200 anche se il token non esiste (idempotente)', async () => {
    // Prova a fare logout con un token inesistente MA valido per lo schema (64 char esadecimali).
    // Se usi un token \"sporco\" (es. 'token_inesistente') lo schema lo blocca prima e ottieni 400.
    const res = await request(app)
      .post('/api/auth/logout')
      .send({
        refreshToken:
          'ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff',
      });

    // La rotta non controlla se il token esiste, quindi restituisce sempre 200
    expect(res.status).toBe(200);
  });

  it('restituisce 400 con dati mancanti', async () => {
    const res = await request(app)
      .post('/api/auth/logout')
      .send({
        // refreshToken mancante
      });

    expect(res.status).toBe(400);
  });
});

describe('POST /api/auth/refresh', () => {
  beforeAll(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('genera nuovo access token e refresh token con successo', async () => {
    // Ottieni un refresh token fresco facendo login
    const loginRes = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        password: testCredentials.admin.password,
      });
    
    const refreshToken = loginRes.body.refreshToken;

    // Usa il refresh token per ottenere nuovi token
    const res = await request(app)
      .post('/api/auth/refresh')
      .send({
        refreshToken,
      });

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('token'); // Nuovo access token
    expect(res.body).toHaveProperty('refreshToken'); // Nuovo refresh token
    expect(res.body.refreshToken).not.toBe(refreshToken); // Deve essere diverso dal vecchio
  });

  it('restituisce 401 con refresh token non valido', async () => {
    // Token valido per lo schema (64 char hex) ma NON presente nel DB -> la rotta arriva alla logica e risponde 401
    const res = await request(app)
      .post('/api/auth/refresh')
      .send({
        refreshToken:
          'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
      });

    expect(res.status).toBe(401);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 401 con refresh token già revocato', async () => {
    // Ottieni un refresh token fresco
    const loginRes = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        password: testCredentials.admin.password,
      });
    
    const refreshToken = loginRes.body.refreshToken;

    // Prima revoca il token con logout
    await request(app)
      .post('/api/auth/logout')
      .send({ refreshToken });

    // Poi prova a usarlo per refresh
    const res = await request(app)
      .post('/api/auth/refresh')
      .send({
        refreshToken,
      });

    expect(res.status).toBe(401);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 400 con dati mancanti', async () => {
    const res = await request(app)
      .post('/api/auth/refresh')
      .send({
        // refreshToken mancante
      });

    expect(res.status).toBe(400);
  });
});

describe('POST /api/auth/change-password', () => {
  beforeAll(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('restituisce 401 senza token', async () => {
    const res = await request(app)
      .post('/api/auth/change-password')
      .send({
        currentPassword: 'password123',
        newPassword: 'nuovaPassword1',
      });

    expect(res.status).toBe(401);
  });

  it('restituisce 400 con password attuale sbagliata', async () => {
    const { token } = await loginAsAdmin();

    const res = await request(app)
      .post('/api/auth/change-password')
      .set('Authorization', `Bearer ${token}`)
      .send({
        currentPassword: 'password_sbagliata',
        newPassword: 'nuovaPassword1',
      });

    expect(res.status).toBe(400);
    expect(res.body.error).toBe('Password attuale non corretta');
  });

  it('restituisce 400 se la nuova password coincide con quella attuale', async () => {
    const { token } = await loginAsAdmin();

    const res = await request(app)
      .post('/api/auth/change-password')
      .set('Authorization', `Bearer ${token}`)
      .send({
        currentPassword: testCredentials.admin.password,
        newPassword: testCredentials.admin.password,
      });

    expect(res.status).toBe(400);
    expect(res.body.error).toBe('La nuova password deve essere diversa da quella attuale');
  });

  it('aggiorna la password admin e permette il login con la nuova', async () => {
    const { token } = await loginAsAdmin();
    const nuovaPassword = 'nuovaPasswordAdmin';

    const res = await request(app)
      .post('/api/auth/change-password')
      .set('Authorization', `Bearer ${token}`)
      .send({
        currentPassword: testCredentials.admin.password,
        newPassword: nuovaPassword,
      });

    expect(res.status).toBe(200);
    expect(res.body.message).toBe('Password aggiornata con successo');

    const loginVecchia = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        password: testCredentials.admin.password,
      });
    expect(loginVecchia.status).toBe(401);

    const loginNuova = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.admin.email,
        password: nuovaPassword,
      });
    expect(loginNuova.status).toBe(200);
  });

  it('aggiorna la password collaboratore e permette il login con la nuova', async () => {
    const { token } = await loginAsCollaboratore();
    const nuovaPassword = 'nuovaPasswordCollab';

    const res = await request(app)
      .post('/api/auth/change-password')
      .set('Authorization', `Bearer ${token}`)
      .send({
        currentPassword: testCredentials.collaboratore.password,
        newPassword: nuovaPassword,
      });

    expect(res.status).toBe(200);

    const loginNuova = await request(app)
      .post('/api/auth/login')
      .send({
        identifier: testCredentials.collaboratore.username,
        password: nuovaPassword,
      });
    expect(loginNuova.status).toBe(200);
  });
});

