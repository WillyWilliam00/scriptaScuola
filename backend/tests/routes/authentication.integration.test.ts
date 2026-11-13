import { describe, it, expect, beforeAll, afterAll, beforeEach } from 'vitest';
import request from 'supertest';
import { app } from '../../src/index.js';
import { cleanTestDb, testDb } from '../db/test-setup.js';
import { seedTestData, testCredentials } from '../db/test-seed.js';
import { refreshTokens } from '../../src/db/schema.js';
import { eq } from 'drizzle-orm';
import { getRefreshCookie, loginAsAdmin, loginAsCollaboratore, refreshCookieHeader } from '../utils/test-helpers.js';
import { hashRefreshToken } from '../../src/db/utils/refreshTokens.js';

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
    // Il refresh token viaggia solo nel cookie httpOnly, mai nel corpo
    expect(res.body).not.toHaveProperty('refreshToken');
    expect(getRefreshCookie(res)).toMatch(/^[a-f0-9]{64}$/);
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
    // Il refresh token viaggia solo nel cookie httpOnly, mai nel corpo
    expect(res.body).not.toHaveProperty('refreshToken');
    expect(getRefreshCookie(res)).toMatch(/^[a-f0-9]{64}$/);
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

const login = (identifier: string, password: string) =>
  request(app).post('/api/auth/login').send({ identifier, password });
const loginAdmin = () => login(testCredentials.admin.email, testCredentials.admin.password);
const refresh = (token: string) =>
  request(app).post('/api/auth/refresh').set('Cookie', refreshCookieHeader(token));
const logout = (token: string) =>
  request(app).post('/api/auth/logout').set('Cookie', refreshCookieHeader(token));

describe('Cookie del refresh token', () => {
  beforeAll(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('il login imposta un cookie httpOnly, SameSite=Strict, limitato a /api/auth, della durata di 7 giorni', async () => {
    const res = await loginAdmin();
    const cookie = (res.headers['set-cookie'] as unknown as string[]).find((c) => c.startsWith('refreshToken='))!;

    expect(cookie).toMatch(/HttpOnly/i);
    expect(cookie).toMatch(/SameSite=Strict/i);
    expect(cookie).toMatch(/Path=\/api\/auth/i);
    expect(cookie).toMatch(/Max-Age=604800/i);
  });

  it('nei test il cookie non è Secure (supertest usa http)', async () => {
    const res = await loginAdmin();
    const cookie = (res.headers['set-cookie'] as unknown as string[]).find((c) => c.startsWith('refreshToken='))!;
    expect(cookie).not.toMatch(/Secure/i);
  });

  it('fuori dai test il cookie è Secure', async () => {
    const precedente = process.env.NODE_ENV;
    process.env.NODE_ENV = 'production';
    try {
      const res = await loginAdmin();
      const cookie = (res.headers['set-cookie'] as unknown as string[]).find((c) => c.startsWith('refreshToken='))!;
      expect(cookie).toMatch(/Secure/i);
    } finally {
      process.env.NODE_ENV = precedente;
    }
  });

  it('un agente che conserva i cookie rinnova la sessione senza inviare nulla nel corpo', async () => {
    const agent = request.agent(app);
    const loginRes = await agent.post('/api/auth/login').send({
      identifier: testCredentials.admin.email,
      password: testCredentials.admin.password,
    });
    expect(loginRes.status).toBe(200);

    const res = await agent.post('/api/auth/refresh');
    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('token');
  });

  it('il refresh token inviato nel corpo, come prima, non viene più accettato', async () => {
    const refreshToken = getRefreshCookie(await loginAdmin())!;

    const res = await request(app).post('/api/auth/refresh').send({ refreshToken });
    expect(res.status).toBe(401);
  });

  it('rifiuta con 403 una richiesta di refresh che arriva da un\'altra origine', async () => {
    const refreshToken = getRefreshCookie(await loginAdmin())!;

    const res = await refresh(refreshToken).set('Origin', 'https://sito-malevolo.example');
    expect(res.status).toBe(403);

    // Il token non è stato consumato: dalla nostra origine funziona ancora
    const ok = await refresh(refreshToken).set('Origin', 'http://localhost:3000');
    expect(ok.status).toBe(200);
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

  it('revoca il refresh token e cancella il cookie', async () => {
    const refreshToken = getRefreshCookie(await loginAdmin())!;

    // Verifica che il token esista e non sia ancora revocato
    const [tokenBefore] = await testDb
      .select()
      .from(refreshTokens)
      .where(eq(refreshTokens.tokenHash, hashRefreshToken(refreshToken)));

    expect(tokenBefore).toBeDefined();
    expect(tokenBefore?.revokedAt).toBeNull(); // Non ancora revocato

    const res = await logout(refreshToken);

    expect(res.status).toBe(200);
    expect(res.body.message).toBe('Logout effettuato con successo');
    // Il server chiede al browser di cancellare il cookie
    expect(getRefreshCookie(res)).toBeNull();
    expect(String(res.headers['set-cookie'])).toMatch(/refreshToken=;/);

    // Verifica che il token sia stato effettivamente revocato nel DB
    const [tokenAfter] = await testDb
      .select()
      .from(refreshTokens)
      .where(eq(refreshTokens.tokenHash, hashRefreshToken(refreshToken)));

    expect(tokenAfter?.revokedAt).toBeInstanceOf(Date);

    // La data deve essere stata impostata "da poco" (finestra larga per macchine lente/CI)
    const secondsDiff = (Date.now() - tokenAfter!.revokedAt!.getTime()) / 1000;
    expect(secondsDiff).toBeGreaterThanOrEqual(0);
    expect(secondsDiff).toBeLessThan(30);
  });

  it('restituisce 200 anche se il token non esiste (idempotente)', async () => {
    const res = await logout('f'.repeat(64));
    expect(res.status).toBe(200);
  });

  it('restituisce 200 anche senza cookie o con un cookie malformato', async () => {
    const senzaCookie = await request(app).post('/api/auth/logout');
    expect(senzaCookie.status).toBe(200);

    const malformato = await logout('token_non_valido');
    expect(malformato.status).toBe(200);
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

  it('genera un nuovo access token e ruota il refresh token nel cookie', async () => {
    const refreshToken = getRefreshCookie(await loginAdmin())!;

    const res = await refresh(refreshToken);

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('token'); // Nuovo access token
    expect(res.body).not.toHaveProperty('refreshToken');
    const nuovo = getRefreshCookie(res);
    expect(nuovo).toMatch(/^[a-f0-9]{64}$/);
    expect(nuovo).not.toBe(refreshToken); // Deve essere diverso dal vecchio

    // Il vecchio token è stato consumato dalla rotazione
    const riuso = await refresh(refreshToken);
    expect(riuso.status).toBe(401);
  });

  it('restituisce 401 con refresh token non presente nel database', async () => {
    const res = await refresh('a'.repeat(64));

    expect(res.status).toBe(401);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 401 con refresh token già revocato', async () => {
    const refreshToken = getRefreshCookie(await loginAdmin())!;

    await logout(refreshToken);

    const res = await refresh(refreshToken);
    expect(res.status).toBe(401);
    expect(res.body).toHaveProperty('error');
  });

  it('restituisce 401 senza cookie', async () => {
    const res = await request(app).post('/api/auth/refresh');
    expect(res.status).toBe(401);
  });

  it('restituisce 401 con un cookie che non ha il formato di un refresh token', async () => {
    for (const valore of ['a'.repeat(63), 'a'.repeat(65), 'A'.repeat(64), 'a'.repeat(63) + 'g']) {
      const res = await refresh(valore);
      expect(res.status).toBe(401);
    }
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


describe('Refresh token: hash nel database e revoca delle sessioni', () => {
  beforeEach(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  it('non salva mai il refresh token in chiaro', async () => {
    const refreshToken = getRefreshCookie(await loginAdmin())!;

    const rows = await testDb.select().from(refreshTokens);
    expect(rows.length).toBeGreaterThan(0);
    expect(rows.some((r) => r.tokenHash === refreshToken)).toBe(false);
    expect(rows.some((r) => r.tokenHash === hashRefreshToken(refreshToken))).toBe(true);

    // Anche dopo la rotazione il nuovo token è salvato solo come hash
    const refreshRes = await refresh(refreshToken);
    expect(refreshRes.status).toBe(200);
    const nuovo = getRefreshCookie(refreshRes)!;
    const rowsDopo = await testDb.select().from(refreshTokens);
    expect(rowsDopo.some((r) => r.tokenHash === nuovo)).toBe(false);
    expect(rowsDopo.some((r) => r.tokenHash === hashRefreshToken(nuovo))).toBe(true);
  });

  it('un hash letto dal database non funziona come refresh token', async () => {
    const hash = hashRefreshToken(getRefreshCookie(await loginAdmin())!);

    const res = await refresh(hash);
    expect(res.status).toBe(401);
  });

  it('dopo il cambio password i vecchi refresh token restituiscono 401 e quello nuovo funziona', async () => {
    // Due sessioni dello stesso utente (es. due dispositivi)
    const sessioneA = await loginAdmin();
    const sessioneB = await loginAdmin();

    const res = await request(app)
      .post('/api/auth/change-password')
      .set('Authorization', `Bearer ${sessioneA.body.token}`)
      .send({ currentPassword: testCredentials.admin.password, newPassword: 'nuovaPassword1' });
    expect(res.status).toBe(200);
    expect(res.body).not.toHaveProperty('refreshToken');
    const nuovoToken = getRefreshCookie(res);
    expect(nuovoToken).toMatch(/^[a-f0-9]{64}$/);

    for (const vecchia of [sessioneA, sessioneB]) {
      const refreshRes = await refresh(getRefreshCookie(vecchia)!);
      expect(refreshRes.status).toBe(401);
    }

    // La sessione di chi ha cambiato la password continua con il nuovo cookie
    const nuovo = await refresh(nuovoToken!);
    expect(nuovo.status).toBe(200);
  });

  it('quando l\'admin imposta una nuova password a un utente, le sue sessioni vengono revocate', async () => {
    const admin = await loginAsAdmin();
    const collab = await login(testCredentials.collaboratore.username, testCredentials.collaboratore.password);
    expect(collab.status).toBe(200);

    const res = await request(app)
      .put(`/api/utenti/update-utente/${collab.body.utente.id}`)
      .set('Authorization', `Bearer ${admin.token}`)
      .send({ ruolo: 'collaboratore', username: testCredentials.collaboratore.username, password: 'passwordNuova1' });
    expect(res.status).toBe(200);

    const refreshRes = await refresh(getRefreshCookie(collab)!);
    expect(refreshRes.status).toBe(401);
  });

  it('una modifica dell\'utente senza nuova password non revoca le sessioni', async () => {
    const admin = await loginAsAdmin();
    const collab = await login(testCredentials.collaboratore.username, testCredentials.collaboratore.password);

    const res = await request(app)
      .put(`/api/utenti/update-utente/${collab.body.utente.id}`)
      .set('Authorization', `Bearer ${admin.token}`)
      .send({ ruolo: 'collaboratore', username: 'collaboratore_rinominato' });
    expect(res.status).toBe(200);

    const refreshRes = await refresh(getRefreshCookie(collab)!);
    expect(refreshRes.status).toBe(200);
  });
});
