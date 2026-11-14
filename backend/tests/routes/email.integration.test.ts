import { describe, it, expect, beforeEach, afterAll } from 'vitest';
import request from 'supertest';
import { eq } from 'drizzle-orm';
import { app } from '../../src/index.js';
import { cleanTestDb, testDb } from '../db/test-setup.js';
import { seedTestData, testCredentials } from '../db/test-seed.js';
import { tokenEmail, utenti } from '../../src/db/schema.js';
import { emailInviate, svuotaEmailInviate } from '../../src/email/index.js';
import { getRefreshCookie, loginAsAdmin, refreshCookieHeader } from '../utils/test-helpers.js';

/** Le email partono in background: aspetta che il giro di eventi le abbia registrate. */
async function attendiEmail(n = 1) {
  for (let i = 0; i < 50 && emailInviate.length < n; i++) {
    await new Promise((r) => setTimeout(r, 10));
  }
}

function tokenDaEmail(indice = emailInviate.length - 1): string {
  const match = emailInviate[indice]?.testo.match(/token=([a-f0-9]{64})/);
  expect(match).toBeTruthy();
  return match![1]!;
}

async function login(identifier: string, password: string) {
  return request(app).post('/api/auth/login').send({ identifier, password });
}

async function verificaEmailAdmin() {
  await testDb.update(utenti).set({ emailVerificataAt: new Date() }).where(eq(utenti.email, testCredentials.admin.email));
}

describe('Email: verifica e recupero password degli admin', () => {
  beforeEach(async () => {
    await cleanTestDb();
    await seedTestData();
    svuotaEmailInviate();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  describe('Verifica dell\'indirizzo', () => {
    it('la registrazione della scuola manda l\'email di verifica e il link conferma l\'indirizzo', async () => {
      const setup = await request(app).post('/api/auth/setup-scuola').send({
        istituto: { nome: 'Scuola email', codiceIstituto: 'EMAIL00001' },
        utente: { email: 'nuovo@scuola.it', password: 'password123' },
      });
      expect(setup.status).toBe(201);
      await attendiEmail();
      expect(emailInviate).toHaveLength(1);
      expect(emailInviate[0]!.a).toBe('nuovo@scuola.it');
      expect(emailInviate[0]!.testo).toContain('/verifica-email?token=');

      const prima = await login('nuovo@scuola.it', 'password123');
      expect(prima.body.utente.emailVerificata).toBe(false);

      const token = tokenDaEmail();
      const verifica = await request(app).post('/api/auth/verifica-email').send({ token });
      expect(verifica.status).toBe(200);

      const dopo = await login('nuovo@scuola.it', 'password123');
      expect(dopo.body.utente.emailVerificata).toBe(true);

      // Il link vale una volta sola
      const ancora = await request(app).post('/api/auth/verifica-email').send({ token });
      expect(ancora.status).toBe(400);
    });

    it('nel database c\'è solo l\'hash del token', async () => {
      const { token } = await loginAsAdmin();
      await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${token}`);
      const inChiaro = tokenDaEmail();
      const [riga] = await testDb.select().from(tokenEmail);
      expect(riga!.tokenHash).toMatch(/^[a-f0-9]{64}$/);
      expect(riga!.tokenHash).not.toBe(inChiaro);
    });

    it('si può chiedere di nuovo l\'email: vale solo l\'ultimo link', async () => {
      const { token } = await loginAsAdmin();
      const r1 = await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${token}`);
      expect(r1.status).toBe(200);
      const vecchio = tokenDaEmail();
      await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${token}`);
      const nuovo = tokenDaEmail();

      expect((await request(app).post('/api/auth/verifica-email').send({ token: vecchio })).status).toBe(400);
      expect((await request(app).post('/api/auth/verifica-email').send({ token: nuovo })).status).toBe(200);

      const giaVerificata = await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${token}`);
      expect(giaVerificata.status).toBe(400);
    });

    it('un collaboratore non ha email da verificare', async () => {
      const collab = await login(testCredentials.collaboratore.username, testCredentials.collaboratore.password);
      expect(collab.body.utente.emailVerificata).toBe(false);
      const res = await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${collab.body.token}`);
      expect(res.status).toBe(400);
      expect(emailInviate).toHaveLength(0);
    });

    it('un link scaduto non vale', async () => {
      const { token } = await loginAsAdmin();
      await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${token}`);
      await testDb.update(tokenEmail).set({ expiresAt: new Date(Date.now() - 1000) });
      const res = await request(app).post('/api/auth/verifica-email').send({ token: tokenDaEmail() });
      expect(res.status).toBe(400);
    });

    it('se l\'admin cambia email, la verifica si azzera e il vecchio link non vale più', async () => {
      await verificaEmailAdmin();
      const { token, utente } = await loginAsAdmin();
      // Un secondo admin modifica l'email del primo
      await request(app).post('/api/utenti/new-utente').set('Authorization', `Bearer ${token}`)
        .send({ ruolo: 'admin', email: 'secondo@test.com' });
      const [secondo] = await testDb.select().from(utenti).where(eq(utenti.email, 'secondo@test.com'));
      await testDb.update(utenti).set({ deveCambiarePassword: false }).where(eq(utenti.id, secondo!.id));

      await testDb.update(utenti).set({ emailVerificataAt: null }).where(eq(utenti.id, utente.id));
      await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${token}`);
      const linkVecchio = tokenDaEmail();

      const modifica = await request(app).put(`/api/utenti/update-utente/${utente.id}`)
        .set('Authorization', `Bearer ${token}`)
        .send({ ruolo: 'admin', email: 'cambiata@test.com' });
      expect(modifica.status).toBe(200);

      const res = await request(app).post('/api/auth/verifica-email').send({ token: linkVecchio });
      expect(res.status).toBe(400);
      const [riga] = await testDb.select().from(utenti).where(eq(utenti.id, utente.id));
      expect(riga!.emailVerificataAt).toBeNull();
    });
  });

  describe('Password dimenticata di un admin', () => {
    it('manda il link, la nuova password funziona e le sessioni vecchie vengono chiuse', async () => {
      const sessione = await login(testCredentials.admin.email, testCredentials.admin.password);
      const cookieVecchio = getRefreshCookie(sessione)!;

      const res = await request(app).post('/api/auth/password-dimenticata').send({ email: testCredentials.admin.email });
      expect(res.status).toBe(200);
      await attendiEmail();
      expect(emailInviate[0]!.a).toBe(testCredentials.admin.email);
      expect(emailInviate[0]!.testo).toContain('/reimposta-password?token=');

      const token = tokenDaEmail();
      const reimposta = await request(app).post('/api/auth/reimposta-password').send({ token, newPassword: 'nuovaPassword1' });
      expect(reimposta.status).toBe(200);

      expect((await login(testCredentials.admin.email, testCredentials.admin.password)).status).toBe(401);
      const nuova = await login(testCredentials.admin.email, 'nuovaPassword1');
      expect(nuova.status).toBe(200);
      // Aver usato il link conferma anche l'email
      expect(nuova.body.utente.emailVerificata).toBe(true);

      const refresh = await request(app).post('/api/auth/refresh').set('Cookie', refreshCookieHeader(cookieVecchio));
      expect(refresh.status).toBe(401);

      // Il link vale una volta sola
      const ancora = await request(app).post('/api/auth/reimposta-password').send({ token, newPassword: 'altraPassword1' });
      expect(ancora.status).toBe(400);
    });

    it('risponde allo stesso modo per email inesistenti e collaboratori, senza mandare nulla', async () => {
      const esistente = await request(app).post('/api/auth/password-dimenticata').send({ email: testCredentials.admin.email });
      await attendiEmail();
      svuotaEmailInviate();
      const inesistente = await request(app).post('/api/auth/password-dimenticata').send({ email: 'nessuno@test.com' });
      expect(inesistente.status).toBe(200);
      expect(inesistente.body).toEqual(esistente.body);
      await new Promise((r) => setTimeout(r, 50));
      expect(emailInviate).toHaveLength(0);
    });

    it('vale solo l\'ultimo link chiesto', async () => {
      await request(app).post('/api/auth/password-dimenticata').send({ email: testCredentials.admin.email });
      await attendiEmail(1);
      const primo = tokenDaEmail(0);
      await request(app).post('/api/auth/password-dimenticata').send({ email: testCredentials.admin.email });
      await attendiEmail(2);
      const secondo = tokenDaEmail(1);

      expect((await request(app).post('/api/auth/reimposta-password').send({ token: primo, newPassword: 'nuovaPassword1' })).status).toBe(400);
      expect((await request(app).post('/api/auth/reimposta-password').send({ token: secondo, newPassword: 'nuovaPassword1' })).status).toBe(200);
    });

    it('un link scaduto o inventato non vale', async () => {
      await request(app).post('/api/auth/password-dimenticata').send({ email: testCredentials.admin.email });
      await attendiEmail();
      await testDb.update(tokenEmail).set({ expiresAt: new Date(Date.now() - 1000) });
      expect((await request(app).post('/api/auth/reimposta-password').send({ token: tokenDaEmail(), newPassword: 'nuovaPassword1' })).status).toBe(400);
      expect((await request(app).post('/api/auth/reimposta-password').send({ token: 'a'.repeat(64), newPassword: 'nuovaPassword1' })).status).toBe(400);
      expect((await request(app).post('/api/auth/reimposta-password').send({ token: 'non-esadecimale', newPassword: 'nuovaPassword1' })).status).toBe(400);
    });

    it('un link di verifica non vale per reimpostare la password', async () => {
      const { token } = await loginAsAdmin();
      await request(app).post('/api/auth/invia-verifica-email').set('Authorization', `Bearer ${token}`);
      const res = await request(app).post('/api/auth/reimposta-password').send({ token: tokenDaEmail(), newPassword: 'nuovaPassword1' });
      expect(res.status).toBe(400);
    });

    it('il link chiude anche una password temporanea in sospeso', async () => {
      await testDb.update(utenti).set({ deveCambiarePassword: true, passwordTemporaneaScadeAt: new Date(Date.now() + 3600_000) })
        .where(eq(utenti.email, testCredentials.admin.email));
      await request(app).post('/api/auth/password-dimenticata').send({ email: testCredentials.admin.email });
      await attendiEmail();
      await request(app).post('/api/auth/reimposta-password').send({ token: tokenDaEmail(), newPassword: 'nuovaPassword1' });
      const res = await login(testCredentials.admin.email, 'nuovaPassword1');
      expect(res.body.utente.deveCambiarePassword).toBe(false);
    });
  });

  describe('Avviso agli admin per le richieste dei collaboratori', () => {
    it('manda un\'email agli admin con l\'email verificata', async () => {
      await verificaEmailAdmin();
      await request(app).post('/api/auth/richiesta-reset').send({ username: testCredentials.collaboratore.username, nota: 'In segreteria <b>ora</b>' });
      await attendiEmail();
      expect(emailInviate).toHaveLength(1);
      expect(emailInviate[0]!.a).toBe(testCredentials.admin.email);
      expect(emailInviate[0]!.testo).toContain(testCredentials.collaboratore.username);
      // La nota arriva come testo, mai come HTML
      expect(emailInviate[0]!.html).toContain('&lt;b&gt;ora&lt;/b&gt;');
      expect(emailInviate[0]!.html).not.toContain('<b>ora</b>');
    });

    it('non manda nulla agli admin con l\'email non verificata', async () => {
      await request(app).post('/api/auth/richiesta-reset').send({ username: testCredentials.collaboratore.username });
      await new Promise((r) => setTimeout(r, 50));
      expect(emailInviate).toHaveLength(0);
    });

    it('una seconda richiesta mentre la prima è aperta non manda altre email', async () => {
      await verificaEmailAdmin();
      await request(app).post('/api/auth/richiesta-reset').send({ username: testCredentials.collaboratore.username });
      await request(app).post('/api/auth/richiesta-reset').send({ username: testCredentials.collaboratore.username });
      await new Promise((r) => setTimeout(r, 50));
      expect(emailInviate).toHaveLength(1);
    });
  });

  it('GET /api/auth/me restituisce i dati aggiornati senza l\'hash della password', async () => {
    const { token } = await loginAsAdmin();
    await verificaEmailAdmin();
    const res = await request(app).get('/api/auth/me').set('Authorization', `Bearer ${token}`);
    expect(res.status).toBe(200);
    expect(res.body.emailVerificata).toBe(true);
    expect(res.body.haPassword).toBe(true);
    expect(JSON.stringify(res.body)).not.toMatch(/password_?hash|\$2[aby]\$/i);
  });
});
