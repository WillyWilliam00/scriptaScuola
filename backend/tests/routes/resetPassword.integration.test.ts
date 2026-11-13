import { describe, it, expect, beforeEach, afterAll } from 'vitest';
import request from 'supertest';
import { eq } from 'drizzle-orm';
import { app } from '../../src/index.js';
import { cleanTestDb, testDb } from '../db/test-setup.js';
import { seedTestData, testCredentials } from '../db/test-seed.js';
import { istituti, richiesteResetPassword, utenti } from '../../src/db/schema.js';
import { loginAsAdmin } from '../utils/test-helpers.js';
import { generaPasswordTemporanea } from '../../src/db/utils/passwordTemporanea.js';

const FORMATO_PASSWORD_TEMPORANEA = /^[a-z2-9]{4}-[a-z2-9]{4}-[a-z2-9]{4}$/;

async function login(identifier: string, password: string) {
  return request(app).post('/api/auth/login').send({ identifier, password });
}

async function richiestaReset(username: string, nota?: string) {
  return request(app).post('/api/auth/richiesta-reset').send({ username, nota });
}

async function idCollaboratore() {
  const [row] = await testDb.select({ id: utenti.id }).from(utenti)
    .where(eq(utenti.username, testCredentials.collaboratore.username));
  return row!.id;
}

async function richiesteAperte(token: string) {
  const res = await request(app).get('/api/utenti/richieste-reset').set('Authorization', `Bearer ${token}`);
  expect(res.status).toBe(200);
  return res.body.data as Array<{ id: string; utenteId: string; username: string | null; nota: string | null }>;
}

describe('Password temporanea', () => {
  it('ha un formato leggibile e cambia a ogni chiamata', () => {
    const a = generaPasswordTemporanea();
    const b = generaPasswordTemporanea();
    expect(a).toMatch(FORMATO_PASSWORD_TEMPORANEA);
    expect(a).not.toBe(b);
  });
});

describe('Reset password dei collaboratori', () => {
  beforeEach(async () => {
    await cleanTestDb();
    await seedTestData();
  });

  afterAll(async () => {
    await cleanTestDb();
  });

  describe('POST /api/auth/richiesta-reset', () => {
    it('crea una richiesta visibile agli admin, con la nota', async () => {
      const res = await richiestaReset(testCredentials.collaboratore.username, 'Sono in segreteria');
      expect(res.status).toBe(200);

      const { token } = await loginAsAdmin();
      const richieste = await richiesteAperte(token);
      expect(richieste).toHaveLength(1);
      expect(richieste[0]!.username).toBe(testCredentials.collaboratore.username);
      expect(richieste[0]!.nota).toBe('Sono in segreteria');
    });

    it('risponde allo stesso modo se lo username non esiste, senza creare nulla', async () => {
      const esistente = await richiestaReset(testCredentials.collaboratore.username);
      const inesistente = await richiestaReset('nessuno_si_chiama_cosi');
      expect(inesistente.status).toBe(200);
      expect(inesistente.body).toEqual(esistente.body);

      const righe = await testDb.select().from(richiesteResetPassword);
      expect(righe).toHaveLength(1);
    });

    it('lo username è case-sensitive', async () => {
      await richiestaReset(testCredentials.collaboratore.username.toUpperCase());
      const righe = await testDb.select().from(richiesteResetPassword);
      expect(righe).toHaveLength(0);
    });

    it('non crea richieste per gli admin', async () => {
      const [admin] = await testDb.select().from(utenti).where(eq(utenti.email, testCredentials.admin.email));
      // Un admin con username (caso limite): la richiesta per username vale solo per i collaboratori
      await testDb.update(utenti).set({ username: 'admin_con_username' }).where(eq(utenti.id, admin!.id));
      await richiestaReset('admin_con_username');
      const righe = await testDb.select().from(richiesteResetPassword);
      expect(righe).toHaveLength(0);
    });

    it('tiene al massimo una richiesta aperta per utenza', async () => {
      await richiestaReset(testCredentials.collaboratore.username, 'prima');
      await richiestaReset(testCredentials.collaboratore.username, 'seconda');
      const righe = await testDb.select().from(richiesteResetPassword);
      expect(righe).toHaveLength(1);
      expect(righe[0]!.nota).toBe('prima');
    });

    it('rifiuta una nota più lunga di 200 caratteri', async () => {
      const res = await richiestaReset(testCredentials.collaboratore.username, 'x'.repeat(201));
      expect(res.status).toBe(400);
    });

    it('le richieste più vecchie di 7 giorni non sono più mostrate', async () => {
      await richiestaReset(testCredentials.collaboratore.username);
      const ottoGiorniFa = new Date(Date.now() - 8 * 24 * 60 * 60 * 1000);
      await testDb.update(richiesteResetPassword).set({ createdAt: ottoGiorniFa });

      const { token } = await loginAsAdmin();
      expect(await richiesteAperte(token)).toHaveLength(0);

      // Scaduta la vecchia, se ne può aprire una nuova
      await richiestaReset(testCredentials.collaboratore.username);
      expect(await richiesteAperte(token)).toHaveLength(1);
    });
  });

  describe('Richieste di reset lato admin', () => {
    it('un admin non vede le richieste di un altro istituto', async () => {
      await richiestaReset(testCredentials.collaboratore.username);

      const setup = await request(app).post('/api/auth/setup-scuola').send({
        istituto: { nome: 'Altro istituto', codiceIstituto: 'ALTRO00001' },
        utente: { email: 'admin@altro.it', password: 'password123' },
      });
      expect(setup.status).toBe(201);
      const altroAdmin = await login('admin@altro.it', 'password123');
      expect(await richiesteAperte(altroAdmin.body.token)).toHaveLength(0);
    });

    it('un collaboratore non può vedere le richieste', async () => {
      const collab = await login(testCredentials.collaboratore.username, testCredentials.collaboratore.password);
      const res = await request(app).get('/api/utenti/richieste-reset').set('Authorization', `Bearer ${collab.body.token}`);
      expect(res.status).toBe(403);
    });

    it('ignora una richiesta, che sparisce dall\'elenco', async () => {
      await richiestaReset(testCredentials.collaboratore.username);
      const { token, utente } = await loginAsAdmin();
      const [richiesta] = await richiesteAperte(token);

      const res = await request(app)
        .post(`/api/utenti/richieste-reset/${richiesta!.id}/ignora`)
        .set('Authorization', `Bearer ${token}`);
      expect(res.status).toBe(200);
      expect(await richiesteAperte(token)).toHaveLength(0);

      const [riga] = await testDb.select().from(richiesteResetPassword);
      expect(riga!.esito).toBe('ignorata');
      expect(riga!.gestitaDa).toBe(utente.id);

      // Già gestita: 404
      const ancora = await request(app)
        .post(`/api/utenti/richieste-reset/${richiesta!.id}/ignora`)
        .set('Authorization', `Bearer ${token}`);
      expect(ancora.status).toBe(404);
    });
  });

  describe('POST /api/utenti/:id/password-temporanea', () => {
    it('genera la password, chiude la richiesta e obbliga al cambio', async () => {
      await richiestaReset(testCredentials.collaboratore.username);
      const { token } = await loginAsAdmin();
      const id = await idCollaboratore();

      const res = await request(app)
        .post(`/api/utenti/${id}/password-temporanea`)
        .set('Authorization', `Bearer ${token}`);
      expect(res.status).toBe(200);
      expect(res.body.passwordTemporanea).toMatch(FORMATO_PASSWORD_TEMPORANEA);
      const scadeAt = new Date(res.body.scadeAt).getTime();
      expect(scadeAt).toBeGreaterThan(Date.now() + 47 * 60 * 60 * 1000);
      expect(scadeAt).toBeLessThan(Date.now() + 49 * 60 * 60 * 1000);

      // La richiesta è chiusa come "password generata"
      expect(await richiesteAperte(token)).toHaveLength(0);
      const [riga] = await testDb.select().from(richiesteResetPassword);
      expect(riga!.esito).toBe('password_generata');

      // La vecchia password non vale più, la temporanea sì
      const vecchia = await login(testCredentials.collaboratore.username, testCredentials.collaboratore.password);
      expect(vecchia.status).toBe(401);
      const nuova = await login(testCredentials.collaboratore.username, res.body.passwordTemporanea);
      expect(nuova.status).toBe(200);
      expect(nuova.body.utente.deveCambiarePassword).toBe(true);
    });

    it('con la password temporanea si può solo cambiare la password', async () => {
      const { token } = await loginAsAdmin();
      const id = await idCollaboratore();
      const gen = await request(app).post(`/api/utenti/${id}/password-temporanea`).set('Authorization', `Bearer ${token}`);
      const collab = await login(testCredentials.collaboratore.username, gen.body.passwordTemporanea);

      const bloccata = await request(app).get('/api/docenti').set('Authorization', `Bearer ${collab.body.token}`);
      expect(bloccata.status).toBe(403);
      expect(bloccata.body.code).toBe('CAMBIO_PASSWORD_OBBLIGATORIO');

      const cambio = await request(app)
        .post('/api/auth/change-password')
        .set('Authorization', `Bearer ${collab.body.token}`)
        .send({ currentPassword: gen.body.passwordTemporanea, newPassword: 'nuovaPassword1' });
      expect(cambio.status).toBe(200);

      const sbloccata = await request(app).get('/api/docenti').set('Authorization', `Bearer ${collab.body.token}`);
      expect(sbloccata.status).toBe(200);

      const dopo = await login(testCredentials.collaboratore.username, 'nuovaPassword1');
      expect(dopo.status).toBe(200);
      expect(dopo.body.utente.deveCambiarePassword).toBe(false);
    });

    it('una password temporanea scaduta non permette l\'accesso', async () => {
      const { token } = await loginAsAdmin();
      const id = await idCollaboratore();
      const gen = await request(app).post(`/api/utenti/${id}/password-temporanea`).set('Authorization', `Bearer ${token}`);
      await testDb.update(utenti).set({ passwordTemporaneaScadeAt: new Date(Date.now() - 1000) }).where(eq(utenti.id, id));

      const res = await login(testCredentials.collaboratore.username, gen.body.passwordTemporanea);
      expect(res.status).toBe(401);
      expect(res.body.error).toMatch(/scaduta/);
    });

    it('un admin non può generarla per sé stesso', async () => {
      const { token, utente } = await loginAsAdmin();
      const res = await request(app).post(`/api/utenti/${utente.id}/password-temporanea`).set('Authorization', `Bearer ${token}`);
      expect(res.status).toBe(400);
    });

    it('un admin non può generarla per un utente di un altro istituto', async () => {
      const id = await idCollaboratore();
      await request(app).post('/api/auth/setup-scuola').send({
        istituto: { nome: 'Altro istituto', codiceIstituto: 'ALTRO00002' },
        utente: { email: 'admin@altro2.it', password: 'password123' },
      });
      const altroAdmin = await login('admin@altro2.it', 'password123');
      const res = await request(app).post(`/api/utenti/${id}/password-temporanea`).set('Authorization', `Bearer ${altroAdmin.body.token}`);
      expect(res.status).toBe(404);

      // La password del collaboratore non è cambiata
      const collab = await login(testCredentials.collaboratore.username, testCredentials.collaboratore.password);
      expect(collab.status).toBe(200);
    });
  });

  describe('Creazione di un\'utenza', () => {
    it('la piattaforma genera la password temporanea, da cambiare al primo accesso', async () => {
      const { token } = await loginAsAdmin();
      const res = await request(app)
        .post('/api/utenti/new-utente')
        .set('Authorization', `Bearer ${token}`)
        .send({ ruolo: 'collaboratore', username: 'nuovo_collab' });
      expect(res.status).toBe(201);
      expect(res.body.passwordTemporanea).toMatch(FORMATO_PASSWORD_TEMPORANEA);
      expect(res.body.data).not.toHaveProperty('passwordHash');

      const primoAccesso = await login('nuovo_collab', res.body.passwordTemporanea);
      expect(primoAccesso.status).toBe(200);
      expect(primoAccesso.body.utente.deveCambiarePassword).toBe(true);
    });

    it('una password inviata dal client viene rifiutata (400)', async () => {
      const { token } = await loginAsAdmin();
      const crea = await request(app)
        .post('/api/utenti/new-utente')
        .set('Authorization', `Bearer ${token}`)
        .send({ ruolo: 'collaboratore', username: 'altro_collab', password: 'sceltaDalClient1' });
      expect(crea.status).toBe(400);

      const id = await idCollaboratore();
      const modifica = await request(app)
        .put(`/api/utenti/update-utente/${id}`)
        .set('Authorization', `Bearer ${token}`)
        .send({ ruolo: 'collaboratore', username: testCredentials.collaboratore.username, campoExtra: 'x' });
      expect(modifica.status).toBe(400);
    });

    it('la registrazione dell\'istituto usa la password scelta, senza obbligo di cambio', async () => {
      await request(app).post('/api/auth/setup-scuola').send({
        istituto: { nome: 'Scuola nuova', codiceIstituto: 'NUOVA00001' },
        utente: { email: 'admin@nuova.it', password: 'password123' },
      });
      const res = await login('admin@nuova.it', 'password123');
      expect(res.status).toBe(200);
      expect(res.body.utente.deveCambiarePassword).toBe(false);
      const [istituto] = await testDb.select().from(istituti).where(eq(istituti.codiceIstituto, 'NUOVA00001'));
      expect(istituto).toBeDefined();
    });
  });
});
