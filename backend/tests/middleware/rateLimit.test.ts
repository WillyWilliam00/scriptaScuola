import { describe, it, expect, afterEach } from 'vitest';
import express from 'express';
import request from 'supertest';
import { createAuthLimiter, getTrustProxy } from '../../src/middleware/rateLimit.js';

function buildApp(limit: number) {
  const app = express();
  app.set('trust proxy', 1);
  // skip: () => false per attivare il limite anche con NODE_ENV=test
  app.post('/login', createAuthLimiter({ limit, skip: () => false }), (_req, res) => {
    res.status(200).json({ ok: true });
  });
  return app;
}

describe('createAuthLimiter', () => {
  it('con il limite predefinito l\'undicesimo tentativo restituisce 429', async () => {
    const app = express();
    app.post('/login', createAuthLimiter({ skip: () => false }), (_req, res) => {
      res.status(200).json({ ok: true });
    });

    for (let i = 0; i < 10; i++) {
      const res = await request(app).post('/login');
      expect(res.status).toBe(200);
    }
    const res = await request(app).post('/login');
    expect(res.status).toBe(429);
    expect(res.body.error).toBe('Troppi tentativi. Riprova tra qualche minuto.');
    expect(res.headers).toHaveProperty('retry-after');
  });

  it('conta i tentativi separatamente per ogni IP', async () => {
    const app = buildApp(2);

    for (let i = 0; i < 2; i++) {
      await request(app).post('/login').set('X-Forwarded-For', '203.0.113.1');
    }
    const bloccato = await request(app).post('/login').set('X-Forwarded-For', '203.0.113.1');
    expect(bloccato.status).toBe(429);

    const altroIp = await request(app).post('/login').set('X-Forwarded-For', '203.0.113.2');
    expect(altroIp.status).toBe(200);
  });

  it('è disattivato in ambiente di test, così i test di integrazione non vengono bloccati', async () => {
    const app = express();
    app.post('/login', createAuthLimiter({ limit: 1 }), (_req, res) => {
      res.status(200).json({ ok: true });
    });

    for (let i = 0; i < 5; i++) {
      const res = await request(app).post('/login');
      expect(res.status).toBe(200);
    }
  });
});

describe('getTrustProxy', () => {
  const originale = process.env.TRUST_PROXY;
  afterEach(() => {
    if (originale === undefined) delete process.env.TRUST_PROXY;
    else process.env.TRUST_PROXY = originale;
  });

  it('usa 2 come valore predefinito (Vercel + Render)', () => {
    delete process.env.TRUST_PROXY;
    expect(getTrustProxy()).toBe(2);
  });

  it('legge il numero di proxy da TRUST_PROXY', () => {
    process.env.TRUST_PROXY = '1';
    expect(getTrustProxy()).toBe(1);
  });

  it('ignora valori non validi', () => {
    process.env.TRUST_PROXY = 'abc';
    expect(getTrustProxy()).toBe(2);
  });
});
