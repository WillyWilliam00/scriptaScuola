import { describe, it, expect } from 'vitest';
import { getAllowedOrigins } from '../../src/index.js';

describe('getAllowedOrigins', () => {
  it('in sviluppo ammette localhost:3000 di default', () => {
    expect(getAllowedOrigins({})).toEqual(['http://localhost:3000']);
  });

  it('aggiunge FRONTEND_ORIGIN e le origini di ALLOWED_ORIGINS, senza spazi né doppioni', () => {
    expect(getAllowedOrigins({
      FRONTEND_ORIGIN: 'https://scriptascuola.it',
      ALLOWED_ORIGINS: ' http://192.168.1.10:3000, ,https://scriptascuola.it',
    })).toEqual(['https://scriptascuola.it', 'http://192.168.1.10:3000', 'http://localhost:3000']);
  });

  it('in produzione non ammette localhost se non indicato', () => {
    expect(getAllowedOrigins({ NODE_ENV: 'production', FRONTEND_ORIGIN: 'https://scriptascuola.it' }))
      .toEqual(['https://scriptascuola.it']);
  });
});
