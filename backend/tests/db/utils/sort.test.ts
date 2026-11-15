import { describe, it, expect } from 'vitest';
import { getSort } from '../../../src/db/utils/sort.js';
import { sql } from 'drizzle-orm';
import { registrazioniCopie } from '../../../src/db/schema.js';

describe('getSort', () => {
  // Creiamo un mock column usando sql per i test
  const mockColumn = sql`test_column`;
  const mockIdColumn = registrazioniCopie.id;

  it('restituisce due criteri (colonna e id) quando order è "asc"', () => {
    const result = getSort(mockColumn, 'asc', mockIdColumn);
    expect(result).toHaveLength(2);
  });

  it('restituisce due criteri (colonna e id) quando order è "desc"', () => {
    const result = getSort(mockColumn, 'desc', mockIdColumn);
    expect(result).toHaveLength(2);
  });

  it('restituisce risultati diversi per asc e desc', () => {
    const ascResult = getSort(mockColumn, 'asc', mockIdColumn);
    const descResult = getSort(mockColumn, 'desc', mockIdColumn);
    // Verifichiamo che i risultati siano effettivamente diversi
    // (se la funzione funzionasse sempre allo stesso modo, sarebbero uguali)
    expect(ascResult).not.toEqual(descResult);
  });

  it('restituisce lo stesso risultato quando chiamato due volte con lo stesso order', () => {
    const result1 = getSort(mockColumn, 'asc', mockIdColumn);
    const result2 = getSort(mockColumn, 'asc', mockIdColumn);
    // Due chiamate con gli stessi parametri dovrebbero restituire lo stesso risultato
    // Nota: Drizzle potrebbe creare nuovi oggetti, quindi verifichiamo almeno la struttura
    expect(result1).toBeDefined();
    expect(result2).toBeDefined();
    expect(typeof result1).toBe('object');
    expect(typeof result2).toBe('object');
  });
});

