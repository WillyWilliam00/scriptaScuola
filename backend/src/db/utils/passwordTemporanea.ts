import crypto from 'crypto';
import bcrypt from 'bcrypt';

/** Validità di una password temporanea generata dalla piattaforma. */
export const PASSWORD_TEMPORANEA_ORE = 48;

/** Una richiesta di reset aperta da più di 7 giorni non viene più mostrata agli admin. */
export const RICHIESTA_RESET_GIORNI = 7;

// Niente caratteri che si confondono quando la password viene dettata o copiata a mano (0/o, 1/l/i).
const ALFABETO = 'abcdefghjkmnpqrstuvwxyz23456789';

/**
 * Genera una password temporanea leggibile, per esempio `kq7m-x3pa-tz9e`.
 * 12 caratteri casuali su 31 simboli: circa 59 bit, abbondanti per una password
 * che scade dopo 48 ore e va cambiata al primo accesso.
 */
export function generaPasswordTemporanea(): string {
  const gruppi: string[] = [];
  for (let g = 0; g < 3; g++) {
    let gruppo = '';
    for (let i = 0; i < 4; i++) {
      gruppo += ALFABETO[crypto.randomInt(ALFABETO.length)];
    }
    gruppi.push(gruppo);
  }
  return gruppi.join('-');
}

/** Scadenza di una password temporanea generata ora. */
export function scadenzaPasswordTemporanea(now = new Date()): Date {
  return new Date(now.getTime() + PASSWORD_TEMPORANEA_ORE * 60 * 60 * 1000);
}

/** Inizio della finestra in cui una richiesta di reset è ancora aperta. */
export function inizioFinestraRichieste(now = new Date()): Date {
  return new Date(now.getTime() - RICHIESTA_RESET_GIORNI * 24 * 60 * 60 * 1000);
}

/**
 * Prepara i valori da salvare su `utenti` per una nuova password temporanea:
 * hash, obbligo di cambio al primo accesso e scadenza.
 */
export async function valoriPasswordTemporanea() {
  const passwordTemporanea = generaPasswordTemporanea();
  const passwordTemporaneaScadeAt = scadenzaPasswordTemporanea();
  const passwordHash = await bcrypt.hash(passwordTemporanea, 10);
  return {
    passwordTemporanea,
    valori: {
      passwordHash,
      deveCambiarePassword: true,
      passwordTemporaneaScadeAt,
    },
  };
}
