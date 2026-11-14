/**
 * Invio delle email transazionali (verifica dell'indirizzo, recupero password, avvisi agli admin).
 *
 * - Con RESEND_API_KEY impostata le email partono davvero tramite l'API di Resend.
 * - Senza chiave (sviluppo locale e test) nessuna email esce: il messaggio viene
 *   scritto nel log e conservato in memoria (`emailInviate`), così i test possono leggerlo.
 *
 * Variabili d'ambiente:
 * - RESEND_API_KEY: chiave API di Resend.
 * - EMAIL_MITTENTE: mittente, per esempio "ScriptaScuola <no-reply@scriptascuola.it>".
 *   Il dominio deve essere verificato su Resend.
 * - FRONTEND_ORIGIN: indirizzo del sito, usato per costruire i link nelle email.
 */

export interface Email {
  a: string;
  oggetto: string;
  testo: string;
  html: string;
}

const RESEND_URL = 'https://api.resend.com/emails';
const MITTENTE_PREDEFINITO = 'ScriptaScuola <no-reply@scriptascuola.it>';

/** Email "inviate" senza Resend: solo per sviluppo e test. */
export const emailInviate: Email[] = [];

export function svuotaEmailInviate() {
  emailInviate.length = 0;
}

export function usaResend(): boolean {
  return !!process.env.RESEND_API_KEY && process.env.NODE_ENV !== 'test';
}

/** Indirizzo del sito, senza barra finale, per i link nelle email. */
export function urlSito(): string {
  return (process.env.FRONTEND_ORIGIN || 'http://localhost:3000').replace(/\/+$/, '');
}

/**
 * Invia un'email. Lancia un errore se Resend la rifiuta:
 * chi la chiama decide se mostrarlo o solo scriverlo nel log.
 */
export async function inviaEmail(email: Email): Promise<void> {
  if (!usaResend()) {
    emailInviate.push(email);
    if (process.env.NODE_ENV !== 'test') {
      console.log(`[email simulata] a: ${email.a} | oggetto: ${email.oggetto}\n${email.testo}`);
    }
    return;
  }

  const risposta = await fetch(RESEND_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.RESEND_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: process.env.EMAIL_MITTENTE || MITTENTE_PREDEFINITO,
      to: [email.a],
      subject: email.oggetto,
      text: email.testo,
      html: email.html,
    }),
    signal: AbortSignal.timeout(10_000),
  });

  if (!risposta.ok) {
    const dettaglio = await risposta.text().catch(() => '');
    throw new Error(`Invio email non riuscito (Resend ${risposta.status}): ${dettaglio.slice(0, 300)}`);
  }
}

/**
 * Invia senza far attendere né fallire la richiesta HTTP: un problema con Resend
 * finisce solo nel log. Usato dove la risposta deve essere sempre la stessa
 * (per esempio "password dimenticata", per non rivelare quali email esistono).
 */
export function inviaEmailInBackground(email: Email): void {
  inviaEmail(email).catch((err) => {
    console.error('Errore durante l\'invio di un\'email:', err);
  });
}
