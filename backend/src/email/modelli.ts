import type { Email } from './index.js';
import { urlSito } from './index.js';

/** Testo inserito in HTML: niente tag dall'esterno (nomi di istituti, note, username). */
function escapeHtml(valore: string): string {
  return valore
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Impaginazione comune: una colonna, testo leggibile anche senza immagini. */
function layout(titolo: string, paragrafi: string[], bottone?: { testo: string; url: string }): string {
  const corpo = paragrafi.map((p) => `<p style="margin:0 0 16px;line-height:1.5">${p}</p>`).join('');
  const cta = bottone
    ? `<p style="margin:24px 0"><a href="${bottone.url}" style="background:#1c3a5f;color:#ffffff;padding:12px 20px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block">${escapeHtml(bottone.testo)}</a></p>
       <p style="margin:0 0 16px;font-size:13px;color:#5a6478">Se il pulsante non funziona, copia questo indirizzo nel browser:<br><span style="word-break:break-all">${bottone.url}</span></p>`
    : '';
  return `<!doctype html><html lang="it"><body style="margin:0;padding:24px;background:#f4f5f7;font-family:Arial,Helvetica,sans-serif;color:#172033">
<div style="max-width:520px;margin:0 auto;background:#ffffff;border-radius:8px;padding:28px">
<p style="margin:0 0 20px;font-weight:700;color:#1c3a5f;font-size:18px">ScriptaScuola</p>
<h1 style="margin:0 0 16px;font-size:20px">${escapeHtml(titolo)}</h1>
${corpo}${cta}
<p style="margin:24px 0 0;font-size:12px;color:#5a6478">Hai ricevuto questa email perché il tuo indirizzo è registrato su ScriptaScuola. Se non riguarda te, puoi ignorarla.</p>
</div></body></html>`;
}

export function emailVerificaIndirizzo(a: string, token: string, ore: number): Email {
  const url = `${urlSito()}/verifica-email?token=${token}`;
  return {
    a,
    oggetto: 'Conferma il tuo indirizzo email su ScriptaScuola',
    testo: `Per confermare il tuo indirizzo email apri questo link entro ${ore} ore:\n${url}\n\nSe non hai creato tu l'account, ignora questa email.`,
    html: layout('Conferma il tuo indirizzo email', [
      'Conferma che questo indirizzo è tuo: servirà per recuperare la password se la dimentichi.',
      `Il link vale ${ore} ore.`,
    ], { testo: 'Conferma email', url }),
  };
}

export function emailReimpostaPassword(a: string, token: string, minuti: number): Email {
  const url = `${urlSito()}/reimposta-password?token=${token}`;
  return {
    a,
    oggetto: 'Reimposta la tua password di ScriptaScuola',
    testo: `Hai chiesto di reimpostare la password. Apri questo link entro ${minuti} minuti:\n${url}\n\nSe non sei stato tu, ignora questa email: la password attuale resta valida.`,
    html: layout('Reimposta la password', [
      'Hai chiesto di reimpostare la password del tuo account amministratore.',
      `Il link vale ${minuti} minuti e si può usare una sola volta.`,
      'Se non sei stato tu, ignora questa email: la password attuale resta valida.',
    ], { testo: 'Scegli una nuova password', url }),
  };
}

export function emailRichiestaResetCollaboratore(a: string, username: string, nota: string | null): Email {
  const url = `${urlSito()}/gestione-utenze`;
  const notaTesto = nota ? `\nNota: "${nota}"` : '';
  return {
    a,
    oggetto: `${username} ha chiesto una nuova password`,
    testo: `Il collaboratore ${username} ha dimenticato la password e chiede una password temporanea.${notaTesto}\n\nPuoi generarla da Gestione utenze:\n${url}`,
    html: layout('Richiesta di nuova password', [
      `Il collaboratore <strong>${escapeHtml(username)}</strong> ha dimenticato la password e chiede una password temporanea.`,
      ...(nota ? [`Nota: <em>“${escapeHtml(nota)}”</em>`] : []),
      'Puoi generarla da Gestione utenze e consegnarla alla persona.',
    ], { testo: 'Apri Gestione utenze', url }),
  };
}
