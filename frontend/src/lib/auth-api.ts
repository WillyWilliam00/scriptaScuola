import { api, isSessionExpiredError, refreshAccessToken } from './api.js';
import { useAuthStore } from '../store/auth-store.js';
import type { LoginData, RegisterData, InsertIstituto, ChangePassword } from '../../../shared/validation.js';
import type { LoginResponse } from '../../../shared/types.js';
import type { QueryClient } from '@tanstack/react-query';

/**
 * Funzioni API per autenticazione
 * 
 * Spiegazione:
 * - Queste funzioni usano l'istanza Axios configurata (api)
 * - Gli interceptors gestiscono automaticamente token e refresh
 * - Dopo operazioni riuscite, aggiornano lo store Zustand
 * - Usano i tipi da shared per coerenza con il backend
 */

/**
 * Login: autentica l'utente con email/username e password
 * 
 * Spiegazione:
 * - Chiama /auth/login con identifier (email o username) e password
 * - Se ha successo, salva access token e dati utente nello store
 * - Il refresh token non compare nella risposta: il server lo mette in un cookie httpOnly
 * - Decodifica il JWT per estrarre il payload
 * - Restituisce la risposta per eventuale uso nel componente
 * 
 * @param identifier - Email o username dell'utente
 * @param password - Password dell'utente
 * @returns La risposta del login con token e dati utente
 */
export async function login(identifier: string, password: string): Promise<LoginResponse> {
  const loginData: LoginData = { identifier, password };
  
  const response = await api.post<LoginResponse>('/auth/login', loginData);
  // Salva nello store
  useAuthStore.getState().login(response.data);
  
  return response.data;
}

/**
 * Register: registra un nuovo istituto e crea l'utente admin
 * 
 * Spiegazione:
 * - Chiama /auth/setup-scuola con dati istituto e utente admin
 * - Non fa login automatico (l'utente deve fare login dopo)
 * - Restituisce la risposta per mostrare messaggio di successo
 * 
 * @param istituto - Dati dell'istituto da creare
 * @param utente - Dati dell'utente admin da creare
 * @returns La risposta della registrazione
 */
export async function register(
  istituto: InsertIstituto,
  utente: RegisterData
): Promise<{ message: string; istituto: unknown; utente: unknown }> {
  const response = await api.post('/auth/setup-scuola', {
    istituto,
    utente,
  });
  
  return response.data;
}

/**
 * Logout: revoca il refresh token e fa logout
 * 
 * Spiegazione:
 * - Chiama /auth/logout: il browser allega il cookie httpOnly con il refresh token,
 *   il server lo revoca nel database e cancella il cookie
 * - Pulisce lo store Zustand (logout)
 * - Pulisce la cache delle query React Query se queryClient è fornito
 * 
 * Nota: 
 * - Il try-catch è necessario perché anche se la chiamata API fallisce (es. server down),
 *   dobbiamo comunque pulire lo store locale per permettere all'utente di fare logout
 *   e accedere alla pagina di login (altrimenti PublicRoute lo bloccherebbe)
 * 
 * @param queryClient - Opzionale: QueryClient di React Query per pulire la cache
 * @returns La risposta del logout
 */
export async function logout(queryClient?: QueryClient): Promise<{ message: string }> {
  try {
    await api.post('/auth/logout');
  } catch (error) {
    // Anche se il logout fallisce (es. server down, rete assente), 
    // puliamo comunque lo store locale per permettere all'utente di accedere al login
    // Altrimenti PublicRoute lo bloccherebbe perché isAuthenticated sarebbe ancora true
    console.error('Errore durante logout API:', error);
  }
  
  // Pulisci lo store (logout locale) - sempre, anche se l'API fallisce
  useAuthStore.getState().logout();
  
  // Pulisci la cache delle query se fornito
  if (queryClient) {
    queryClient.clear();
  }
  
  return { message: 'Logout effettuato con successo' };
}

/**
 * Cambia la password dell'utente autenticato.
 * Richiede la password attuale: il backend verifica hash e rifiuta se la nuova è uguale.
 *
 * Il backend revoca tutte le sessioni dell'utente e ne apre una nuova per questo dispositivo,
 * sostituendo il cookie: qui non c'è nulla da salvare.
 */
export async function changePassword(
  currentPassword: string,
  newPassword: string
): Promise<{ message: string }> {
  const payload: ChangePassword = { currentPassword, newPassword };
  const response = await api.post<{ message: string }>('/auth/change-password', payload);
  // Una password scelta dall'utente sostituisce sempre quella temporanea
  useAuthStore.getState().setDeveCambiarePassword(false);
  return response.data;
}

/**
 * Ripristina la sessione all'avvio dell'app (es. dopo un ricaricamento della pagina).
 *
 * Spiegazione:
 * - L'access token vive solo in memoria, quindi dopo un ricaricamento non c'è più
 * - Se in localStorage risulta un utente collegato, chiediamo al server un nuovo access token:
 *   il browser allega da solo il cookie httpOnly con il refresh token
 * - Sessione non più valida (cookie assente, scaduto o revocato): logout locale
 * - Rete assente o backend non raggiungibile: l'utente resta collegato; alla prima richiesta
 *   l'interceptor di api.ts riproverà il rinnovo
 * - In ogni caso, alla fine l'app può uscire dallo stato di caricamento iniziale
 */
export async function restoreSession(): Promise<void> {
  const { utente, token, logout: logoutLocale, setInitialized } = useAuthStore.getState();
  try {
    if (utente && !token) {
      await refreshAccessToken();
    }
  } catch (error) {
    if (isSessionExpiredError(error)) {
      logoutLocale();
    } else {
      console.warn('Ripristino della sessione non riuscito, riproverò alla prima richiesta:', error);
    }
  } finally {
    setInitialized();
  }
}
