import axios, { type AxiosError, type InternalAxiosRequestConfig } from 'axios';
import { useAuthStore } from '../store/auth-store.js';
import type { AxiosResponse } from 'axios';
import { CAMBIO_PASSWORD_OBBLIGATORIO, type RefreshTokenResponse } from '../../../shared/types.js';

/**
 * Istanza Axios configurata con interceptors per gestione token e refresh automatico
 * 
 * Spiegazione Axios:
 * - Axios è un client HTTP che semplifica le chiamate API rispetto a fetch
 * - Gli interceptors permettono di modificare richieste/risposte automaticamente
 * - Request interceptor: aggiunge token a ogni richiesta senza doverlo specificare manualmente
 * - Response interceptor: intercetta errori 401 e gestisce refresh token automaticamente
 */


/**
 * Indirizzo base dell'API.
 *
 * - In produzione è sempre `/api` sullo stesso dominio del sito: Vercel inoltra le chiamate
 *   al backend (vedi la rewrite in vercel.json). Così il cookie di sessione è del nostro dominio
 *   e i browser non lo scartano come cookie di terze parti.
 * - In sviluppo è `/api` attraverso il proxy di Vite (vedi vite.config.ts), oppure il valore
 *   di VITE_API_BASE_URL se si vuole puntare a un altro backend.
 */
function getBaseUrl(): string {
  if (import.meta.env.DEV && import.meta.env.VITE_API_BASE_URL) {
    return import.meta.env.VITE_API_BASE_URL;
  }
  return '/api';
}

const BASE_URL = getBaseUrl();

// Creiamo l'istanza Axios
export const api = axios.create({
  baseURL: BASE_URL,
  // Invia e riceve i cookie anche quando in sviluppo l'API è su un'altra origine
  withCredentials: true,
  headers: {
    'Content-Type': 'application/json',
  },
});

// Rotte che non devono mai far partire un rinnovo della sessione:
// un 401 qui significa "credenziali sbagliate" o "sessione finita", non "access token scaduto".
const ROTTE_SENZA_RINNOVO = ['/auth/login', '/auth/setup-scuola', '/auth/refresh', '/auth/logout'];

// Rinnovo in corso: tutte le richieste che lo aspettano condividono la stessa promessa,
// così il refresh parte una volta sola anche se falliscono più richieste insieme.
let refreshPromise: Promise<string> | null = null;

/**
 * Chiede al server un nuovo access token.
 *
 * Il refresh token non viene inviato dal codice: sta nel cookie httpOnly e lo allega il browser.
 * Il server risponde con un nuovo access token e ruota il cookie.
 */
export function refreshAccessToken(): Promise<string> {
  if (!refreshPromise) {
    refreshPromise = axios
      .post<RefreshTokenResponse>(`${BASE_URL}/auth/refresh`, undefined, { withCredentials: true })
      .then((response) => {
        useAuthStore.getState().setToken(response.data.token);
        return response.data.token;
      })
      .finally(() => {
        refreshPromise = null;
      });
  }
  return refreshPromise;
}

/**
 * True se il rinnovo è fallito perché la sessione non è più valida (cookie assente, scaduto o revocato).
 * Errori di rete, timeout ed errori del server (5xx, backend che si sta riavviando) NON contano:
 * in quei casi l'utente resta collegato e potrà riprovare.
 */
export function isSessionExpiredError(error: unknown): boolean {
  if (isCambioPasswordObbligatorio(error)) return false;
  const status = (error as AxiosError)?.response?.status;
  return status === 401 || status === 403;
}

/** True se il server ha bloccato la richiesta perché va prima cambiata la password temporanea. */
export function isCambioPasswordObbligatorio(error: unknown): boolean {
  const response = (error as AxiosError)?.response;
  return response?.status === 403
    && (response.data as { code?: string } | undefined)?.code === CAMBIO_PASSWORD_OBBLIGATORIO;
}

/**
 * Request Interceptor: aggiunge il token di autenticazione a ogni richiesta
 * 
 * Spiegazione:
 * - Viene eseguito PRIMA di ogni richiesta HTTP
 * - Recupera il token dallo store Zustand
 * - Aggiunge l'header Authorization se il token è presente
 * - Non serve specificare il token manualmente in ogni chiamata API
 */
api.interceptors.request.use(
  (config: InternalAxiosRequestConfig) => {
    const token = useAuthStore.getState().token
    if(token) {
      config.headers.Authorization = `Bearer ${token}`;
    }
    return config;
  },
  (error) => {
    return Promise.reject(error);
  }
)

/**
 * Response Interceptor: gestisce errori 401 e rinnovo automatico della sessione
 * 
 * Spiegazione:
 * - Viene eseguito DOPO ogni risposta HTTP
 * - Se riceve 401 (access token scaduto o assente dopo un ricaricamento), chiede un nuovo token
 * - Se il rinnovo ha successo, ripete la richiesta originale con il nuovo token
 * - Se la sessione non è più valida, fa logout e redirect a login
 * - Evita che l'utente debba rifare login quando l'access token scade (15 minuti)
 */
api.interceptors.response.use(
  (response: AxiosResponse) => {
    return response;
  },
  async (error: AxiosError) => {
    const originalRequest = error.config as (InternalAxiosRequestConfig & { _retry?: boolean }) | undefined;

    // Password temporanea ancora da cambiare: ProtectedRoute mostra la pagina del cambio password
    if (isCambioPasswordObbligatorio(error)) {
      useAuthStore.getState().setDeveCambiarePassword(true);
      return Promise.reject(error);
    }

    // Se l'errore non è 401 o la richiesta è già stata ritentata, rifiuta
    if (!originalRequest || error.response?.status !== 401 || originalRequest._retry) {
      return Promise.reject(error);
    }

    // Login, registrazione, refresh e logout: il 401 va restituito al chiamante così com'è
    const url = originalRequest.url ?? '';
    if (ROTTE_SENZA_RINNOVO.some((rotta) => url.includes(rotta))) {
      return Promise.reject(error);
    }

    // Senza un utente collegato non c'è nessuna sessione da rinnovare
    if (!useAuthStore.getState().isAuthenticated) {
      return Promise.reject(error);
    }

    // Controlla se l'errore indica che l'utente è stato eliminato
    // In questo caso, logout immediato senza tentare refresh
    const errorMessage = (error.response?.data as { error?: string })?.error || '';
    const isUserDeleted = errorMessage.includes('Utente non trovato') || 
                          errorMessage.includes('account eliminato') ||
                          errorMessage.includes('non trovato');

    if (isUserDeleted) {
      // Utente eliminato: logout immediato e redirect
      useAuthStore.getState().logout();
      if (typeof window !== 'undefined') {
        window.location.href = '/login';
      }
      return Promise.reject(error);
    }

    originalRequest._retry = true;

    try {
      // Se un rinnovo è già in corso, questa richiesta aspetta lo stesso risultato
      const newToken = await refreshAccessToken();

      // Ripeti la richiesta originale con il nuovo token
      if (originalRequest.headers) {
        originalRequest.headers.Authorization = `Bearer ${newToken}`;
      }

      return api(originalRequest);
    } catch (refreshError) {
      // Rete assente, timeout o backend non raggiungibile: NON fare logout.
      // L'utente potrebbe avere una sessione valida ma problemi di connessione: può riprovare.
      if (!isSessionExpiredError(refreshError)) {
        console.warn('Rinnovo della sessione non riuscito, mantengo lo stato corrente:', refreshError);
        return Promise.reject(refreshError);
      }

      // Sessione non più valida (cookie scaduto o revocato, utente eliminato): logout
      console.error('Sessione scaduta o revocata:', refreshError);
      useAuthStore.getState().logout();
      
      // Redirect a login se siamo in un contesto browser
      if (typeof window !== 'undefined') {
        window.location.href = '/login';
      }

      return Promise.reject(refreshError);
    }
  }
);
