import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import type { LoginResponse, JwtPayload } from '../../../shared/types.js';
import { decodeJWT } from '../lib/jwt-utils.js';

/**
 * Store Zustand per gestione stato autenticazione
 * 
 * Spiegazione Zustand:
 * - Zustand è una libreria di state management più leggera di Redux
 * - Il middleware persist salva in localStorage solo i dati dell'utente e dell'istituto
 * - I token NON vengono mai salvati in localStorage:
 *   - l'access token (15 minuti) vive solo in memoria
 *   - il refresh token sta in un cookie httpOnly gestito dal server, invisibile a JavaScript
 * - All'avvio dell'app i dati utente vengono ripristinati da localStorage e la sessione
 *   viene riattivata chiedendo al server un nuovo access token (vedi restoreSession in lib/auth-api.ts)
 * - Pattern create<T>()() con doppie parentesi per TypeScript corretto
 * - Salvare sia dati utente che JWT payload permette accesso rapido senza decodificare ogni volta
 */

interface AuthState {
  // Access token: solo in memoria, mai in localStorage
  token: string | null;
  
  // Dati utente completi (da LoginResponse)
  utente: {
    id: string;
    email: string | null;
    username: string | null;
    ruolo: 'admin' | 'collaboratore';
  } | null;
  
  // Dati istituto (da LoginResponse)
  istituto: {
    id: number;
    nome: string;
    codiceIstituto: string;
  } | null;
  
  // JWT payload decodificato (per accesso rapido a userId, ruolo, istitutoId)
  jwtPayload: JwtPayload | null;
  
  // Computed: verifica se l'utente è autenticato
  isAuthenticated: boolean;

  // True finché la sessione non è stata ripristinata all'avvio (vedi restoreSession)
  isInitializing: boolean;
  
  // Azioni
  login: (response: LoginResponse) => void;
  logout: () => void;
  setToken: (token: string) => void;
  setInitialized: () => void;
  setUser: (utente: AuthState['utente']) => void;
}

export const useAuthStore = create<AuthState>()(
  persist(
    (set, get) => ({
      // Stato iniziale
      token: null,
      utente: null,
      istituto: null,
      jwtPayload: null,
      isAuthenticated: false,
      isInitializing: true,

      /**
       * Login: salva access token (in memoria), dati utente e payload del JWT
       * 
       * Spiegazione:
       * - Viene chiamato dopo un login riuscito
       * - Decodifica il JWT per estrarre il payload (userId, ruolo, istitutoId)
       * - In localStorage finiscono solo utente, istituto e payload (vedi partialize)
       * - Il refresh token non passa da qui: il server lo ha messo in un cookie httpOnly
       */
      login: (response: LoginResponse) => {
        const jwtPayload = decodeJWT(response.token);
        
        set({
          token: response.token,
          utente: response.utente,
          istituto: response.istituto,
          jwtPayload,
          isAuthenticated: true,
        });
      },

      /**
       * Logout: pulisce tutto lo stato
       * 
       * Spiegazione:
       * - Viene chiamato quando l'utente fa logout
       * - Pulisce tutti i dati dallo store
       * - I dati vengono automaticamente rimossi da localStorage grazie al middleware persist
       */
      logout: () => {
        set({
          token: null,
          utente: null,
          istituto: null,
          jwtPayload: null,
          isAuthenticated: false,
        });
      },

      /**
       * setToken: aggiorna l'access token dopo un rinnovo della sessione
       * 
       * Spiegazione:
       * - Viene chiamato quando il server rilascia un nuovo access token (refresh)
       * - isAuthenticated segue la stessa regola usata al ripristino: token e utente presenti
       */
      setToken: (token: string) => {
        set({
          token,
          jwtPayload: decodeJWT(token),
          isAuthenticated: !!token && !!get().utente,
        });
      },

      /** Segnala che il ripristino della sessione all'avvio è terminato. */
      setInitialized: () => {
        set({ isInitializing: false });
      },

      /**
       * setUser: aggiorna solo i dati utente (raro, ma utile se necessario)
       */
      setUser: (utente: AuthState['utente']) => {
        set({ utente });
      },
    }),
    {
      name: 'auth-storage', // Chiave in localStorage
      storage: createJSONStorage(() => localStorage),
      // Versione 1: i token non sono più persistiti.
      // migrate ripulisce i dati salvati dalle versioni precedenti, che contenevano token e refreshToken.
      version: 1,
      migrate: (persisted) => {
        const { utente, istituto, jwtPayload } = (persisted ?? {}) as Partial<AuthState>;
        return { utente: utente ?? null, istituto: istituto ?? null, jwtPayload: jwtPayload ?? null };
      },
      // Solo questi campi vengono persistiti: nessun token
      partialize: (state) => ({
        utente: state.utente,
        istituto: state.istituto,
        jwtPayload: state.jwtPayload,
      }),
      // Al ripristino da localStorage sappiamo solo chi era l'utente: lo consideriamo autenticato
      // e lasciamo isInitializing a true finché restoreSession non ha ottenuto un access token
      // (o ha scoperto che la sessione non è più valida).
      onRehydrateStorage: () => (state) => {
        if (state) {
          state.isAuthenticated = !!state.utente;
        }
      },
    }
  )
);
