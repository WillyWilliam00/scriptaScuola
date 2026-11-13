export interface PaginationParams {
  page: number;
  limit: number;
}

export interface PaginationResponse<T> {
  data: T[];
  pagination: {
    page: number;
    pageSize: number;
    totalItems: number;
    totalPages: number;
    hasNextPage: boolean;
    hasPreviousPage: boolean;
  }
}

export type SortOrder = 'asc' | 'desc';
export interface DocentiFilters {
  nome?: string;
  cognome?: string;
}	
export interface DocentiSort {
  field: 'nome' | 'cognome' | 'limiteCopie' | 'copieEffettuate' | 'copieRimanenti' | 'createdAt' | 'updatedAt';
  order: SortOrder;
}
export interface UtentiFilters {
  email?: string;
  username?: string;
  ruolo?: 'admin' | 'collaboratore';
}
export interface UtentiSort {
  field: 'ruolo' | 'identificativo';
  order: SortOrder;
}
export interface RegistrazioniCopieSort {
  field: 'docenteId' | 'utenteId' | 'createdAt' | 'updatedAt' | 'docenteNome' | 'docenteCognome' | 'copieEffettuate' | 'utente';
  order: SortOrder;
}
export interface RegistrazioniCopieFilters {
  docenteId?: number;
  utenteId?: string;
}
// --- Istituto ---
export interface Istituto {
  id: number;
  nome: string;
  codiceIstituto: string;
  createdAt: Date;
  updatedAt: Date;
}

// --- Docente ---
export interface Docente {
  id: number;
  nome: string;
  cognome: string;
  limiteCopie: number;
  istitutoId: number;
  createdAt: Date;
  updatedAt: Date;
}

// --- Utente (base comune) ---
interface UtenteBase {
  id: string;
  istitutoId: number;
  createdAt: Date;
  updatedAt: Date;
}

// Admin: ha email, username opzionale
interface UtenteAdmin extends UtenteBase {
  ruolo: 'admin';
  email: string;
}

// Collaboratore: ha username, email opzionale
interface UtenteCollaboratore extends UtenteBase {
  ruolo: 'collaboratore';
  username: string;
}

export type Utente = UtenteAdmin | UtenteCollaboratore;

// --- Registrazione Copie ---
export interface Registrazione {
  id: number;
  docenteId: number;
  copieEffettuate: number;
  istitutoId: number;
  utenteId: string | null;
  note: string | null;
  createdAt: Date;
  updatedAt: Date;
}

// --- Registrazione Copie con Dettagli (con join) ---
export interface RegistrazioneConDettagli extends Registrazione {
  docenteNome: string | null;
  docenteCognome: string | null;
  utenteUsername: string | null;
  utenteEmail: string | null;
}

// --- Tipi composti ---
export interface DocenteConRegistrazioni extends Docente {
  copieEffettuate: number;    // Somma totale delle copie
  copieRimanenti: number; 
}


// --- JWT Payload ---
export interface JwtPayload {
  userId: string;
  ruolo: 'admin' | 'collaboratore';
  istitutoId: number;
  exp?: number;
}

// --- Risposta login API ---
export interface LoginResponse {
  token: string;
  refreshToken: string;
  utente: {
    id: string;
    email: string | null;
    username: string | null;
    ruolo: 'admin' | 'collaboratore';
  };
  istituto: {
    id: number;
    nome: string;
    codiceIstituto: string;
  };
}

/** Risposta del cambio password: le vecchie sessioni sono revocate, questo è il refresh token della sessione corrente. */
export interface ChangePasswordResponse {
  message: string;
  refreshToken: string;
}

export interface RefreshTokenResponse {
  token: string;
  refreshToken?: string; //per rotazione token
}


export interface DocentiPaginatedResponse extends PaginationResponse<DocenteConRegistrazioni> {}
export interface UtentiPaginatedResponse extends PaginationResponse<Utente> {}
export interface RegistrazioniCopiePaginatedResponse extends PaginationResponse<RegistrazioneConDettagli> {}