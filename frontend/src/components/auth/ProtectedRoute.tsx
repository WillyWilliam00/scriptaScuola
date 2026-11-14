import { Navigate } from 'react-router-dom';
import { useAuthStore } from "@/store/auth-store";
import type { ReactNode } from 'react';
import CambioPasswordObbligatorio from './CambioPasswordObbligatorio';

/**
 * Componente ProtectedRoute per proteggere route con autenticazione e ruolo
 * 
 * Spiegazione:
 * - Wrapper riusabile e dichiarativo per proteggere route
 * - Verifica se l'utente è autenticato usando authStore
 * - Opzionalmente verifica il ruolo (admin o collaboratore)
 * - Se non autenticato: redirect a /login
 * - Se ruolo non corrisponde: mostra errore 403
 * - Pattern standard in React Router
 * 
 * Vantaggi:
 * - Separazione responsabilità: logica auth isolata
 * - Facile da testare
 * - Riusabile: <ProtectedRoute><Component /></ProtectedRoute>
 */
interface ProtectedRouteProps {
  children: ReactNode;
  requiredRole?: 'admin' | 'collaboratore';
}

export default function ProtectedRoute({ children, requiredRole }: ProtectedRouteProps) {
  const { isAuthenticated, jwtPayload, utente } = useAuthStore();

  // Se non autenticato, redirect a login
  if (!isAuthenticated) {
    return <Navigate to="/login" replace />;
  }

  // Password temporanea: prima di tutto va scelta una password nuova
  if (utente?.deveCambiarePassword) {
    return <CambioPasswordObbligatorio />;
  }

  // Se richiesto un ruolo specifico, verifica che corrisponda
  if (requiredRole && jwtPayload?.ruolo !== requiredRole) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center">
          <h1 className="text-2xl font-bold text-red-600 mb-2">Accesso Negato</h1>
          <p className="text-muted-foreground">
            Non hai i permessi necessari per accedere a questa pagina.
          </p>
          <p className="text-sm text-muted-foreground mt-2">
            Ruolo richiesto: <strong>{requiredRole}</strong>
          </p>
        </div>
      </div>
    );
  }

  // Autenticato e con ruolo corretto, renderizza il children
  return <>{children}</>;
}
