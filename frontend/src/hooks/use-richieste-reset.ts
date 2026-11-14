import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query';
import { toast } from 'sonner';
import { api } from '@/lib/api.js';
import { formatError } from '@/lib/utils.js';
import { useAuthStore } from '@/store/auth-store.js';
import type { PasswordTemporaneaResponse, RichiestaResetPassword } from '@shared/types';

export const RICHIESTE_RESET_KEY = ['richieste-reset'] as const;

/**
 * Richieste di reset password aperte dei collaboratori dell'istituto.
 * Solo per gli admin; si aggiorna da sola ogni minuto (badge in sidebar e card in Gestione utenze).
 */
export function useRichiesteReset() {
  const isAdmin = useAuthStore((state) => state.jwtPayload?.ruolo === 'admin');
  return useQuery({
    queryKey: RICHIESTE_RESET_KEY,
    queryFn: async () => {
      const response = await api.get<{ data: RichiestaResetPassword[] }>('/utenti/richieste-reset');
      return response.data.data;
    },
    enabled: isAdmin,
    refetchInterval: 60_000,
    staleTime: 30_000,
  });
}

/** Genera una nuova password temporanea per un'utenza (la risposta va mostrata una sola volta). */
export function useGeneraPasswordTemporanea() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (utenteId: string) => {
      const response = await api.post<PasswordTemporaneaResponse>(`/utenti/${utenteId}/password-temporanea`);
      return response.data;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: RICHIESTE_RESET_KEY });
    },
    onError: (err) => {
      toast.error(formatError(err, 'Errore durante la generazione della password temporanea.'));
    },
  });
}

/** Chiude una richiesta di reset senza generare una password. */
export function useIgnoraRichiestaReset() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (richiestaId: string) => {
      await api.post(`/utenti/richieste-reset/${richiestaId}/ignora`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: RICHIESTE_RESET_KEY });
      toast.success('Richiesta ignorata');
    },
    onError: (err) => {
      toast.error(formatError(err, 'Errore durante l\'aggiornamento della richiesta.'));
    },
  });
}
