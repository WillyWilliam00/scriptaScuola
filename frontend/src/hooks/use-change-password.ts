import { useMutation } from '@tanstack/react-query';
import { changePassword } from '@/lib/auth-api';
import { toast } from 'sonner';
import { formatError } from '@/lib/utils.js';
import type { ChangePassword } from '@shared/validation';

export function useChangePassword() {
  return useMutation({
    mutationFn: async (data: ChangePassword) => {
      return changePassword(data.currentPassword, data.newPassword);
    },
    onSuccess: () => {
      toast.success('Password aggiornata con successo');
    },
    onError: (err) => {
      toast.error(formatError(err, 'Errore durante il cambio password.'));
    },
  });
}
