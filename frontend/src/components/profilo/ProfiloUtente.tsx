
import { useAuthStore } from "@/store/auth-store";

import { Card, CardHeader, CardTitle, CardDescription, CardContent } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import { Button } from '@/components/ui/button';
import { useEffect, useState } from 'react';
import { AlertDialog, AlertDialogCancel, AlertDialogFooter } from '@/components/ui/alert-dialog.js';
import { AlertDialogContent, AlertDialogDescription, AlertDialogHeader, AlertDialogMedia, AlertDialogTitle } from '@/components/ui/alert-dialog';
import { HugeiconsIcon } from '@hugeicons/react';
import { DeleteIcon } from '@hugeicons/core-free-icons';
import { DialogContent, DialogDescription, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Dialog } from '@/components/ui/dialog';
import { useDeleteIstituto } from '@/hooks/use-istituto.js';
import { formatError } from '@/lib/utils';
import { useQueryClient } from '@tanstack/react-query';
import { logout } from '@/lib/auth-api';
import { useNavigate } from 'react-router-dom';
import { useDeleteUtente } from '@/hooks/use-utenti';
import CambiaPasswordCard from '@/components/profilo/CambiaPasswordCard';

type DeleteDialogType = 'account' | 'istituto' | null;

export default function ProfiloUtente() {
  const { utente, istituto } = useAuthStore();
  const [deleteDialogType, setDeleteDialogType] = useState<DeleteDialogType>(null);
  const [isLogoutDialogOpen, setIsLogoutDialogOpen] = useState(false);
  const firtsTowLetters = utente?.email?.slice(0, 2).toUpperCase() ?? utente?.username?.slice(0, 2).toUpperCase();
  const isAdmin = utente?.ruolo === 'admin';
  const userId = utente?.id;
  const [timeToLogout, setTimeToLogout] = useState(15);
  const queryClient = useQueryClient();
  const navigate = useNavigate();

  const openDialogLogout = () => {
    setTimeToLogout(5);
    setIsLogoutDialogOpen(true);
  };

  const { mutate: deleteIstituto, isPending: isDeleteIstitutoPending, isError: isDeleteIstitutoError, error: deleteIstitutoError, reset: resetDeleteIstituto } = useDeleteIstituto(openDialogLogout);
  const { mutate: deleteUtente, isPending: isDeleteUtentePending, isError: isDeleteUtenteError, error: deleteUtenteError, reset: resetDeleteUtente } = useDeleteUtente(openDialogLogout);

  const openDialogDeleteIstituto = () => {
    resetDeleteIstituto();
    setDeleteDialogType('istituto');
  };

  const openDialogDeleteAccount = () => {
    resetDeleteUtente();
    setDeleteDialogType('account');
  };

  const closeDeleteDialog = () => setDeleteDialogType(null);

  const handleDeleteConfirm = () => {
    if (deleteDialogType === 'istituto') {
      deleteIstituto(undefined, { onSuccess: closeDeleteDialog });
    } else if (deleteDialogType === 'account') {
      deleteUtente(userId ?? '', { onSuccess: closeDeleteDialog });
    }
  };

  const isDeletePending = isDeleteIstitutoPending || isDeleteUtentePending;
  const deleteError = deleteDialogType === 'istituto' ? deleteIstitutoError : deleteUtenteError;
  const isDeleteError = deleteDialogType === 'istituto' ? isDeleteIstitutoError : isDeleteUtenteError;

  useEffect(() => {
    if (!isLogoutDialogOpen ) return;
  
    const timer = setInterval(() => {
      setTimeToLogout((prev) => {
        if (prev <= 1) {
          clearInterval(timer);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
  
    return () => clearInterval(timer);
  }, [isLogoutDialogOpen]);

  useEffect(() => {
    if (timeToLogout === 0) {
      logout(queryClient)
      navigate('/login')
    }
  }, [timeToLogout]);

  const displayIdentifier = isAdmin ? utente?.email : utente?.username;

  return (
    <div className="container mx-auto p-6 max-w-2xl">
      <Card>
        <CardHeader>
          <CardTitle className="flex flex-col items-center mb-6 ">
            <p className="text-lg font-bold ">
              Il mio profilo
            </p>
            <p className="text-lg text-primary font-bold px-4 py-2 bg-primary/10 rounded-full w-fit">
              {firtsTowLetters}
            </p>
          </CardTitle>
          <CardDescription>Visualizza e gestisci le informazioni del tuo profilo</CardDescription>


        </CardHeader>

        <CardContent className="space-y-6">
          {/* Istituto */}
          <div className="space-y-2">
            <h3 className="text-sm font-medium text-muted-foreground">Istituto</h3>
            <div className="space-y-1">
              <p className="text-base font-medium">{istituto?.nome}</p>
              <p className="text-sm text-muted-foreground">Codice: {istituto?.codiceIstituto}</p>
            </div>
          </div>

          <Separator />

          {/* Identificativo */}
          <div className="space-y-2">
            <h3 className="text-sm font-medium text-muted-foreground">
              {isAdmin ? 'Email' : 'Username'}
            </h3>
            <p className="text-base">{displayIdentifier}</p>
          </div>

          <Separator />

          {/* Ruolo */}
          <div className="space-y-2">
            <h3 className="text-sm font-medium text-muted-foreground">Ruolo</h3>
            <Badge variant={isAdmin ? 'default' : 'secondary'}>
              {isAdmin ? 'Amministratore' : 'Collaboratore'}
            </Badge>
          </div>
        </CardContent>

      </Card>

      <CambiaPasswordCard />

      {isAdmin && (
        <>
        <div className="container mx-auto p-6 max-w-2xl bg-red-500/10 mt-6 rounded-lg">
          <div className=''>
            <p className='text-sm text-red-500 font-medium'>
              Elimina Account
            </p>
            <p className='text-sm text-muted-foreground mt-2'>
              Attenzione: questa azione eliminerà il tuo account e tutti i tuoi dati, compresi i docenti e gli utenti.
            </p>
          </div>
          <div className='mt-4'>
            <Button className="w-full bg-red-500 text-white hover:bg-red-500/80" onClick={openDialogDeleteAccount}>
              {isDeleteUtentePending ? 'Eliminazione in corso...' : 'Elimina Account'}
            </Button>
          </div>
        </div>
      
     


        <div className="container mx-auto p-6 max-w-2xl bg-red-500/10 mt-6 rounded-lg">
          <div className=''>
            <p className='text-sm text-red-500 font-medium'>
              Elimina Istituto
            </p>
            <p className='text-sm text-muted-foreground mt-2'>
              Attenzione: questa azione eliminerà l'istituto e tutti i suoi dati, compresi i docenti e gli utenti.
            </p>
          </div>
          <div className='mt-4'>
            <Button className="w-full bg-red-500 text-white hover:bg-red-500/80" onClick={openDialogDeleteIstituto}>
              Elimina Istituto
            </Button>
          </div>
        </div>
        </>
      )}
      

      <AlertDialog open={deleteDialogType !== null} onOpenChange={(open) => !open && closeDeleteDialog()}>
        <AlertDialogContent size="sm">
          <AlertDialogHeader>
            <AlertDialogMedia className="bg-destructive/10 text-destructive dark:bg-destructive/20 dark:text-destructive">
              <HugeiconsIcon icon={DeleteIcon} strokeWidth={2} />
            </AlertDialogMedia>
            <AlertDialogTitle>
              {deleteDialogType === 'istituto'
                ? `Eliminare l'istituto ${istituto?.nome}?`
                : 'Eliminare il tuo account?'}
            </AlertDialogTitle>
            <AlertDialogDescription>
              {deleteDialogType === 'istituto'
                ? "Questa azione eliminerà l'istituto e tutti i suoi dati, compresi i docenti e gli utenti. Non può essere annullata."
                : "Questa azione eliminerà il tuo account e tutti i tuoi dati. Non può essere annullata."}
            </AlertDialogDescription>
            {isDeleteError && deleteError && (
              <p className="text-sm text-destructive px-6" role="alert">
                {formatError(deleteError, deleteDialogType === 'istituto' ? "Errore durante l'eliminazione dell'istituto" : "Errore durante l'eliminazione del tuo account")}
              </p>
            )}
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={isDeletePending} variant="outline" onClick={closeDeleteDialog}>
              Annulla
            </AlertDialogCancel>
            <Button
              variant="destructive"
              className="w-full"
              onClick={handleDeleteConfirm}
              disabled={isDeletePending}
            >
              {isDeletePending ? 'Eliminazione in corso...' : 'Elimina'}
            </Button>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      <Dialog open={isLogoutDialogOpen} onOpenChange={setIsLogoutDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Disconnessione in corso <div className='w-full h-1 bg-primary/10 rounded-full animate-pulse' /></DialogTitle>
          </DialogHeader>
          <DialogDescription>
            <div className='flex flex-col items-center justify-center'>
              <p className='text-sm text-muted-foreground'>Verrai disconnesso tra:</p>
              <p className="text-sm text-muted-foreground">
                {timeToLogout} secondi
              </p>
            </div>
          </DialogDescription>
        </DialogContent>
      </Dialog>

    </div>
  );
}
