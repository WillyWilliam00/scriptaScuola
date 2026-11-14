import { useEffect, useState } from "react";
import { useMutation, useQuery } from "@tanstack/react-query";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { api } from "@/lib/api.js";
import { formatError } from "@/lib/utils";
import { useAuthStore } from "@/store/auth-store";
import type { LoginResponse } from "@shared/types";

/**
 * Stato dell'email di un amministratore nel profilo: confermata o no,
 * con il pulsante per ricevere di nuovo il link di conferma.
 * I dati arrivano da GET /auth/me, così il badge si aggiorna anche se l'email
 * è stata confermata in un'altra scheda.
 */
export default function StatoEmailAdmin() {
  const utente = useAuthStore((s) => s.utente);
  const setUser = useAuthStore((s) => s.setUser);
  const [inviata, setInviata] = useState(false);

  const { data: me } = useQuery({
    queryKey: ["auth", "me"],
    queryFn: async () => (await api.get<LoginResponse["utente"]>("/auth/me")).data,
    refetchOnWindowFocus: true,
  });

  useEffect(() => {
    if (!me) return;
    const attuale = useAuthStore.getState().utente;
    if (attuale && (attuale.emailVerificata !== me.emailVerificata || attuale.haPassword !== me.haPassword)) {
      setUser({ ...attuale, emailVerificata: me.emailVerificata, haPassword: me.haPassword });
    }
  }, [me, setUser]);

  const invia = useMutation({
    mutationFn: async () => (await api.post<{ message: string }>("/auth/invia-verifica-email")).data,
    onSuccess: () => setInviata(true),
  });

  const verificata = me?.emailVerificata ?? utente?.emailVerificata ?? false;

  return (
    <div className="space-y-2" data-testid="stato-email">
      <div className="flex flex-wrap items-center gap-2">
        <p className="text-base">{utente?.email}</p>
        {verificata ? (
          <Badge className="bg-green-600 text-white">Confermata</Badge>
        ) : (
          <Badge variant="outline" className="border-amber-500 text-amber-700">Da confermare</Badge>
        )}
      </div>
      {!verificata && (
        <div className="space-y-2">
          <p className="text-sm text-muted-foreground">
            Conferma l'email per poter recuperare la password da solo se la dimentichi.
            Senza conferma non ricevi nemmeno gli avvisi delle richieste di nuova password dei collaboratori.
          </p>
          {inviata ? (
            <p className="text-sm text-green-700" role="status">
              Ti abbiamo inviato il link: controlla la posta (anche lo spam). Vale 48 ore.
            </p>
          ) : (
            <Button variant="outline" size="sm" onClick={() => invia.mutate()} disabled={invia.isPending}>
              {invia.isPending ? "Invio in corso..." : "Invia email di conferma"}
            </Button>
          )}
          {invia.isError && (
            <p className="text-sm text-destructive" role="alert">
              {formatError(invia.error, "Non è stato possibile inviare l'email. Riprova più tardi.")}
            </p>
          )}
        </div>
      )}
    </div>
  );
}
