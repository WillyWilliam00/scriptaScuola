import { useEffect, useRef, useState } from "react";
import { Link, useSearchParams } from "react-router-dom";
import { HugeiconsIcon } from "@hugeicons/react";
import { MailValidation01Icon } from "@hugeicons/core-free-icons";
import { Button } from "@/components/ui/button";
import Footer from "@/components/layout/Footer";
import { api } from "@/lib/api.js";
import { formatError } from "@/lib/utils";
import { useAuthStore } from "@/store/auth-store";

const FORMATO_TOKEN = /^[a-f0-9]{64}$/;

type Stato = { tipo: "in-corso" } | { tipo: "ok" } | { tipo: "errore"; messaggio: string };

/**
 * Pagina aperta dal link "Conferma email" ricevuto dagli amministratori.
 * Conferma l'indirizzo appena si apre; se l'utente è già connesso aggiorna anche lo stato locale.
 */
export default function VerificaEmail() {
  const [searchParams] = useSearchParams();
  const token = searchParams.get("token") ?? "";
  const tokenValido = FORMATO_TOKEN.test(token);
  const [stato, setStato] = useState<Stato>(
    tokenValido ? { tipo: "in-corso" } : { tipo: "errore", messaggio: "Il link non è valido. Copia l'indirizzo completo dall'email." },
  );
  const isAuthenticated = useAuthStore((s) => s.isAuthenticated);
  // Il token è monouso: in sviluppo StrictMode esegue l'effetto due volte, la seconda fallirebbe.
  const inviato = useRef(false);

  useEffect(() => {
    if (!tokenValido || inviato.current) return;
    inviato.current = true;
    api.post("/auth/verifica-email", { token })
      .then(() => {
        setStato({ tipo: "ok" });
        const { utente, setUser } = useAuthStore.getState();
        if (utente) setUser({ ...utente, emailVerificata: true });
      })
      .catch((err) => {
        setStato({ tipo: "errore", messaggio: formatError(err, "Non è stato possibile confermare l'email. Riprova più tardi.") });
      });
  }, [token, tokenValido]);

  const destinazione = isAuthenticated ? "/profilo" : "/login";

  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-primary px-4 pt-10 pb-16">
      <div className="bg-secondary rounded-lg p-6 w-full max-w-xl">
        <div className="flex items-center justify-center w-20 h-20 rounded-full bg-primary p-2 mx-auto">
          <HugeiconsIcon icon={MailValidation01Icon} strokeWidth={2} className="text-secondary w-12 h-12" />
        </div>
        <h1 className="text-2xl font-semibold text-primary text-center mt-4">Conferma email</h1>

        <div className="mt-6 space-y-4" data-testid={`verifica-${stato.tipo}`}>
          {stato.tipo === "in-corso" && (
            <p className="text-center text-muted-foreground">Conferma in corso...</p>
          )}
          {stato.tipo === "ok" && (
            <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-3 rounded">
              Indirizzo confermato. Se dimentichi la password potrai recuperarla da "Password dimenticata?".
            </div>
          )}
          {stato.tipo === "errore" && (
            <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded" role="alert">
              {stato.messaggio} Puoi chiedere un nuovo link dalla pagina Profilo.
            </div>
          )}
          {stato.tipo !== "in-corso" && (
            <Button asChild className="w-full rounded-none">
              <Link to={destinazione}>{isAuthenticated ? "Vai al profilo" : "Vai al login"}</Link>
            </Button>
          )}
        </div>
      </div>
      <Footer />
    </div>
  );
}
