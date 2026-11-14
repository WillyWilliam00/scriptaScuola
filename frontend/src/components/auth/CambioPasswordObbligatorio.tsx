import { useNavigate } from "react-router-dom";
import { useQueryClient } from "@tanstack/react-query";
import { HugeiconsIcon } from "@hugeicons/react";
import { KeyIcon } from "@hugeicons/core-free-icons";
import { Button } from "@/components/ui/button";
import { logout } from "@/lib/auth-api";
import CambiaPasswordCard from "@/components/profilo/CambiaPasswordCard";
import Footer from "@/components/layout/Footer";

/**
 * Mostrata al posto dell'app finché l'utente usa una password temporanea
 * (utenza appena creata o password reimpostata da un admin).
 * Il backend blocca comunque tutte le altre rotte con 403 CAMBIO_PASSWORD_OBBLIGATORIO.
 */
export default function CambioPasswordObbligatorio() {
  const navigate = useNavigate();
  const queryClient = useQueryClient();

  const esci = async () => {
    await logout(queryClient);
    navigate("/login");
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-primary px-4 pt-10 pb-16" data-testid="cambio-password-obbligatorio">
      <div className="bg-secondary rounded-lg p-6 w-full max-w-xl">
        <div className="flex items-center justify-center w-20 h-20 rounded-full bg-primary p-2 mx-auto">
          <HugeiconsIcon icon={KeyIcon} strokeWidth={2} className="text-secondary w-12 h-12" />
        </div>
        <h1 className="text-2xl font-semibold text-primary text-center mt-4">Scegli una nuova password</h1>
        <p className="text-sm text-muted-foreground text-center mt-1">
          Stai usando una password temporanea. Prima di continuare scegline una tua.
        </p>
        <CambiaPasswordCard
          titolo="Nuova password"
          descrizione="Nel primo campo scrivi la password temporanea che ti hanno dato gli amministratori."
          etichettaPasswordAttuale="Password temporanea"
        />
        <div className="text-center mt-4">
          <Button variant="link" onClick={esci}>Esci</Button>
        </div>
      </div>
      <Footer />
    </div>
  );
}
