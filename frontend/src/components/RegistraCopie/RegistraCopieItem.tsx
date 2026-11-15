import { useMemo, useState } from "react";
import { useForm } from "@tanstack/react-form";
import { Plus } from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";
import { InputGroup, InputGroupAddon, InputGroupInput } from "@/components/ui/input-group";
import { Item, ItemContent, ItemDescription, ItemTitle } from "@/components/ui/item";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import { useCreateRegistrazione } from "@/hooks/use-registrazioni.js";
import { useAuthStore } from "@/store/auth-store.js";
import { createRegistrazioneFormSchema } from "../../../../shared/validation.js";
import type { DocenteConRegistrazioni } from "../../../../shared/types.js";
import { cn, formatError } from "@/lib/utils.js";

export interface RegistraCopieItemProps {
  docente: DocenteConRegistrazioni;
  isSelected: boolean;
  setSelectedDocenteId: (id: number | null) => void;
  onOpenChange: (open: boolean) => void;
}

export default function RegistraCopieItem({
  docente,
  setSelectedDocenteId,
  isSelected,
  onOpenChange,
}: RegistraCopieItemProps) {
  const createRegistrazione = useCreateRegistrazione();
  const utenteId = useAuthStore((s) => s.utente?.id);
  const [serverError, setServerError] = useState<string | null>(null);
  const isOverRegistrazioni = docente.copieRimanenti === 0;
  const missing5Percent = Math.max(1, (docente.limiteCopie * 0.05)) >= docente.copieRimanenti;

  // Schema Zod per il form con i limiti dinamici del docente
  const formSchema = useMemo(
    () => createRegistrazioneFormSchema(docente.limiteCopie, docente.copieRimanenti),
    [docente.limiteCopie, docente.copieRimanenti]
  );

  const form = useForm({
    defaultValues: {
      copie: "",
      note: "",
    },
    validators: {
      onChange: formSchema,
    },
    onSubmit: async ({ value }) => {
      setServerError(null); // Reset errore server prima di ogni submit

      // Validazione utenteId (non può essere validata con Zod perché viene dallo store)
      if (!utenteId) {
        setServerError("Sessione non valida. Effettua di nuovo l'accesso.");
        return;
      }

      // Convertiamo la stringa in numero per la mutation
      const copieNumero = Number(value.copie);

      // Costruiamo l'oggetto completo per la mutation
      createRegistrazione.mutate(
        {
          docenteId: docente.id,
          copieEffettuate: copieNumero,
          note: value.note.trim() || undefined,
        },
        {
          onSuccess: () => {
            // Reset del form solo se la mutation ha successo
            form.reset();
            setSelectedDocenteId(null);
          },
          onError: (err) => {
            setServerError(formatError(err, "Errore durante la registrazione."));
          },
        }
      );
    },
  });

  return (
    <Dialog open={isSelected} onOpenChange={onOpenChange}>
      <DialogTrigger asChild>
        <button
          type="button"
          className="cursor-pointer w-full h-full text-left"
          disabled={isOverRegistrazioni}
        >
          <Item
            variant="outline"
            className={cn(
              "h-full hover:bg-muted transition-all duration-200 ease-linear",
              isSelected && "ring-2 ring-primary",
              isOverRegistrazioni && "opacity-50 bg-red-200/50 cursor-not-allowed hover:bg-red-200/50 border-red-500",
              missing5Percent && !isOverRegistrazioni && "bg-yellow-100/50 hover:bg-yellow-200/50 border-yellow-500"
            )}
          >
            <ItemContent>
              <ItemTitle className={cn(missing5Percent && !isOverRegistrazioni && "text-yellow-500" , isOverRegistrazioni && "text-red-500", 'md:text-lg ')}>
                {docente.nome} {docente.cognome}
              </ItemTitle>
              <ItemDescription className={cn(missing5Percent && !isOverRegistrazioni && "text-yellow-500" , isOverRegistrazioni && "text-red-500", 'md:text-base ')}>
                Copie: {docente.copieEffettuate}/{docente.limiteCopie}
              </ItemDescription>
            </ItemContent>
            <ItemContent>
              <ItemDescription className={cn("md:text-base text-sm", isOverRegistrazioni && "text-red-500", missing5Percent && !isOverRegistrazioni && "text-yellow-500")}>
                Copie Rimanenti: {docente.copieRimanenti}
              </ItemDescription>
            </ItemContent>
          </Item>
        </button>
      </DialogTrigger>

      <DialogContent className="sm:max-w-lg">
        <DialogHeader>
          <DialogTitle>
            Registra copie — {docente.nome} {docente.cognome}
          </DialogTitle>
          <DialogDescription>
            Copie effettuate: {docente.copieEffettuate}/{docente.limiteCopie} · Rimaste:{" "}
            {docente.copieRimanenti}
          </DialogDescription>
        </DialogHeader>

        <form
          onSubmit={(e) => {
            e.preventDefault();
            form.handleSubmit();
          }}
          className="flex flex-col gap-3"
        >
          <div className="flex gap-2 flex-wrap">
            <form.Field name="copie">
              {(field) => (
                <div className="flex-1 min-w-55">
                  <InputGroup>
                    <InputGroupInput
                      type="number"
                      min={1}
                      max={docente.limiteCopie}
                      placeholder="Copie effettuate"
                      value={field.state.value}
                      onChange={(e) => field.handleChange(e.target.value)}
                      onBlur={field.handleBlur}
                    />
                    <InputGroupAddon align="inline-end">
                      /{docente.limiteCopie}
                    </InputGroupAddon>
                  </InputGroup>
                  {field.state.meta.errors.length > 0 && (
                    <p className="text-sm text-destructive mt-1">
                      {field.state.meta.errors[0]?.message}
                    </p>
                  )}
                </div>
              )}
            </form.Field>

            <form.Field name="note">
              {(field) => (
                <div className="flex-1 min-w-55">
                  <InputGroup>
                    <InputGroupInput
                      type="text"
                      placeholder="Note (opzionale)"
                      value={field.state.value}
                      onChange={(e) => field.handleChange(e.target.value)}
                      onBlur={field.handleBlur}
                    />
                  </InputGroup>
                </div>
              )}
            </form.Field>
          </div>

          {serverError && (
            <p className="text-sm text-destructive">{serverError}</p>
          )}

          <div className="flex justify-end">
            <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting]}>
              {([canSubmit, isSubmitting]) => (
                <Button
                  type="submit"
                  variant="default"
                  disabled={!canSubmit || isSubmitting || createRegistrazione.isPending}
                >
                  <HugeiconsIcon icon={Plus} strokeWidth={2} />
                  <span className="ml-2">
                    {createRegistrazione.isPending ? "Salvataggio..." : "Registra"}
                  </span>
                </Button>
              )}
            </form.Subscribe>
          </div>
        </form>
      </DialogContent>
    </Dialog>
  );
}
