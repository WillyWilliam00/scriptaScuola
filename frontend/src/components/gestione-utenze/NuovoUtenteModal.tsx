import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogClose, DialogFooter, DialogTrigger } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { useCreateUtente } from "@/hooks/use-utenti";
import { formatError } from "@/lib/utils.js";
import { HugeiconsIcon } from "@hugeicons/react";
import { Plus } from "@hugeicons/core-free-icons";
import { UtenteFormFields } from "./UtenteFormFields";
import { useUtenteForm } from "./utente-form";
import { useState } from "react";
import { PasswordTemporaneaDialog, type PasswordTemporaneaInfo } from "./PasswordTemporaneaDialog";

interface NuovoUtenteModalProps {
    typeForm: "view" | "add" | "edit" | "delete" | null;
    onClose: () => void;
    onOpenAddForm: () => void;
}

export default function NuovoUtenteModal({ typeForm, onClose, onOpenAddForm }: NuovoUtenteModalProps) {
    const createUtente = useCreateUtente();
    const [passwordInfo, setPasswordInfo] = useState<PasswordTemporaneaInfo | null>(null);

    const form = useUtenteForm({
        onSubmit: (data) => {
            createUtente.mutate(data, {
                onSuccess: (res) => {
                    setPasswordInfo({
                        identificativo: data.ruolo === "admin" ? data.email : data.username,
                        passwordTemporanea: res.passwordTemporanea,
                        scadeAt: res.scadeAt,
                    });
                    onClose();
                    createUtente.reset();
                    form.reset();
                }
            })
        }
    })

    return (
        <>
        <Dialog
            open={typeForm === "add"}
            onOpenChange={(open) => {
                if (!open) {
                    onClose();
                    createUtente.reset();
                    form.reset();
                }
            }}
        >
            <DialogTrigger asChild>
                <Button variant="default" onClick={onOpenAddForm}>
                    Aggiungi utenza
                    <HugeiconsIcon icon={Plus} strokeWidth={2} className="size-4" />
                </Button>
            </DialogTrigger>
            <DialogContent>
                <DialogHeader>
                    <DialogTitle>
                        Aggiungi utenza
                    </DialogTitle>
                    <DialogDescription>
                        Inserisci ruolo ed email o username. La password temporanea la genera la piattaforma: te la mostriamo subito dopo.
                    </DialogDescription>
                </DialogHeader>
                <div className="space-y-4 py-4">
                    <UtenteFormFields
                        form={form}
                        ruoloDisabled={false}
                    />
                    {
                        createUtente.isError && createUtente.error && (
                            <span className="text-red-500 text-xs">
                                {formatError(createUtente.error, "Errore durante l'aggiunta dell'utenza.")}
                            </span>
                        )
                    }
                </div>
                <DialogFooter>
                    <DialogClose asChild>
                        <Button variant="outline">Chiudi</Button>
                    </DialogClose>
                    <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting, state.isDirty]}>
                        {([canSubmit, isSubmitting, isDirty]) => (
                            <Button
                                variant="default"
                                type="submit"
                                onClick={() => form.handleSubmit()}
                                disabled={!canSubmit || isSubmitting || !isDirty || createUtente.isPending}
                            >
                                {isSubmitting ? "Aggiungi in corso..." : "Aggiungi"}
                            </Button>
                        )}
                    </form.Subscribe>
                </DialogFooter>
            </DialogContent>
        </Dialog>
        <PasswordTemporaneaDialog info={passwordInfo} onClose={() => setPasswordInfo(null)} />
        </>
    )
}
