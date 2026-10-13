import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogClose, DialogFooter, DialogTrigger } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { useCreateUtente } from "@/hooks/use-utenti";
import { formatError } from "@/lib/utils.js";
import { HugeiconsIcon } from "@hugeicons/react";
import { Plus } from "@hugeicons/core-free-icons";
import { UtenteFormFields } from "./UtenteFormFields";
import { useUtenteForm } from "./utente-form";

interface NuovoUtenteModalProps {
    typeForm: "view" | "add" | "edit" | "delete" | null;
    onClose: () => void;
    onOpenAddForm: () => void;
}

export default function NuovoUtenteModal({ typeForm, onClose, onOpenAddForm }: NuovoUtenteModalProps) {
    const createUtente = useCreateUtente();

    const form = useUtenteForm({
        passwordRequired: true,
        onSubmit: (data) => {
            createUtente.mutate(data, {
                onSuccess: () => {
                    onClose();
                    createUtente.reset();
                    form.reset();
                }
            })
        }
    })

    return (
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
                        Inserisci ruolo, email o username e password.
                    </DialogDescription>
                </DialogHeader>
                <div className="space-y-4 py-4">
                    <UtenteFormFields
                        form={form}
                        ruoloDisabled={false}
                        passwordLabel="Password"
                        passwordRequired
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
    )
}
