import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogClose, DialogFooter } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import type { Utente } from "@shared/types";
import { formatError, utenteDisplayName } from "@/lib/utils";
import { useUpdateUtente } from "@/hooks/use-utenti";
import { useEffect, useState } from "react";
import { useGeneraPasswordTemporanea } from "@/hooks/use-richieste-reset";
import { PasswordTemporaneaDialog, type PasswordTemporaneaInfo } from "./PasswordTemporaneaDialog";
import { UtenteFormFields } from "./UtenteFormFields";
import { getUtenteFormValues, emptyUtenteFormValues, useUtenteForm } from "./utente-form";

interface EditUtenteModalProps {
    selectedUtente: Utente | null;
    typeForm: "edit" | "view" | "delete" | "add" | null;
    onClose: () => void;
}

export default function EditUtenteModal({ selectedUtente, typeForm, onClose }: EditUtenteModalProps) {
    const isAdmin = selectedUtente ? selectedUtente.ruolo === "admin" : false;
    const isOpen = typeForm === "edit" && selectedUtente !== null;

    const updateUtente = useUpdateUtente();
    const generaPassword = useGeneraPasswordTemporanea();
    const [passwordInfo, setPasswordInfo] = useState<PasswordTemporaneaInfo | null>(null);

    const handleGeneraPassword = () => {
        if (!selectedUtente) return;
        const identificativo = utenteDisplayName(selectedUtente);
        generaPassword.mutate(selectedUtente.id, {
            onSuccess: (res) => {
                setPasswordInfo({ identificativo, ...res });
                updateUtente.reset();
                form.reset(emptyUtenteFormValues);
                onClose();
            },
        });
    };

    const form = useUtenteForm({
        onSubmit: (data) => {
            if (!selectedUtente) return;
            updateUtente.mutate({ id: selectedUtente.id, data }, {
                onSuccess: () => {
                    onClose();
                    updateUtente.reset();
                    form.reset(emptyUtenteFormValues);
                }
            })
        }
    })

    useEffect(() => {
        if (isOpen && selectedUtente) {
            form.reset(getUtenteFormValues(selectedUtente));
        }
    }, [isOpen, selectedUtente, form]);

    return (
        <>
        <Dialog
            open={isOpen}
            onOpenChange={(open) => {
                if (!open) {
                    updateUtente.reset();
                    form.reset(emptyUtenteFormValues);
                    onClose();
                }
            }}
        >
            <DialogContent>
                <DialogHeader>
                    <DialogTitle>
                        Modifica utenza
                    </DialogTitle>
                    <DialogDescription>
                        Modifica i dati dell’utenza di {selectedUtente ? utenteDisplayName(selectedUtente) : '...'}.
                    </DialogDescription>
                </DialogHeader>
                <UtenteFormFields
                    form={form}
                    ruoloDisabled={isAdmin}
                />
                <div className="rounded-md border p-3 space-y-2">
                    <p className="text-sm font-medium">Password</p>
                    <p className="text-xs text-muted-foreground">
                        Se ha dimenticato la password, generane una temporanea: le sessioni aperte vengono chiuse
                        e al primo accesso dovrà sceglierne una nuova.
                    </p>
                    <Button
                        type="button"
                        variant="outline"
                        size="sm"
                        onClick={handleGeneraPassword}
                        disabled={generaPassword.isPending}
                    >
                        {generaPassword.isPending ? "Generazione..." : "Genera password temporanea"}
                    </Button>
                </div>
                {
                    updateUtente.isError && updateUtente.error && (
                        <span className="text-red-500 text-xs">
                            {formatError(updateUtente.error, "Errore durante la modifica dell'utenza.")}
                        </span>
                    )
                }

                <DialogFooter>
                    <DialogClose asChild>
                        <Button variant="outline" onClick={onClose}>
                            Chiudi
                        </Button>
                    </DialogClose>
                    <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting, state.isDirty]}>
                        {([canSubmit, isSubmitting, isDirty]) => (
                            <Button
                                variant="default"
                                type="submit"
                                onClick={() => form.handleSubmit()}
                                disabled={!canSubmit || isSubmitting || !isDirty || updateUtente.isPending}
                            >
                                {isSubmitting ? "Salvataggio in corso..." : "Salva"}
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
