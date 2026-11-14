import { useState } from "react";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";

export interface PasswordTemporaneaInfo {
    /** Email o username dell'utenza */
    identificativo: string;
    passwordTemporanea: string;
    scadeAt: string;
}

interface PasswordTemporaneaDialogProps {
    info: PasswordTemporaneaInfo | null;
    onClose: () => void;
}

function formatScadenza(iso: string): string {
    return new Date(iso).toLocaleString("it-IT", {
        weekday: "long",
        day: "numeric",
        month: "long",
        hour: "2-digit",
        minute: "2-digit",
    });
}

/**
 * Mostra la password temporanea appena generata. È l'unica volta in cui si vede:
 * il server salva solo l'hash, quindi chiusa la finestra non si può più recuperare.
 */
export function PasswordTemporaneaDialog({ info, onClose }: PasswordTemporaneaDialogProps) {
    const [copiata, setCopiata] = useState(false);

    const copia = async () => {
        if (!info) return;
        try {
            await navigator.clipboard.writeText(info.passwordTemporanea);
            setCopiata(true);
        } catch {
            setCopiata(false);
        }
    };

    return (
        <Dialog
            open={info !== null}
            onOpenChange={(open) => {
                if (!open) {
                    setCopiata(false);
                    onClose();
                }
            }}
        >
            <DialogContent>
                <DialogHeader>
                    <DialogTitle>Password temporanea</DialogTitle>
                    <DialogDescription>
                        Consegnala a <strong>{info?.identificativo}</strong>. Al primo accesso dovrà sceglierne una nuova.
                    </DialogDescription>
                </DialogHeader>
                {info && (
                    <div className="space-y-3 py-2">
                        <div className="flex items-center gap-2">
                            <code
                                className="flex-1 rounded-md border bg-muted px-3 py-2 text-center font-mono text-lg tracking-wider select-all"
                                data-testid="password-temporanea"
                            >
                                {info.passwordTemporanea}
                            </code>
                            <Button type="button" variant="outline" onClick={copia}>
                                {copiata ? "Copiata" : "Copia"}
                            </Button>
                        </div>
                        <p className="text-sm text-muted-foreground">
                            Vale fino a {formatScadenza(info.scadeAt)}. Dopo serve generarne una nuova.
                        </p>
                        <p className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-2">
                            Non la potrai rivedere: chiusa questa finestra, la piattaforma non la conserva.
                        </p>
                    </div>
                )}
                <DialogFooter>
                    <Button onClick={() => { setCopiata(false); onClose(); }}>Ho preso nota</Button>
                </DialogFooter>
            </DialogContent>
        </Dialog>
    );
}
