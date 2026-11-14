import { useState } from "react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { useGeneraPasswordTemporanea, useIgnoraRichiestaReset, useRichiesteReset } from "@/hooks/use-richieste-reset";
import type { RichiestaResetPassword } from "@shared/types";
import { PasswordTemporaneaDialog, type PasswordTemporaneaInfo } from "./PasswordTemporaneaDialog";

function identificativo(richiesta: RichiestaResetPassword): string {
    return richiesta.username ?? richiesta.email ?? "utenza";
}

function formatData(iso: string): string {
    return new Date(iso).toLocaleString("it-IT", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" });
}

/**
 * Richieste di reset password aperte: compare in Gestione utenze solo se ce n'è almeno una.
 * Da qui l'admin genera la password temporanea o ignora la richiesta.
 */
export function RichiesteResetCard() {
    const { data: richieste } = useRichiesteReset();
    const genera = useGeneraPasswordTemporanea();
    const ignora = useIgnoraRichiestaReset();
    const [passwordInfo, setPasswordInfo] = useState<PasswordTemporaneaInfo | null>(null);

    if (!richieste || richieste.length === 0) {
        return <PasswordTemporaneaDialog info={passwordInfo} onClose={() => setPasswordInfo(null)} />;
    }

    const occupato = genera.isPending || ignora.isPending;

    return (
        <>
            <Card className="mt-4 border-amber-300 bg-amber-50/60" data-testid="richieste-reset">
                <CardHeader>
                    <CardTitle className="text-base">
                        {richieste.length === 1
                            ? "1 richiesta di nuova password"
                            : `${richieste.length} richieste di nuova password`}
                    </CardTitle>
                    <CardDescription>
                        Hanno dimenticato la password. Genera una password temporanea e consegnala alla persona.
                    </CardDescription>
                </CardHeader>
                <CardContent>
                    <ul className="divide-y">
                        {richieste.map((richiesta) => (
                            <li key={richiesta.id} className="flex flex-col gap-2 py-3 sm:flex-row sm:items-center sm:justify-between">
                                <div className="min-w-0">
                                    <p className="font-semibold break-all">{identificativo(richiesta)}</p>
                                    <p className="text-xs text-muted-foreground">Richiesta il {formatData(richiesta.createdAt)}</p>
                                    {richiesta.nota && (
                                        <p className="text-sm mt-1 italic break-words">“{richiesta.nota}”</p>
                                    )}
                                </div>
                                <div className="flex gap-2 shrink-0">
                                    <Button
                                        size="sm"
                                        disabled={occupato}
                                        onClick={() =>
                                            genera.mutate(richiesta.utenteId, {
                                                onSuccess: (res) =>
                                                    setPasswordInfo({ identificativo: identificativo(richiesta), ...res }),
                                            })
                                        }
                                    >
                                        Genera password temporanea
                                    </Button>
                                    <Button
                                        size="sm"
                                        variant="outline"
                                        disabled={occupato}
                                        onClick={() => ignora.mutate(richiesta.id)}
                                    >
                                        Ignora
                                    </Button>
                                </div>
                            </li>
                        ))}
                    </ul>
                </CardContent>
            </Card>
            <PasswordTemporaneaDialog info={passwordInfo} onClose={() => setPasswordInfo(null)} />
        </>
    );
}
