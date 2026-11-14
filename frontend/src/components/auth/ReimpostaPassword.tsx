import { useState } from "react";
import { Link, useSearchParams } from "react-router-dom";
import { useForm } from "@tanstack/react-form";
import { HugeiconsIcon } from "@hugeicons/react";
import { KeyIcon } from "@hugeicons/core-free-icons";
import { reimpostaPasswordFormSchema } from "@shared/validation";
import { Field, FieldContent, FieldGroup, FieldLabel } from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import Footer from "@/components/layout/Footer";
import { api } from "@/lib/api.js";
import { formatError } from "@/lib/utils";

const FORMATO_TOKEN = /^[a-f0-9]{64}$/;

/**
 * Pagina aperta dal link "Reimposta la password" ricevuto via email (solo amministratori).
 * Il token è monouso e vale 60 minuti; dopo il cambio tutte le sessioni aperte vengono chiuse.
 */
export default function ReimpostaPassword() {
  const [searchParams] = useSearchParams();
  const token = searchParams.get("token") ?? "";
  const tokenValido = FORMATO_TOKEN.test(token);
  const [completata, setCompletata] = useState(false);
  const [serverError, setServerError] = useState<string | null>(null);

  const form = useForm({
    defaultValues: { newPassword: "", confirmPassword: "" },
    validators: { onChange: reimpostaPasswordFormSchema },
    onSubmit: async ({ value }) => {
      setServerError(null);
      try {
        await api.post("/auth/reimposta-password", { token, newPassword: value.newPassword });
        setCompletata(true);
      } catch (err) {
        setServerError(formatError(err, "Non è stato possibile reimpostare la password. Riprova più tardi."));
      }
    },
  });

  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-primary px-4 pt-10 pb-16">
      <div className="bg-secondary rounded-lg p-6 w-full max-w-xl">
        <div className="flex items-center justify-center w-20 h-20 rounded-full bg-primary p-2 mx-auto">
          <HugeiconsIcon icon={KeyIcon} strokeWidth={2} className="text-secondary w-12 h-12" />
        </div>
        <h1 className="text-2xl font-semibold text-primary text-center mt-4">Reimposta la password</h1>

        {!tokenValido ? (
          <div className="mt-6 space-y-4" data-testid="link-non-valido">
            <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded">
              Il link non è valido. Copia l'indirizzo completo dall'email oppure chiedi un nuovo link.
            </div>
            <Button asChild className="w-full rounded-none">
              <Link to="/password-dimenticata">Chiedi un nuovo link</Link>
            </Button>
          </div>
        ) : completata ? (
          <div className="mt-6 space-y-4" data-testid="password-reimpostata">
            <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-3 rounded">
              Password aggiornata. Per sicurezza abbiamo chiuso tutte le sessioni aperte: accedi con la nuova password.
            </div>
            <Button asChild className="w-full rounded-none">
              <Link to="/login">Vai al login</Link>
            </Button>
          </div>
        ) : (
          <>
            <p className="text-sm text-center text-muted-foreground mt-1">
              Scegli una nuova password di almeno 8 caratteri.
            </p>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                form.handleSubmit();
              }}
              className="mt-6"
            >
              <FieldGroup>
                <form.Field name="newPassword">
                  {(field) => (
                    <Field>
                      <FieldLabel htmlFor="new-password">Nuova password</FieldLabel>
                      <FieldContent>
                        <Input
                          id="new-password"
                          type="password"
                          autoComplete="new-password"
                          className="rounded-none h-12"
                          value={field.state.value}
                          onChange={(e) => field.handleChange(e.target.value)}
                          onBlur={field.handleBlur}
                        />
                        {field.state.meta.isTouched && field.state.meta.errors.map((error, index) => (
                          <span key={index} className="text-red-500 text-xs">{error?.message}</span>
                        ))}
                      </FieldContent>
                    </Field>
                  )}
                </form.Field>

                <form.Field name="confirmPassword">
                  {(field) => (
                    <Field>
                      <FieldLabel htmlFor="confirm-password">Conferma nuova password</FieldLabel>
                      <FieldContent>
                        <Input
                          id="confirm-password"
                          type="password"
                          autoComplete="new-password"
                          className="rounded-none h-12"
                          value={field.state.value}
                          onChange={(e) => field.handleChange(e.target.value)}
                          onBlur={field.handleBlur}
                        />
                        {field.state.meta.isTouched && field.state.meta.errors.map((error, index) => (
                          <span key={index} className="text-red-500 text-xs">{error?.message}</span>
                        ))}
                      </FieldContent>
                    </Field>
                  )}
                </form.Field>

                {serverError && (
                  <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded" role="alert">
                    {serverError}
                  </div>
                )}

                <div className="pt-2 flex flex-col items-center">
                  <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting, state.isDirty]}>
                    {([canSubmit, isSubmitting, isDirty]) => (
                      <Button type="submit" className="w-full rounded-none" disabled={!canSubmit || isSubmitting || !isDirty}>
                        {isSubmitting ? "Salvataggio in corso..." : "Salva la nuova password"}
                      </Button>
                    )}
                  </form.Subscribe>
                  <Link to="/login" className="text-sm text-muted-foreground underline hover:text-primary mt-3">
                    Torna al login
                  </Link>
                </div>
              </FieldGroup>
            </form>
          </>
        )}
      </div>
      <Footer />
    </div>
  );
}
