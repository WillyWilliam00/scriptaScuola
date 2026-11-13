import { useState } from "react";
import { Link } from "react-router-dom";
import { useForm } from "@tanstack/react-form";
import { HugeiconsIcon } from "@hugeicons/react";
import { KeyIcon } from "@hugeicons/core-free-icons";
import { richiestaResetFormSchema } from "@shared/validation";
import { Field, FieldContent, FieldGroup, FieldLabel } from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Button } from "@/components/ui/button";
import Footer from "@/components/layout/Footer";
import { api } from "@/lib/api.js";
import { formatError } from "@/lib/utils";

const NOTA_MAX = 200;

/**
 * Pagina "Password dimenticata?" per i collaboratori.
 * La richiesta arriva agli admin dell'istituto, che generano una password temporanea
 * e la consegnano di persona. Il server risponde sempre allo stesso modo,
 * anche se lo username non esiste.
 */
export default function PasswordDimenticata() {
  const [inviata, setInviata] = useState<string | null>(null);
  const [serverError, setServerError] = useState<string | null>(null);

  const form = useForm({
    defaultValues: { username: "", nota: "" },
    validators: { onChange: richiestaResetFormSchema },
    onSubmit: async ({ value }) => {
      setServerError(null);
      try {
        const res = await api.post<{ message: string }>("/auth/richiesta-reset", {
          username: value.username,
          nota: value.nota || undefined,
        });
        setInviata(res.data.message);
      } catch (err) {
        setServerError(formatError(err, "Non è stato possibile inviare la richiesta. Riprova più tardi."));
      }
    },
  });

  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-4 py-10">
      <div className="bg-secondary rounded-lg p-6 w-full max-w-xl">
        <div className="flex items-center justify-center w-20 h-20 rounded-full bg-primary p-2 mx-auto">
          <HugeiconsIcon icon={KeyIcon} strokeWidth={2} className="text-secondary w-12 h-12" />
        </div>
        <h1 className="text-2xl font-semibold text-primary text-center mt-4">Password dimenticata</h1>

        {inviata ? (
          <div className="mt-6 space-y-4" data-testid="richiesta-inviata">
            <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-3 rounded">
              {inviata}
            </div>
            <p className="text-sm text-muted-foreground text-center">
              Quando la ricevi, accedi con quella password: ti chiederemo subito di sceglierne una nuova.
            </p>
            <Button asChild className="w-full rounded-none">
              <Link to="/login">Torna al login</Link>
            </Button>
          </div>
        ) : (
          <>
            <p className="text-sm text-center text-muted-foreground mt-1">
              Scrivi il tuo username: avviseremo gli amministratori del tuo istituto, che ti daranno una password temporanea.
            </p>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                form.handleSubmit();
              }}
              className="mt-6"
            >
              <FieldGroup>
                <form.Field name="username">
                  {(field) => (
                    <Field>
                      <FieldLabel htmlFor="username">Username</FieldLabel>
                      <FieldContent>
                        <Input
                          id="username"
                          autoComplete="username"
                          placeholder="Il tuo username"
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

                <form.Field name="nota">
                  {(field) => (
                    <Field>
                      <FieldLabel htmlFor="nota">Nota per gli amministratori (facoltativa)</FieldLabel>
                      <FieldContent>
                        <Textarea
                          id="nota"
                          placeholder="Per esempio: sono in sala stampa fino alle 13"
                          maxLength={NOTA_MAX}
                          rows={3}
                          value={field.state.value}
                          onChange={(e) => field.handleChange(e.target.value)}
                          onBlur={field.handleBlur}
                        />
                        <span className="text-xs text-muted-foreground text-right">
                          {field.state.value.length}/{NOTA_MAX}
                        </span>
                        {field.state.meta.errors.map((error, index) => (
                          <span key={index} className="text-red-500 text-xs">{error?.message}</span>
                        ))}
                      </FieldContent>
                    </Field>
                  )}
                </form.Field>

                {serverError && (
                  <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded">{serverError}</div>
                )}

                <div className="pt-2 flex flex-col items-center">
                  <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting]}>
                    {([canSubmit, isSubmitting]) => (
                      <Button type="submit" className="w-full rounded-none" disabled={!canSubmit || isSubmitting}>
                        {isSubmitting ? "Invio in corso..." : "Avvisa gli amministratori"}
                      </Button>
                    )}
                  </form.Subscribe>
                  <p className="text-xs text-muted-foreground text-center mt-3">
                    Sei un amministratore? Per ora chiedi a un altro amministratore dell'istituto di generarti una password temporanea.
                  </p>
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
