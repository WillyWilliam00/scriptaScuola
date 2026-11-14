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

/** Controllo leggero lato client: la validazione vera la fa il server. */
function isEmail(valore: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(valore);
}

/**
 * Pagina "Password dimenticata?".
 * - Email (amministratori): arriva un link per scegliere una nuova password.
 * - Username (collaboratori): la richiesta arriva agli admin dell'istituto, che generano
 *   una password temporanea e la consegnano di persona.
 * Il server risponde sempre allo stesso modo, anche se l'utente non esiste.
 */
export default function PasswordDimenticata() {
  const [inviata, setInviata] = useState<{ messaggio: string; perEmail: boolean } | null>(null);
  const [serverError, setServerError] = useState<string | null>(null);

  const form = useForm({
    defaultValues: { username: "", nota: "" },
    validators: { onChange: richiestaResetFormSchema },
    onSubmit: async ({ value }) => {
      setServerError(null);
      const identificativo = value.username.trim();
      try {
        // Con una @ è l'email di un amministratore: riceve il link per email.
        // Senza @ è lo username di un collaboratore: la richiesta va agli admin dell'istituto.
        if (isEmail(identificativo)) {
          const res = await api.post<{ message: string }>("/auth/password-dimenticata", { email: identificativo });
          setInviata({ messaggio: res.data.message, perEmail: true });
          return;
        }
        if (identificativo.includes("@")) {
          setServerError("L'indirizzo email non sembra valido.");
          return;
        }
        const res = await api.post<{ message: string }>("/auth/richiesta-reset", {
          username: identificativo,
          nota: value.nota || undefined,
        });
        setInviata({ messaggio: res.data.message, perEmail: false });
      } catch (err) {
        setServerError(formatError(err, "Non è stato possibile inviare la richiesta. Riprova più tardi."));
      }
    },
  });

  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-primary px-4 pt-10 pb-16">
      <div className="bg-secondary rounded-lg p-6 w-full max-w-xl">
        <div className="flex items-center justify-center w-20 h-20 rounded-full bg-primary p-2 mx-auto">
          <HugeiconsIcon icon={KeyIcon} strokeWidth={2} className="text-secondary w-12 h-12" />
        </div>
        <h1 className="text-2xl font-semibold text-primary text-center mt-4">Password dimenticata</h1>

        {inviata ? (
          <div className="mt-6 space-y-4" data-testid="richiesta-inviata">
            <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-3 rounded">
              {inviata.messaggio}
            </div>
            <p className="text-sm text-muted-foreground text-center">
              {inviata.perEmail
                ? "Il link vale 60 minuti e si può usare una sola volta."
                : "Quando la ricevi, accedi con quella password: ti chiederemo subito di sceglierne una nuova."}
            </p>
            <Button asChild className="w-full rounded-none">
              <Link to="/login">Torna al login</Link>
            </Button>
          </div>
        ) : (
          <>
            <p className="text-sm text-center text-muted-foreground mt-1">
              Amministratori: scrivete la vostra email e riceverete un link per scegliere una nuova password.
              Collaboratori: scrivete il vostro username e avviseremo gli amministratori dell'istituto.
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
                      <FieldLabel htmlFor="username">Username o email</FieldLabel>
                      <FieldContent>
                        <Input
                          id="username"
                          autoComplete="username"
                          placeholder="Username oppure email da amministratore"
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

                <form.Subscribe selector={(state) => state.values.username.includes("@")}>
                  {(perEmail) => perEmail ? null : (
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
                  )}
                </form.Subscribe>

                {serverError && (
                  <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded">{serverError}</div>
                )}

                <div className="pt-2 flex flex-col items-center">
                  <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting, state.values.username.includes("@")] as const}>
                    {([canSubmit, isSubmitting, perEmail]) => (
                      <Button type="submit" className="w-full rounded-none" disabled={!canSubmit || isSubmitting}>
                        {isSubmitting ? "Invio in corso..." : perEmail ? "Inviami il link" : "Avvisa gli amministratori"}
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
