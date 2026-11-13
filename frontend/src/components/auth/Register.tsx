import { useForm } from '@tanstack/react-form';
import { insertIstitutoSchema, registerSchema } from '@shared/validation';
import { Field, FieldContent, FieldGroup, FieldLabel } from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { CameraIcon, EyeIcon, EyeOff } from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";
import Footer from "@/components/layout/Footer";
import { Link, useNavigate } from "react-router-dom";
import { register } from "@/lib/auth-api.js";
import { formatError } from "@/lib/utils";
import { useState } from "react";
import { ZodError } from "zod";

/**
 * Componente Register con validazione Zod e gestione errori
 * 
 * Spiegazione:
 * - Usa TanStack Form per gestione form state
 * - Valida manualmente con Zod prima di inviare (coerenza con backend)
 * - Chiama auth-api.register() invece di fetch diretto
 * - Dopo registrazione riuscita: redirect a /login con messaggio di successo
 * - Gestisce errori specifici (es. istituto già esistente, email già in uso)
 */
export default function Register() {
  const navigate = useNavigate();
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const form = useForm({
    defaultValues: {
      istituto: {
        nome: '',
        codiceIstituto: '',
      },
      utente: {
        email: '',
        password: '',
      },
    },
    onSubmit: async ({ value }) => {
      setIsLoading(true);
      setError(null);

      try {
        // Validazione Zod manuale prima di inviare
        const validatedIstituto = insertIstitutoSchema.parse(value.istituto);
        const validatedUtente = registerSchema.parse(value.utente);

        // Chiama API register
        await register(validatedIstituto, validatedUtente);

        // Registrazione riuscita, redirect a login con:
        // - messaggio di successo
        // - email precompilata nel form di login
        navigate('/login', {
          state: {
            message: 'Configurazione completata con successo! Ora puoi accedere.',
            identifier: validatedUtente.email,
          },
        });
      } catch (err) {
        // Gestione errori Zod o API
        if (err instanceof ZodError) {
          // Errore di validazione Zod (in Zod 4 gli errori stanno in `issues`)
          const firstError = err.issues[0];
          const fieldPath = firstError?.path.join('.') ?? '';
          setError(`${fieldPath}: ${firstError?.message ?? 'Dati non validi'}`);
        } else {
          setError(formatError(err, "Errore durante la registrazione"));
        }
      } finally {
        setIsLoading(false);
      }
    },
  });

  return (
    <div className="min-h-screen flex items-center justify-center bg-primary ">
      <div className="bg-secondary rounded-lg p-6 w-full max-w-xl">
        <div className="flex items-center justify-center w-20 h-20 rounded-full bg-primary p-2 mx-auto">
          <HugeiconsIcon icon={CameraIcon} strokeWidth={2} className="text-secondary w-12 h-12" />
        </div>
        <h1 className="text-2xl font-semibold text-primary text-center mt-4">
          Configurazione Registro Copie
        </h1>
        <h3 className="text-sm text-center text-muted-foreground">
          Configura il tuo istituto e crea l&apos;account amministratore
        </h3>

        <form
          onSubmit={(e) => {
            e.preventDefault();
            form.handleSubmit();
          }}
          className="mt-6"
        >
          <FieldGroup>
            {/* Campo Nome Istituto */}
            <form.Field name="istituto.nome">
              {(field) => (
                <Field>
                  <FieldLabel htmlFor="istituto-nome">Nome Istituto</FieldLabel>
                  <FieldContent>
                    <Input
                      id="istituto-nome"
                      type="text"
                      placeholder="Inserisci il nome dell'istituto"
                      className="rounded-none h-12"
                      value={field.state.value}
                      onChange={(e) => field.handleChange(e.target.value)}
                      onBlur={field.handleBlur}
                    />
                  </FieldContent>
                </Field>
              )}
            </form.Field>

            {/* Campo Codice Istituto */}
            <form.Field name="istituto.codiceIstituto">
              {(field) => (
                <Field>
                  <FieldLabel htmlFor="istituto-codice">Codice Istituto</FieldLabel>
                  <FieldContent>
                    <Input
                      id="istituto-codice"
                      type="text"
                      placeholder="Inserisci il codice meccanografico (10 caratteri)"
                      className="rounded-none h-12"
                      value={field.state.value}
                      onChange={(e) => field.handleChange(e.target.value)}
                      onBlur={field.handleBlur}
                      maxLength={10}
                    />
                  </FieldContent>
                </Field>
              )}
            </form.Field>

            {/* Campo Email Amministratore */}
            <form.Field name="utente.email">
              {(field) => (
                <Field>
                  <FieldLabel htmlFor="utente-email">Email amministratore</FieldLabel>
                  <FieldContent>
                    <Input
                      id="utente-email"
                      type="email"
                      placeholder="Inserisci l'email dell'amministratore"
                      className="rounded-none h-12"
                      value={field.state.value}
                      onChange={(e) => field.handleChange(e.target.value)}
                      onBlur={field.handleBlur}
                    />
                  </FieldContent>
                </Field>
              )}
            </form.Field>

            {/* Campo Password Amministratore */}
            <form.Field name="utente.password">
              {(field) => (
                <Field>
                  <FieldLabel htmlFor="utente-password">Password amministratore</FieldLabel>
                  <FieldContent>
                    <div className='flex items-center'>
                      <Input
                        id="utente-password"
                        type={showPassword ? "text" : "password"}
                        placeholder="Inserisci la password (minimo 8 caratteri)"
                        className="rounded-none h-12"
                        value={field.state.value}
                        onChange={(e) => field.handleChange(e.target.value)}
                        onBlur={field.handleBlur}
                      />
                      <Button type="button" variant="default" size="icon" className='rounded-none h-12 w-12 ' onClick={() => setShowPassword(!showPassword)}>
                        {showPassword ? <HugeiconsIcon icon={EyeIcon} strokeWidth={2} className="w-4 h-4" /> : <HugeiconsIcon icon={EyeOff} strokeWidth={2} className="w-4 h-4" />}
                      </Button>
                    </div>

                  </FieldContent>
                </Field>
              )}
            </form.Field>

            {/* Messaggio di errore generale */}
            {error && (
              <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded">
                {error}
              </div>
            )}

            <div className="pt-2">
              <Button
                type="submit"
                className="w-full rounded-none"
                size="lg"
                variant="default"
                disabled={isLoading}
              >
                {isLoading ? 'Configurazione in corso...' : 'Configura istituto'}
              </Button>
              <Link to="/login" className="text-sm text-muted-foreground block text-center mt-2">
                <span className="text-sm text-muted-foreground underline hover:text-primary">Accedi</span>
              </Link>
            </div>
          </FieldGroup>
        </form>
      </div>
      <Footer />
    </div>
  );
}
