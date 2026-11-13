import { useForm } from '@tanstack/react-form';
import { loginSchema } from '@shared/validation';
import { Field, FieldContent, FieldGroup, FieldLabel } from "@/components/ui/field";
import { CameraIcon, EyeIcon, EyeOff } from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Link, useNavigate, useLocation } from "react-router-dom";
import Footer from "@/components/layout/Footer";
import { login } from "@/lib/auth-api.js";
import { formatError } from "@/lib/utils";
import { useState } from "react";

/**
 * Componente Login con validazione Zod e integrazione auth-store
 * 
 * Spiegazione:
 * - Usa TanStack Form per gestione form state
 * - Valida manualmente con Zod prima di inviare (evita chiamate API con dati invalidi)
 * - Mostra errori specifici per ogni campo (es. "Password deve essere almeno 8 caratteri")
 * - Chiama auth-api.login() che salva automaticamente nello store Zustand
       * - Dopo login riuscito, redirect a /registra-copie (route protetta)
 */
export default function Login() {
  const navigate = useNavigate();
  const location = useLocation();
  const prefillIdentifier = location.state?.identifier ?? '';
  const [serverError, setServerError] = useState<string | null>(null);
  const [showPassword, setShowPassword] = useState(false);

  // Messaggio di successo da Register (se presente)
  const successMessage = location.state?.message;

  const form = useForm({
    defaultValues: {
      identifier: prefillIdentifier,
      password: '',
    },
    validators: {
      onChange: loginSchema,
    },
    onSubmit: async ({ value }) => {
      setServerError(null); // Reset errore server prima di ogni submit

      try {
        // Chiama API login
        await login(value.identifier, value.password);

        // Login riuscito, redirect alla pagina di registrazione copie
        navigate('/registra-copie');
      } catch (err: unknown) {
        console.log('errore', err);
        setServerError(formatError(err, "Errore durante il login. Riprova più tardi."));
      }
    },
  });

  return (
    <div className="min-h-screen flex items-center justify-center bg-primary ">
      <div className="bg-secondary rounded-lg p-6 w-full max-w-xl">
        <div className="flex items-center justify-center w-20 h-20 rounded-full bg-primary p-2 mx-auto">
          <HugeiconsIcon icon={CameraIcon} strokeWidth={2} className="text-secondary w-12 h-12" />
        </div>
        <h1 className="text-2xl font-semibold text-primary text-center mt-4">Registro Copie Scolastico</h1>
        <h3 className="text-sm text-center text-muted-foreground">Accesso Riservato</h3>

        <form
          onSubmit={(e) => {
            e.preventDefault();
            form.handleSubmit();
          }}
          className="mt-6"
        >
          <FieldGroup>
            {/* Campo Identifier (email o username) */}
            <form.Field
              name="identifier"
            >
              {(field) => (
                <Field>
                  <FieldLabel htmlFor="identifier">Email o Username</FieldLabel>
                  <FieldContent>
                    <Input
                      id="identifier"
                      type="text"
                      placeholder="Inserisci email o username"
                      className="rounded-none h-12"
                      value={field.state.value}
                      onChange={(e) => field.handleChange(e.target.value)}
                      onBlur={field.handleBlur}
                    />
                    {
                      field.state.meta.errors.length > 0 && (
                        field.state.meta.errors.map((error, index) => (
                          <span key={index} className="text-red-500 text-xs">
                            {error?.message}
                          </span>
                        ))
                      )

                    }
                  </FieldContent>
                </Field>
              )}
            </form.Field>

            {/* Campo Password */}
            <form.Field
              name="password"
            >
              {(field) => (
                <Field>
                  <FieldLabel htmlFor="password">Password</FieldLabel>
                  <FieldContent className=''>
                    <div className='flex items-center'>
                      <Input
                        id="password"
                        type={showPassword ? "text" : "password"}
                        placeholder="Inserisci la password"
                        className="rounded-none h-12"
                        value={field.state.value}
                        onChange={(e) => field.handleChange(e.target.value)}
                        onBlur={field.handleBlur}
                      />
                      <Button type="button" variant="default" size="icon" className='rounded-none h-12 w-12' onClick={() => setShowPassword(!showPassword)}>
                        {showPassword ? <HugeiconsIcon icon={EyeIcon} strokeWidth={2} className="w-4 h-4" /> : <HugeiconsIcon icon={EyeOff} strokeWidth={2} className="w-4 h-4" />}
                      </Button>
                    </div>
                    {
                      field.state.meta.errors.length > 0 && field.state.value.length > 0 && (
                        field.state.meta.errors.map((error, index) => (
                          <span key={index} className="text-red-500 text-xs">
                            {error?.message}
                          </span>
                        ))
                      )

                    }
                  </FieldContent>
                </Field>
              )}
            </form.Field>

            {/* Messaggio di successo da Register */}
            {successMessage && (
              <div className="bg-green-50 border border-green-200 text-green-700 px-4 py-3 rounded">
                {successMessage}
              </div>
            )}

            {/* Messaggio di errore dal server */}
            {serverError && (
              <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded">
                {serverError}
              </div>
            )}



            <div className="pt-2 flex flex-col items-center justify-center">
              <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting]}>
                {([canSubmit, isSubmitting]) => (
                  <Button type="submit" className="w-full rounded-none" disabled={!canSubmit || isSubmitting}>
                    {isSubmitting ? 'Accesso in corso...' : 'Accedi'}
                  </Button>
                )}
              </form.Subscribe>
              <Link to="/password-dimenticata" className="text-sm text-muted-foreground block text-center mt-3">
                <span className="text-sm text-muted-foreground underline hover:text-primary">Password dimenticata?</span>
              </Link>
              <Link to="/register" className="text-sm text-muted-foreground block text-center mt-2">
                <span className="text-sm text-muted-foreground underline hover:text-primary">Registrati</span>
              </Link>
            </div>
          </FieldGroup>
        </form>
      </div>
      <Footer />
    </div>
  );
}
