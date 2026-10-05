import { Card, CardHeader, CardTitle, CardDescription, CardContent } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Field, FieldContent, FieldLabel } from '@/components/ui/field';
import { Input } from '@/components/ui/input';
import { useForm } from '@tanstack/react-form';
import { changePasswordFormSchema } from '@shared/validation';
import { useChangePassword } from '@/hooks/use-change-password';
import { formatError } from '@/lib/utils';

const emptyValues = {
  currentPassword: '',
  newPassword: '',
  confirmPassword: '',
};

function fieldErrorMessage(error: unknown): string {
  if (typeof error === 'string') return error;
  if (error && typeof error === 'object' && 'message' in error) {
    return String((error as { message?: string }).message ?? '');
  }
  return '';
}

export default function CambiaPasswordCard() {
  const changePasswordMutation = useChangePassword();

  const form = useForm({
    defaultValues: emptyValues,
    validators: {
      onChange: changePasswordFormSchema,
    },
    onSubmit: async ({ value }) => {
      await changePasswordMutation.mutateAsync(
        {
          currentPassword: value.currentPassword,
          newPassword: value.newPassword,
        },
        {
          onSuccess: () => {
            form.reset(emptyValues);
            changePasswordMutation.reset();
          },
        }
      );
    },
  });

  return (
    <Card className="mt-6">
      <CardHeader>
        <CardTitle className="text-lg">Cambia password</CardTitle>
        <CardDescription>
          Inserisci la password attuale e scegline una nuova di almeno 8 caratteri.
        </CardDescription>
      </CardHeader>
      <CardContent>
        <form
          onSubmit={(e) => {
            e.preventDefault();
            e.stopPropagation();
            form.handleSubmit();
          }}
          className="space-y-4"
        >
          <form.Field name="currentPassword">
            {(field) => (
              <Field>
                <FieldLabel htmlFor="current-password">Password attuale</FieldLabel>
                <FieldContent>
                  <Input
                    id="current-password"
                    type="password"
                    autoComplete="current-password"
                    value={field.state.value}
                    onChange={(e) => field.handleChange(e.target.value)}
                    onBlur={field.handleBlur}
                  />
                  {field.state.meta.errors.length > 0 && field.state.meta.isTouched && (
                    field.state.meta.errors.map((error, index) => (
                      <span key={index} className="text-red-500 text-xs">
                        {fieldErrorMessage(error)}
                      </span>
                    ))
                  )}
                </FieldContent>
              </Field>
            )}
          </form.Field>

          <form.Field name="newPassword">
            {(field) => (
              <Field>
                <FieldLabel htmlFor="new-password">Nuova password</FieldLabel>
                <FieldContent>
                  <Input
                    id="new-password"
                    type="password"
                    autoComplete="new-password"
                    value={field.state.value}
                    onChange={(e) => field.handleChange(e.target.value)}
                    onBlur={field.handleBlur}
                  />
                  {field.state.meta.errors.length > 0 && field.state.meta.isTouched && (
                    field.state.meta.errors.map((error, index) => (
                      <span key={index} className="text-red-500 text-xs">
                        {fieldErrorMessage(error)}
                      </span>
                    ))
                  )}
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
                    value={field.state.value}
                    onChange={(e) => field.handleChange(e.target.value)}
                    onBlur={field.handleBlur}
                  />
                  {field.state.meta.errors.length > 0 && field.state.meta.isTouched && (
                    field.state.meta.errors.map((error, index) => (
                      <span key={index} className="text-red-500 text-xs">
                        {fieldErrorMessage(error)}
                      </span>
                    ))
                  )}
                </FieldContent>
              </Field>
            )}
          </form.Field>

          {changePasswordMutation.isError && changePasswordMutation.error && (
            <p className="text-sm text-destructive" role="alert">
              {formatError(changePasswordMutation.error, 'Errore durante il cambio password.')}
            </p>
          )}

          <form.Subscribe selector={(state) => [state.canSubmit, state.isSubmitting, state.isDirty]}>
            {([canSubmit, isSubmitting, isDirty]) => (
              <Button
                type="submit"
                className="w-full"
                disabled={!canSubmit || isSubmitting || !isDirty || changePasswordMutation.isPending}
              >
                {isSubmitting || changePasswordMutation.isPending
                  ? 'Salvataggio in corso...'
                  : 'Aggiorna password'}
              </Button>
            )}
          </form.Subscribe>
        </form>
      </CardContent>
    </Card>
  );
}
