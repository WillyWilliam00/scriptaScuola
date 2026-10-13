import { Field, FieldLabel, FieldContent } from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { Select, SelectTrigger, SelectValue, SelectContent, SelectGroup, SelectItem } from "@/components/ui/select";
import type { UtenteFormApi, UtenteFormValues } from "./utente-form";

interface UtenteFormFieldsProps {
    form: UtenteFormApi;
    ruoloDisabled: boolean;
    passwordLabel: string;
    passwordRequired: boolean;
}

function FieldErrors({ errors, isTouched }: { errors: unknown[]; isTouched: boolean }) {
    if (!isTouched || errors.length === 0) return null;
    return errors.map((error, index) => (
        <span key={index} className="text-red-500 text-xs">
            {(error as { message?: string } | undefined)?.message}
        </span>
    ));
}

/** Campi del form utente (ruolo, email/username, password), usati da creazione e modifica. */
export function UtenteFormFields({ form, ruoloDisabled, passwordLabel, passwordRequired }: UtenteFormFieldsProps) {
    return (
        <form onSubmit={(e) => {
            e.preventDefault();
            e.stopPropagation();
            form.handleSubmit();
        }}>
            <form.Field name="ruolo" children={(field) => (
                <Field>
                    <FieldLabel>Ruolo</FieldLabel>
                    <Select
                        disabled={ruoloDisabled || form.state.isSubmitting}
                        value={field.state.value}
                        required
                        onValueChange={(value) => field.handleChange(value as UtenteFormValues["ruolo"])}
                    >
                        <SelectTrigger>
                            <SelectValue placeholder="Ruolo" />
                        </SelectTrigger>
                        <SelectContent>
                            <SelectGroup>
                                <SelectItem value="admin">Admin</SelectItem>
                                <SelectItem value="collaboratore">Collaboratore</SelectItem>
                            </SelectGroup>
                        </SelectContent>
                    </Select>
                </Field>
            )} />

            <form.Subscribe selector={(state) => state.values.ruolo}>
                {(ruolo) => (
                    <>
                        {ruolo === "admin" && <form.Field name="email" children={(field) => (
                            <>
                                <Field className="mt-4">
                                    <FieldLabel htmlFor="email-utente">Email</FieldLabel>
                                    <FieldContent>
                                        <Input
                                            id="email-utente"
                                            type="email"
                                            placeholder="email@esempio.it"
                                            required
                                            value={field.state.value}
                                            onChange={(e) => field.handleChange(e.target.value)}
                                        />
                                    </FieldContent>
                                </Field>
                                <FieldErrors errors={field.state.meta.errors} isTouched={field.state.meta.isTouched} />
                            </>
                        )} />}
                        {ruolo === "collaboratore" && <form.Field name="username" children={(field) => (
                            <>
                                <Field className="mt-4">
                                    <FieldLabel htmlFor="username-utente">Username</FieldLabel>
                                    <FieldContent>
                                        <Input
                                            id="username-utente"
                                            placeholder="username (senza @)"
                                            required
                                            value={field.state.value}
                                            onChange={(e) => field.handleChange(e.target.value)}
                                        />
                                    </FieldContent>
                                </Field>
                                <FieldErrors errors={field.state.meta.errors} isTouched={field.state.meta.isTouched} />
                            </>
                        )} />}
                    </>
                )}
            </form.Subscribe>

            <form.Field name="password" children={(field) => (
                <>
                    <Field className="mt-4">
                        <FieldLabel htmlFor="password-utente">{passwordLabel}</FieldLabel>
                        <FieldContent>
                            <Input
                                id="password-utente"
                                type="password"
                                required={passwordRequired}
                                placeholder="Minimo 8 caratteri"
                                value={field.state.value}
                                onChange={(e) => field.handleChange(e.target.value)}
                            />
                        </FieldContent>
                    </Field>
                    <FieldErrors errors={field.state.meta.errors} isTouched={field.state.meta.isTouched} />
                </>
            )} />
        </form>
    );
}
