import { useForm } from "@tanstack/react-form";
import type { Utente } from "@shared/types";
import { createUtenteFormSchema, modifyUtenteFormSchema, type CreateUtente } from "@shared/validation";

export type UtenteFormValues = {
    ruolo: "admin" | "collaboratore";
    email: string;
    username: string;
    password: string;
};

export const emptyUtenteFormValues: UtenteFormValues = {
    ruolo: "collaboratore",
    email: "",
    username: "",
    password: "",
};

export function getUtenteFormValues(utente: Utente): UtenteFormValues {
    if (utente.ruolo === "admin") {
        return { ...emptyUtenteFormValues, ruolo: "admin", email: utente.email };
    }
    return { ...emptyUtenteFormValues, ruolo: "collaboratore", username: utente.username };
}

/** Invia solo i campi del ruolo attivo: email per admin, username per collaboratore. */
function toUtentePayload(value: UtenteFormValues): CreateUtente {
    if (value.ruolo === "admin") {
        return { ruolo: "admin", email: value.email, password: value.password };
    }
    return { ruolo: "collaboratore", username: value.username, password: value.password };
}

interface UseUtenteFormOptions {
    passwordRequired: boolean;
    onSubmit: (data: CreateUtente) => void;
}

/** Form condiviso tra creazione (password obbligatoria) e modifica (password opzionale). */
export function useUtenteForm({ passwordRequired, onSubmit }: UseUtenteFormOptions) {
    const schema = passwordRequired ? createUtenteFormSchema : modifyUtenteFormSchema;

    return useForm({
        defaultValues: emptyUtenteFormValues,
        validators: {
            onChange: schema,
            onMount: schema,
        },
        onSubmit: async ({ value }) => {
            onSubmit(toUtentePayload(value));
        },
    });
}

export type UtenteFormApi = ReturnType<typeof useUtenteForm>;
