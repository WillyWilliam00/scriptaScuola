import { useForm } from "@tanstack/react-form";
import type { Utente } from "@shared/types";
import { utenteFormSchema, type CreateUtente } from "@shared/validation";

export type UtenteFormValues = {
    ruolo: "admin" | "collaboratore";
    email: string;
    username: string;
};

export const emptyUtenteFormValues: UtenteFormValues = {
    ruolo: "collaboratore",
    email: "",
    username: "",
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
        return { ruolo: "admin", email: value.email };
    }
    return { ruolo: "collaboratore", username: value.username };
}

interface UseUtenteFormOptions {
    onSubmit: (data: CreateUtente) => void;
}

/**
 * Form condiviso tra creazione e modifica. La password non si sceglie qui:
 * la genera la piattaforma (password temporanea da cambiare al primo accesso).
 */
export function useUtenteForm({ onSubmit }: UseUtenteFormOptions) {
    const schema = utenteFormSchema;

    return useForm({
        defaultValues: emptyUtenteFormValues,
        // Niente onMount: un errore calcolato al mount (ruolo collaboratore vuoto)
        // resterebbe in canSubmit anche dopo lo switch ad admin.
        validators: {
            onChange: schema,
        },
        onSubmit: async ({ value }) => {
            onSubmit(toUtentePayload(value));
        },
    });
}

export type UtenteFormApi = ReturnType<typeof useUtenteForm>;
