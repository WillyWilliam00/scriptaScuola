import { AxiosError, AxiosHeaders } from "axios"
import { describe, expect, it } from "vitest"
import type { Utente } from "@shared/types"
import { cn, formatError, utenteDisplayName } from "./utils"

function axiosErrorWith(data: unknown, message = "Request failed with status code 400") {
  return new AxiosError(message, "ERR_BAD_REQUEST", undefined, undefined, {
    status: 400,
    statusText: "Bad Request",
    headers: {},
    config: { headers: new AxiosHeaders() },
    data,
  })
}

describe("formatError", () => {
  it("usa il messaggio d'errore restituito dal backend", () => {
    expect(formatError(axiosErrorWith({ error: "Limite copie superato" }))).toBe("Limite copie superato")
  })

  it("senza messaggio del backend usa quello di axios", () => {
    expect(formatError(axiosErrorWith(undefined, "Network Error"))).toBe("Network Error")
  })

  it("usa il messaggio di un Error qualsiasi", () => {
    expect(formatError(new Error("qualcosa è andato storto"))).toBe("qualcosa è andato storto")
  })

  it("per valori sconosciuti usa il messaggio predefinito o quello passato", () => {
    expect(formatError("boh")).toBe("Si è verificato un errore.")
    expect(formatError(null, "Errore personalizzato")).toBe("Errore personalizzato")
  })
})

describe("utenteDisplayName", () => {
  const base = { id: "1", istitutoId: 1, createdAt: new Date(), updatedAt: new Date() }

  it("per un admin mostra l'email", () => {
    const admin: Utente = { ...base, ruolo: "admin", email: "preside@scuola.it" }
    expect(utenteDisplayName(admin)).toBe("preside@scuola.it")
  })

  it("per un collaboratore mostra lo username", () => {
    const collaboratore: Utente = { ...base, ruolo: "collaboratore", username: "mario" }
    expect(utenteDisplayName(collaboratore)).toBe("mario")
  })
})

describe("cn", () => {
  it("unisce le classi e risolve i conflitti di Tailwind", () => {
    const nascosto = false as boolean
    expect(cn("px-2", nascosto && "hidden", "px-4")).toBe("px-4")
  })
})
