import { describe, expect, it } from "vitest"
import type { Utente } from "@shared/types"
import { emptyUtenteFormValues, getUtenteFormValues, toUtentePayload } from "./utente-form"

const base = { id: "1", istitutoId: 1, createdAt: new Date(), updatedAt: new Date() }

describe("getUtenteFormValues", () => {
  it("per un admin precompila ruolo ed email", () => {
    const admin: Utente = { ...base, ruolo: "admin", email: "preside@scuola.it" }
    expect(getUtenteFormValues(admin)).toEqual({ ruolo: "admin", email: "preside@scuola.it", username: "" })
  })

  it("per un collaboratore precompila ruolo e username", () => {
    const collaboratore: Utente = { ...base, ruolo: "collaboratore", username: "mario" }
    expect(getUtenteFormValues(collaboratore)).toEqual({ ruolo: "collaboratore", email: "", username: "mario" })
  })

  it("non modifica i valori vuoti condivisi", () => {
    getUtenteFormValues({ ...base, ruolo: "admin", email: "x@y.it" })
    expect(emptyUtenteFormValues).toEqual({ ruolo: "collaboratore", email: "", username: "" })
  })
})

describe("toUtentePayload", () => {
  it("per un admin invia solo l'email", () => {
    expect(toUtentePayload({ ruolo: "admin", email: "a@b.it", username: "rimasto-dal-cambio-ruolo" })).toEqual({
      ruolo: "admin",
      email: "a@b.it",
    })
  })

  it("per un collaboratore invia solo lo username", () => {
    expect(toUtentePayload({ ruolo: "collaboratore", email: "rimasta@b.it", username: "mario" })).toEqual({
      ruolo: "collaboratore",
      username: "mario",
    })
  })
})
