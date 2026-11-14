import { beforeEach, describe, expect, it } from "vitest"
import type { LoginResponse } from "@shared/types"
import { makeJwt, resetAuthStore } from "@/test/utils"
import { useAuthStore } from "./auth-store"

function loginResponse(token = makeJwt({ userId: "u1", ruolo: "admin", istitutoId: 3 })): LoginResponse {
  return {
    token,
    utente: { id: "u1", email: "preside@scuola.it", username: null, ruolo: "admin", deveCambiarePassword: false },
    istituto: { id: 3, nome: "Liceo Galilei", codiceIstituto: "LG01" },
  }
}

function salvato() {
  const raw = localStorage.getItem("auth-storage")
  return raw ? JSON.parse(raw) : null
}

describe("auth-store", () => {
  beforeEach(() => {
    resetAuthStore()
    localStorage.clear()
  })

  it("parte senza utente collegato", () => {
    const state = useAuthStore.getState()
    expect(state.isAuthenticated).toBe(false)
    expect(state.token).toBeNull()
    expect(state.utente).toBeNull()
  })

  it("login salva token, utente, istituto e payload del JWT", () => {
    const response = loginResponse()
    useAuthStore.getState().login(response)

    const state = useAuthStore.getState()
    expect(state.isAuthenticated).toBe(true)
    expect(state.token).toBe(response.token)
    expect(state.utente).toEqual(response.utente)
    expect(state.istituto).toEqual(response.istituto)
    expect(state.jwtPayload).toMatchObject({ userId: "u1", ruolo: "admin", istitutoId: 3 })
  })

  it("in localStorage finiscono utente e istituto, mai il token", () => {
    const response = loginResponse()
    useAuthStore.getState().login(response)

    const { state } = salvato()
    expect(state.utente).toEqual(response.utente)
    expect(state.istituto).toEqual(response.istituto)
    expect(JSON.stringify(salvato())).not.toContain(response.token)
    expect(state).not.toHaveProperty("token")
    expect(state).not.toHaveProperty("refreshToken")
  })

  it("logout pulisce lo stato e localStorage", () => {
    useAuthStore.getState().login(loginResponse())
    useAuthStore.getState().logout()

    const state = useAuthStore.getState()
    expect(state.isAuthenticated).toBe(false)
    expect(state.token).toBeNull()
    expect(state.utente).toBeNull()
    expect(state.jwtPayload).toBeNull()
    expect(salvato().state.utente).toBeNull()
  })

  it("setToken aggiorna token e payload e lascia collegato l'utente", () => {
    useAuthStore.getState().login(loginResponse())
    const nuovo = makeJwt({ userId: "u1", ruolo: "admin", istitutoId: 3, exp: 2000000000 })
    useAuthStore.getState().setToken(nuovo)

    const state = useAuthStore.getState()
    expect(state.token).toBe(nuovo)
    expect(state.jwtPayload?.exp).toBe(2000000000)
    expect(state.isAuthenticated).toBe(true)
  })

  it("setToken senza un utente non rende autenticati", () => {
    useAuthStore.getState().setToken(makeJwt())
    expect(useAuthStore.getState().isAuthenticated).toBe(false)
  })

  it("setDeveCambiarePassword aggiorna l'utente, e senza utente non fa nulla", () => {
    useAuthStore.getState().setDeveCambiarePassword(true)
    expect(useAuthStore.getState().utente).toBeNull()

    useAuthStore.getState().login(loginResponse())
    useAuthStore.getState().setDeveCambiarePassword(true)
    expect(useAuthStore.getState().utente?.deveCambiarePassword).toBe(true)
  })

  it("setInitialized chiude lo stato di caricamento iniziale", () => {
    expect(useAuthStore.getState().isInitializing).toBe(true)
    useAuthStore.getState().setInitialized()
    expect(useAuthStore.getState().isInitializing).toBe(false)
  })

  describe("ripristino da localStorage", () => {
    it("un utente salvato risulta collegato, ma senza token finché la sessione non è rinnovata", async () => {
      const { utente, istituto } = loginResponse()
      localStorage.setItem("auth-storage", JSON.stringify({ state: { utente, istituto, jwtPayload: null }, version: 1 }))

      await useAuthStore.persist.rehydrate()

      const state = useAuthStore.getState()
      expect(state.utente).toEqual(utente)
      expect(state.isAuthenticated).toBe(true)
      expect(state.token).toBeNull()
    })

    it("i dati delle versioni vecchie perdono token e refreshToken", async () => {
      const { utente, istituto } = loginResponse()
      localStorage.setItem(
        "auth-storage",
        JSON.stringify({
          state: { token: "vecchio-access", refreshToken: "vecchio-refresh", utente, istituto, jwtPayload: null },
          version: 0,
        }),
      )

      await useAuthStore.persist.rehydrate()

      const state = useAuthStore.getState()
      expect(state.token).toBeNull()
      expect(state).not.toHaveProperty("refreshToken")
      expect(state.utente).toEqual(utente)
      expect(JSON.stringify(salvato())).not.toContain("vecchio-")
    })
  })
})
