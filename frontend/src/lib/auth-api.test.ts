import { http, HttpResponse } from "msw"
import { beforeEach, describe, expect, it, vi } from "vitest"
import type { LoginResponse } from "@shared/types"
import { API, server } from "@/test/server"
import { createTestQueryClient, loginAs, makeJwt, resetAuthStore } from "@/test/utils"
import { useAuthStore } from "@/store/auth-store"
import { changePassword, login, logout, register, restoreSession } from "./auth-api"

const rispostaLogin: LoginResponse = {
  token: makeJwt({ userId: "u1", ruolo: "collaboratore" }),
  utente: { id: "u1", email: null, username: "mario", ruolo: "collaboratore", deveCambiarePassword: false },
  istituto: { id: 1, nome: "Liceo Galilei", codiceIstituto: "LG01" },
}

describe("auth-api", () => {
  beforeEach(() => {
    resetAuthStore()
    vi.spyOn(console, "warn").mockImplementation(() => {})
    vi.spyOn(console, "error").mockImplementation(() => {})
  })

  describe("login", () => {
    it("invia identifier e password e collega l'utente", async () => {
      let corpo: unknown
      server.use(
        http.post(`${API}/auth/login`, async ({ request }) => {
          corpo = await request.json()
          return HttpResponse.json(rispostaLogin)
        }),
      )

      const response = await login("mario", "password123")

      expect(corpo).toEqual({ identifier: "mario", password: "password123" })
      expect(response).toEqual(rispostaLogin)
      const state = useAuthStore.getState()
      expect(state.isAuthenticated).toBe(true)
      expect(state.token).toBe(rispostaLogin.token)
      expect(state.utente?.username).toBe("mario")
    })

    it("con credenziali sbagliate restituisce l'errore e non collega nessuno", async () => {
      server.use(http.post(`${API}/auth/login`, () => HttpResponse.json({ error: "Credenziali non valide" }, { status: 401 })))

      await expect(login("mario", "sbagliata")).rejects.toMatchObject({
        response: { status: 401, data: { error: "Credenziali non valide" } },
      })
      expect(useAuthStore.getState().isAuthenticated).toBe(false)
    })
  })

  describe("register", () => {
    it("invia istituto e admin a /auth/setup-scuola senza collegare l'utente", async () => {
      let corpo: unknown
      server.use(
        http.post(`${API}/auth/setup-scuola`, async ({ request }) => {
          corpo = await request.json()
          return HttpResponse.json({ message: "Istituto creato", istituto: {}, utente: {} }, { status: 201 })
        }),
      )

      const istituto = { nome: "Liceo Galilei", codiceIstituto: "ABCD123456" }
      const utente = { email: "preside@scuola.it", password: "password123" }
      const response = await register(istituto, utente)

      expect(corpo).toEqual({ istituto, utente })
      expect(response.message).toBe("Istituto creato")
      expect(useAuthStore.getState().isAuthenticated).toBe(false)
    })
  })

  describe("logout", () => {
    it("chiama il server, scollega l'utente e svuota la cache delle query", async () => {
      loginAs("admin")
      let chiamato = false
      server.use(
        http.post(`${API}/auth/logout`, () => {
          chiamato = true
          return HttpResponse.json({ message: "ok" })
        }),
      )
      const queryClient = createTestQueryClient()
      queryClient.setQueryData(["docenti"], [{ id: 1 }])

      await logout(queryClient)

      expect(chiamato).toBe(true)
      expect(useAuthStore.getState().isAuthenticated).toBe(false)
      expect(queryClient.getQueryData(["docenti"])).toBeUndefined()
    })

    it("scollega l'utente anche se il server non risponde", async () => {
      loginAs("admin")
      server.use(http.post(`${API}/auth/logout`, () => HttpResponse.error()))

      await expect(logout()).resolves.toEqual({ message: "Logout effettuato con successo" })
      expect(useAuthStore.getState().isAuthenticated).toBe(false)
    })
  })

  describe("changePassword", () => {
    it("invia password attuale e nuova e toglie l'obbligo di cambio", async () => {
      loginAs("collaboratore")
      useAuthStore.getState().setDeveCambiarePassword(true)
      let corpo: unknown
      server.use(
        http.post(`${API}/auth/change-password`, async ({ request }) => {
          corpo = await request.json()
          return HttpResponse.json({ message: "Password aggiornata" })
        }),
      )

      const response = await changePassword("vecchia123", "nuova12345")

      expect(corpo).toEqual({ currentPassword: "vecchia123", newPassword: "nuova12345" })
      expect(response.message).toBe("Password aggiornata")
      expect(useAuthStore.getState().utente?.deveCambiarePassword).toBe(false)
    })

    it("se il server rifiuta, l'obbligo di cambio resta", async () => {
      loginAs("collaboratore")
      useAuthStore.getState().setDeveCambiarePassword(true)
      server.use(
        http.post(`${API}/auth/change-password`, () =>
          HttpResponse.json({ error: "Password attuale errata" }, { status: 400 }),
        ),
      )

      await expect(changePassword("sbagliata", "nuova12345")).rejects.toMatchObject({ response: { status: 400 } })
      expect(useAuthStore.getState().utente?.deveCambiarePassword).toBe(true)
    })
  })

  describe("restoreSession", () => {
    /** Simula un ricaricamento della pagina: l'utente è in localStorage, il token in memoria no. */
    function dopoRicaricamento() {
      loginAs("admin")
      useAuthStore.setState({ token: null, isInitializing: true })
    }

    it("con un utente salvato chiede un nuovo access token", async () => {
      dopoRicaricamento()
      const nuovo = makeJwt({ userId: "u1" })
      server.use(http.post(`${API}/auth/refresh`, () => HttpResponse.json({ token: nuovo })))

      await restoreSession()

      const state = useAuthStore.getState()
      expect(state.token).toBe(nuovo)
      expect(state.isAuthenticated).toBe(true)
      expect(state.isInitializing).toBe(false)
    })

    it("se la sessione non è più valida fa logout", async () => {
      dopoRicaricamento()
      server.use(http.post(`${API}/auth/refresh`, () => HttpResponse.json({ error: "Sessione scaduta" }, { status: 401 })))

      await restoreSession()

      const state = useAuthStore.getState()
      expect(state.isAuthenticated).toBe(false)
      expect(state.utente).toBeNull()
      expect(state.isInitializing).toBe(false)
    })

    it("con la rete assente l'utente resta collegato", async () => {
      dopoRicaricamento()
      server.use(http.post(`${API}/auth/refresh`, () => HttpResponse.error()))

      await restoreSession()

      const state = useAuthStore.getState()
      expect(state.isAuthenticated).toBe(true)
      expect(state.utente).not.toBeNull()
      expect(state.isInitializing).toBe(false)
    })

    it("senza utente salvato non chiama il server", async () => {
      let chiamato = false
      server.use(
        http.post(`${API}/auth/refresh`, () => {
          chiamato = true
          return HttpResponse.json({ token: makeJwt() })
        }),
      )

      await restoreSession()

      expect(chiamato).toBe(false)
      expect(useAuthStore.getState().isInitializing).toBe(false)
    })
  })
})
