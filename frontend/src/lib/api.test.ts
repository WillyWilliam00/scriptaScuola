import { http, HttpResponse } from "msw"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { CAMBIO_PASSWORD_OBBLIGATORIO } from "@shared/types"
import { API, server } from "@/test/server"
import { loginAs, makeJwt, resetAuthStore } from "@/test/utils"
import { useAuthStore } from "@/store/auth-store"
import { api, isCambioPasswordObbligatorio, isSessionExpiredError } from "./api"

const tokenVecchio = makeJwt({ userId: "vecchio" })
const tokenNuovo = makeJwt({ userId: "nuovo" })

/** Endpoint protetto finto: risponde 200 solo con il token nuovo, 401 con quello vecchio. */
function endpointProtetto() {
  const autorizzazioni: (string | null)[] = []
  server.use(
    http.get(`${API}/docenti`, ({ request }) => {
      const auth = request.headers.get("Authorization")
      autorizzazioni.push(auth)
      if (auth === `Bearer ${tokenNuovo}`) return HttpResponse.json({ ok: true })
      return HttpResponse.json({ error: "Token scaduto" }, { status: 401 })
    }),
  )
  return autorizzazioni
}

/** Endpoint di refresh finto che conta le chiamate. */
function refresh(risposta: () => Response | Promise<Response>) {
  const chiamate = { count: 0 }
  server.use(
    http.post(`${API}/auth/refresh`, async () => {
      chiamate.count++
      return risposta()
    }),
  )
  return chiamate
}

describe("api", () => {
  beforeEach(() => {
    resetAuthStore()
    vi.stubGlobal("location", { href: "http://localhost:3000/docenti" })
    vi.spyOn(console, "warn").mockImplementation(() => {})
    vi.spyOn(console, "error").mockImplementation(() => {})
  })
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it("aggiunge l'access token a ogni richiesta", async () => {
    loginAs("admin", tokenNuovo)
    const autorizzazioni = endpointProtetto()

    await api.get("/docenti")

    expect(autorizzazioni).toEqual([`Bearer ${tokenNuovo}`])
  })

  it("senza token non invia l'header Authorization", async () => {
    const autorizzazioni = endpointProtetto()
    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 401 } })
    expect(autorizzazioni).toEqual([null])
  })

  it("con un 401 rinnova la sessione e ripete la richiesta con il nuovo token", async () => {
    loginAs("admin", tokenVecchio)
    const autorizzazioni = endpointProtetto()
    const chiamate = refresh(() => HttpResponse.json({ token: tokenNuovo }))

    const response = await api.get("/docenti")

    expect(response.data).toEqual({ ok: true })
    expect(chiamate.count).toBe(1)
    expect(autorizzazioni).toEqual([`Bearer ${tokenVecchio}`, `Bearer ${tokenNuovo}`])
    expect(useAuthStore.getState().token).toBe(tokenNuovo)
    expect(useAuthStore.getState().isAuthenticated).toBe(true)
  })

  it("più richieste che falliscono insieme aspettano un solo rinnovo", async () => {
    loginAs("admin", tokenVecchio)
    endpointProtetto()
    const chiamate = refresh(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20))
      return HttpResponse.json({ token: tokenNuovo })
    })

    const risposte = await Promise.all([api.get("/docenti"), api.get("/docenti"), api.get("/docenti")])

    expect(risposte.map((r) => r.data)).toEqual([{ ok: true }, { ok: true }, { ok: true }])
    expect(chiamate.count).toBe(1)
  })

  it("se il rinnovo fallisce per un errore di rete, l'utente resta collegato", async () => {
    loginAs("admin", tokenVecchio)
    endpointProtetto()
    refresh(() => HttpResponse.error())

    await expect(api.get("/docenti")).rejects.toThrow()

    expect(useAuthStore.getState().isAuthenticated).toBe(true)
    expect(useAuthStore.getState().utente).not.toBeNull()
    expect(window.location.href).toBe("http://localhost:3000/docenti")
  })

  it("se il backend risponde 500 al rinnovo, l'utente resta collegato", async () => {
    loginAs("admin", tokenVecchio)
    endpointProtetto()
    refresh(() => HttpResponse.json({ error: "Errore interno" }, { status: 500 }))

    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 500 } })

    expect(useAuthStore.getState().isAuthenticated).toBe(true)
    expect(window.location.href).toBe("http://localhost:3000/docenti")
  })

  it("se la sessione è scaduta o revocata fa logout e porta a /login", async () => {
    loginAs("admin", tokenVecchio)
    endpointProtetto()
    refresh(() => HttpResponse.json({ error: "Sessione scaduta" }, { status: 401 }))

    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 401 } })

    expect(useAuthStore.getState().isAuthenticated).toBe(false)
    expect(useAuthStore.getState().utente).toBeNull()
    expect(window.location.href).toBe("/login")
  })

  it("«Utente non trovato» fa logout subito, senza tentare il rinnovo", async () => {
    loginAs("admin", tokenVecchio)
    server.use(http.get(`${API}/docenti`, () => HttpResponse.json({ error: "Utente non trovato" }, { status: 401 })))
    const chiamate = refresh(() => HttpResponse.json({ token: tokenNuovo }))

    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 401 } })

    expect(chiamate.count).toBe(0)
    expect(useAuthStore.getState().isAuthenticated).toBe(false)
    expect(window.location.href).toBe("/login")
  })

  it("un 401 sul login (credenziali sbagliate) non fa partire il rinnovo", async () => {
    loginAs("admin", tokenVecchio)
    server.use(http.post(`${API}/auth/login`, () => HttpResponse.json({ error: "Credenziali non valide" }, { status: 401 })))
    const chiamate = refresh(() => HttpResponse.json({ token: tokenNuovo }))

    await expect(api.post("/auth/login", {})).rejects.toMatchObject({ response: { status: 401 } })

    expect(chiamate.count).toBe(0)
    expect(useAuthStore.getState().isAuthenticated).toBe(true)
  })

  it("senza un utente collegato un 401 non fa partire il rinnovo", async () => {
    endpointProtetto()
    const chiamate = refresh(() => HttpResponse.json({ token: tokenNuovo }))

    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 401 } })

    expect(chiamate.count).toBe(0)
  })

  it("se anche la richiesta ripetuta riceve 401, si ferma (nessun ciclo)", async () => {
    loginAs("admin", tokenVecchio)
    let tentativi = 0
    server.use(
      http.get(`${API}/docenti`, () => {
        tentativi++
        return HttpResponse.json({ error: "Token scaduto" }, { status: 401 })
      }),
    )
    const chiamate = refresh(() => HttpResponse.json({ token: tokenNuovo }))

    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 401 } })

    expect(tentativi).toBe(2)
    expect(chiamate.count).toBe(1)
  })

  it("gli altri errori arrivano al chiamante senza rinnovo", async () => {
    loginAs("admin", tokenVecchio)
    server.use(http.get(`${API}/docenti`, () => HttpResponse.json({ error: "Non trovato" }, { status: 404 })))
    const chiamate = refresh(() => HttpResponse.json({ token: tokenNuovo }))

    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 404 } })
    expect(chiamate.count).toBe(0)
  })

  it("un 403 «cambio password obbligatorio» segna l'utente senza fare logout", async () => {
    loginAs("collaboratore", tokenVecchio)
    server.use(
      http.get(`${API}/docenti`, () =>
        HttpResponse.json({ error: "Cambia la password", code: CAMBIO_PASSWORD_OBBLIGATORIO }, { status: 403 }),
      ),
    )

    await expect(api.get("/docenti")).rejects.toMatchObject({ response: { status: 403 } })

    const state = useAuthStore.getState()
    expect(state.isAuthenticated).toBe(true)
    expect(state.utente?.deveCambiarePassword).toBe(true)
  })
})

describe("isSessionExpiredError e isCambioPasswordObbligatorio", () => {
  const conRisposta = (status: number, data?: unknown) => ({ response: { status, data } })

  it("401 e 403 indicano una sessione non più valida", () => {
    expect(isSessionExpiredError(conRisposta(401))).toBe(true)
    expect(isSessionExpiredError(conRisposta(403))).toBe(true)
  })

  it("rete assente e errori del server non contano come sessione scaduta", () => {
    expect(isSessionExpiredError(new Error("Network Error"))).toBe(false)
    expect(isSessionExpiredError(conRisposta(500))).toBe(false)
    expect(isSessionExpiredError(conRisposta(503))).toBe(false)
  })

  it("il 403 del cambio password non è una sessione scaduta", () => {
    const errore = conRisposta(403, { code: CAMBIO_PASSWORD_OBBLIGATORIO })
    expect(isCambioPasswordObbligatorio(errore)).toBe(true)
    expect(isSessionExpiredError(errore)).toBe(false)
    expect(isCambioPasswordObbligatorio(conRisposta(403))).toBe(false)
  })
})
