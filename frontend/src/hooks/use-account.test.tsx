import { act, renderHook, waitFor } from "@testing-library/react"
import { http, HttpResponse } from "msw"
import { toast } from "sonner"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { API, server } from "@/test/server"
import { createWrapper, loginAs, resetAuthStore } from "@/test/utils"
import { useAuthStore } from "@/store/auth-store"
import { useChangePassword } from "./use-change-password"
import { useDeleteIstituto } from "./use-istituto"
import { RICHIESTE_RESET_KEY, useGeneraPasswordTemporanea, useIgnoraRichiestaReset, useRichiesteReset } from "./use-richieste-reset"

vi.mock("sonner", () => ({ toast: { success: vi.fn(), error: vi.fn() } }))

beforeEach(() => {
  resetAuthStore()
})

describe("useChangePassword", () => {
  it("cambia la password e avvisa", async () => {
    loginAs("collaboratore")
    useAuthStore.getState().setDeveCambiarePassword(true)
    let corpo: unknown
    server.use(
      http.post(`${API}/auth/change-password`, async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ message: "ok" })
      }),
    )
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useChangePassword(), { wrapper })
    await act(() => result.current.mutateAsync({ currentPassword: "vecchia123", newPassword: "nuova12345" }))

    expect(corpo).toEqual({ currentPassword: "vecchia123", newPassword: "nuova12345" })
    expect(useAuthStore.getState().utente?.deveCambiarePassword).toBe(false)
    expect(toast.success).toHaveBeenCalledWith("Password aggiornata con successo")
  })

  it("con la password attuale sbagliata mostra l'errore del backend", async () => {
    loginAs("collaboratore")
    server.use(
      http.post(`${API}/auth/change-password`, () => HttpResponse.json({ error: "Password attuale errata" }, { status: 400 })),
    )
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useChangePassword(), { wrapper })
    await act(async () => {
      await result.current.mutateAsync({ currentPassword: "sbagliata", newPassword: "nuova12345" }).catch(() => {})
    })

    expect(toast.error).toHaveBeenCalledWith("Password attuale errata")
    expect(toast.success).not.toHaveBeenCalled()
  })
})

describe("useDeleteIstituto", () => {
  it("elimina l'istituto e chiama la callback", async () => {
    loginAs("admin")
    let chiamato = false
    server.use(
      http.delete(`${API}/istituti/delete-istituto`, () => {
        chiamato = true
        return HttpResponse.json({ message: "ok", id: "1" })
      }),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")
    const onShowDialog = vi.fn()

    const { result } = renderHook(() => useDeleteIstituto(onShowDialog), { wrapper })
    await act(() => result.current.mutateAsync())

    expect(chiamato).toBe(true)
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["istituti"] })
    expect(onShowDialog).toHaveBeenCalledTimes(1)
  })

  it("in caso di errore avvisa e non chiama la callback", async () => {
    loginAs("admin")
    server.use(
      http.delete(`${API}/istituti/delete-istituto`, () => HttpResponse.json({ error: "Password errata" }, { status: 400 })),
    )
    const { wrapper } = createWrapper()
    const onShowDialog = vi.fn()

    const { result } = renderHook(() => useDeleteIstituto(onShowDialog), { wrapper })
    await act(async () => {
      await result.current.mutateAsync().catch(() => {})
    })

    expect(onShowDialog).not.toHaveBeenCalled()
    expect(toast.error).toHaveBeenCalledWith("Password errata")
  })
})

describe("use-richieste-reset", () => {
  const richiesta = { id: "r1", utenteId: "u2", username: "mario", createdAt: "2026-01-01T00:00:00.000Z" }

  it("un admin vede le richieste aperte", async () => {
    loginAs("admin")
    server.use(http.get(`${API}/utenti/richieste-reset`, () => HttpResponse.json({ data: [richiesta] })))
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useRichiesteReset(), { wrapper })

    await waitFor(() => expect(result.current.data).toEqual([richiesta]))
  })

  it("un collaboratore non chiama il server", async () => {
    loginAs("collaboratore")
    let chiamato = false
    server.use(
      http.get(`${API}/utenti/richieste-reset`, () => {
        chiamato = true
        return HttpResponse.json({ data: [] })
      }),
    )
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useRichiesteReset(), { wrapper })

    await new Promise((resolve) => setTimeout(resolve, 50))
    expect(result.current.fetchStatus).toBe("idle")
    expect(chiamato).toBe(false)
  })

  it("genera una password temporanea per l'utenza e aggiorna le richieste", async () => {
    loginAs("admin")
    server.use(
      http.post(`${API}/utenti/u2/password-temporanea`, () =>
        HttpResponse.json({ passwordTemporanea: "Temp-5678", utente: { id: "u2" } }),
      ),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useGeneraPasswordTemporanea(), { wrapper })
    const risposta = await act(() => result.current.mutateAsync("u2"))

    expect(risposta).toMatchObject({ passwordTemporanea: "Temp-5678" })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: RICHIESTE_RESET_KEY })
  })

  it("se la generazione fallisce avvisa", async () => {
    loginAs("admin")
    server.use(
      http.post(`${API}/utenti/u2/password-temporanea`, () => HttpResponse.json({ error: "Utente non valido" }, { status: 400 })),
    )
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useGeneraPasswordTemporanea(), { wrapper })
    await act(async () => {
      await result.current.mutateAsync("u2").catch(() => {})
    })

    expect(toast.error).toHaveBeenCalledWith("Utente non valido")
  })

  it("ignora una richiesta e avvisa", async () => {
    loginAs("admin")
    let chiamato = false
    server.use(
      http.post(`${API}/utenti/richieste-reset/r1/ignora`, () => {
        chiamato = true
        return new HttpResponse(null, { status: 204 })
      }),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useIgnoraRichiestaReset(), { wrapper })
    await act(() => result.current.mutateAsync("r1"))

    expect(chiamato).toBe(true)
    expect(invalidate).toHaveBeenCalledWith({ queryKey: RICHIESTE_RESET_KEY })
    expect(toast.success).toHaveBeenCalledWith("Richiesta ignorata")
  })
})
