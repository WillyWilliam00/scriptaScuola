import { act, renderHook, waitFor } from "@testing-library/react"
import { http, HttpResponse } from "msw"
import { toast } from "sonner"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { API, server } from "@/test/server"
import { stubDownload } from "@/test/download"
import { createSuspenseWrapper } from "@/test/suspense"
import { createWrapper, loginAs, paginated, resetAuthStore } from "@/test/utils"
import { useCreateUtente, useDeleteUtente, useExportUtenti, useUpdateUtente, useUtenti } from "./use-utenti"

vi.mock("sonner", () => ({ toast: { success: vi.fn(), error: vi.fn() } }))

const utente = {
  id: "8f0c2c1e-0000-4000-8000-000000000001",
  istitutoId: 1,
  ruolo: "collaboratore",
  username: "mario",
  createdAt: "2026-01-01T00:00:00.000Z",
  updatedAt: "2026-01-01T00:00:00.000Z",
}

describe("use-utenti", () => {
  beforeEach(() => {
    resetAuthStore()
    loginAs("admin")
  })

  describe("useUtenti", () => {
    it("senza parametri chiede la prima pagina ordinata per ruolo", async () => {
      const richieste: URLSearchParams[] = []
      server.use(
        http.get(`${API}/utenti`, ({ request }) => {
          richieste.push(new URL(request.url).searchParams)
          return HttpResponse.json(paginated([utente]))
        }),
      )
      const { wrapper } = createSuspenseWrapper()

      const { result } = renderHook(() => useUtenti(), { wrapper })

      await waitFor(() => expect(result.current?.data.data).toEqual([utente]))
      expect(Object.fromEntries(richieste[0])).toEqual({ page: "1", pageSize: "20", sortField: "ruolo", sortOrder: "asc" })
    })

    it("invia ricerca, ruolo e ordinamento", async () => {
      const richieste: URLSearchParams[] = []
      server.use(
        http.get(`${API}/utenti`, ({ request }) => {
          richieste.push(new URL(request.url).searchParams)
          return HttpResponse.json(paginated([]))
        }),
      )
      const { wrapper } = createSuspenseWrapper()

      const { result } = renderHook(
        () => useUtenti({ page: 2, pageSize: 10, identifier: "mar", ruolo: "collaboratore", sortField: "identificativo", sortOrder: "desc" }),
        { wrapper },
      )

      await waitFor(() => expect(result.current?.data).toBeDefined())
      expect(Object.fromEntries(richieste[0])).toEqual({
        page: "2",
        pageSize: "10",
        identifier: "mar",
        ruolo: "collaboratore",
        sortField: "identificativo",
        sortOrder: "desc",
      })
    })
  })

  it("useCreateUtente invia il nuovo utente, aggiorna la lista e avvisa", async () => {
    let corpo: unknown
    server.use(
      http.post(`${API}/utenti/new-utente`, async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ message: "ok", data: utente, passwordTemporanea: "Temp-1234" }, { status: 201 })
      }),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useCreateUtente(), { wrapper })
    const risposta = await act(() => result.current.mutateAsync({ ruolo: "collaboratore", username: "mario" }))

    expect(corpo).toEqual({ ruolo: "collaboratore", username: "mario" })
    expect(risposta).toMatchObject({ passwordTemporanea: "Temp-1234" })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["utenti"] })
    expect(toast.success).toHaveBeenCalledWith("Utente creato con successo")
  })

  it("useCreateUtente mostra l'errore del backend (es. username già usato)", async () => {
    server.use(
      http.post(`${API}/utenti/new-utente`, () => HttpResponse.json({ error: "Username già in uso" }, { status: 409 })),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useCreateUtente(), { wrapper })
    await act(async () => {
      await result.current.mutateAsync({ ruolo: "collaboratore", username: "mario" }).catch(() => {})
    })

    expect(toast.error).toHaveBeenCalledWith("Username già in uso")
    expect(invalidate).not.toHaveBeenCalled()
  })

  it("useUpdateUtente usa l'id (UUID) nell'URL", async () => {
    let corpo: unknown
    server.use(
      http.put(`${API}/utenti/update-utente/${utente.id}`, async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ message: "ok", data: utente })
      }),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useUpdateUtente(), { wrapper })
    await act(() => result.current.mutateAsync({ id: utente.id, data: { ruolo: "collaboratore", username: "mario2" } }))

    expect(corpo).toEqual({ ruolo: "collaboratore", username: "mario2" })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["utenti"] })
    expect(toast.success).toHaveBeenCalledWith("Utente aggiornato con successo")
  })

  it("useDeleteUtente elimina l'utente e chiama la callback", async () => {
    server.use(
      http.delete(`${API}/utenti/delete-utente/${utente.id}`, () => HttpResponse.json({ message: "ok", id: utente.id })),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")
    const onShowDialog = vi.fn()

    const { result } = renderHook(() => useDeleteUtente(onShowDialog), { wrapper })
    await act(() => result.current.mutateAsync(utente.id))

    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["utenti"] })
    expect(onShowDialog).toHaveBeenCalledTimes(1)
    expect(toast.success).toHaveBeenCalledWith("Utente eliminato con successo")
  })

  it("useDeleteUtente in caso di errore non chiama la callback", async () => {
    server.use(
      http.delete(`${API}/utenti/delete-utente/${utente.id}`, () =>
        HttpResponse.json({ error: "Non puoi eliminare te stesso" }, { status: 400 }),
      ),
    )
    const { wrapper } = createWrapper()
    const onShowDialog = vi.fn()

    const { result } = renderHook(() => useDeleteUtente(onShowDialog), { wrapper })
    await act(async () => {
      await result.current.mutateAsync(utente.id).catch(() => {})
    })

    expect(onShowDialog).not.toHaveBeenCalled()
    expect(toast.error).toHaveBeenCalledWith("Non puoi eliminare te stesso")
  })

  it("useExportUtenti scarica il CSV con i filtri attivi", async () => {
    const download = stubDownload()
    let parametri: URLSearchParams | undefined
    server.use(
      http.get(`${API}/utenti/export`, ({ request }) => {
        parametri = new URL(request.url).searchParams
        return new HttpResponse("username\nmario", { headers: { "Content-Type": "text/csv" } })
      }),
    )
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useExportUtenti(), { wrapper })
    await act(() => result.current.mutateAsync({ page: 1, pageSize: 20, identifier: "mar", ruolo: "admin", sortField: "ruolo", sortOrder: "asc" }))

    expect(Object.fromEntries(parametri!)).toEqual({ identifier: "mar", ruolo: "admin", sortField: "ruolo", sortOrder: "asc" })
    expect(download.nomiScaricati[0]).toMatch(/^utenti-\d{4}-\d{2}-\d{2}\.csv$/)
    expect(toast.success).toHaveBeenCalledWith("Esportazione utenti completata")
  })
})
