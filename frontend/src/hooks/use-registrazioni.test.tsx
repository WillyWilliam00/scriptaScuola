import { act, renderHook, waitFor } from "@testing-library/react"
import { http, HttpResponse } from "msw"
import { toast } from "sonner"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { API, server } from "@/test/server"
import { stubDownload } from "@/test/download"
import { createSuspenseWrapper } from "@/test/suspense"
import { createWrapper, loginAs, paginated, resetAuthStore } from "@/test/utils"
import {
  useCreateRegistrazione,
  useDeleteAllRegistrazioni,
  useDeleteRegistrazione,
  useExportRegistrazioni,
  useRegistrazioniSuspense,
  useUpdateRegistrazione,
} from "./use-registrazioni"

vi.mock("sonner", () => ({ toast: { success: vi.fn(), error: vi.fn() } }))

const registrazione = {
  id: 5,
  docenteId: 1,
  copieEffettuate: 30,
  istitutoId: 1,
  utenteId: "u1",
  note: null,
  createdAt: "2026-01-01T00:00:00.000Z",
  updatedAt: "2026-01-01T00:00:00.000Z",
  docenteNome: "Mario",
  docenteCognome: "Rossi",
  utenteUsername: "mario",
  utenteEmail: null,
}

describe("use-registrazioni", () => {
  beforeEach(() => {
    resetAuthStore()
    loginAs("admin")
  })

  describe("useRegistrazioniSuspense", () => {
    function lista() {
      const richieste: URLSearchParams[] = []
      server.use(
        http.get(`${API}/registrazioni-copie`, ({ request }) => {
          richieste.push(new URL(request.url).searchParams)
          return HttpResponse.json(paginated([registrazione]))
        }),
      )
      return richieste
    }

    it("senza parametri chiede la prima pagina ordinata per data", async () => {
      const richieste = lista()
      const { wrapper } = createSuspenseWrapper()

      const { result } = renderHook(() => useRegistrazioniSuspense(), { wrapper })

      await waitFor(() => expect(result.current?.data.data).toEqual([registrazione]))
      expect(Object.fromEntries(richieste[0])).toEqual({ page: "1", pageSize: "20", sortField: "createdAt", sortOrder: "asc" })
    })

    it("invia tutti i filtri, comprese 0 copie", async () => {
      const richieste = lista()
      const { wrapper } = createSuspenseWrapper()

      const { result } = renderHook(
        () =>
          useRegistrazioniSuspense({
            page: 2,
            pageSize: 50,
            docenteId: 1,
            utenteId: "u1",
            docenteNome: "Mar",
            docenteCognome: "Ros",
            copieEffettuate: 0,
            utenteIdentifier: "mario",
            sortField: "copieEffettuate",
            sortOrder: "desc",
          }),
        { wrapper },
      )

      await waitFor(() => expect(result.current?.data).toBeDefined())
      expect(Object.fromEntries(richieste[0])).toEqual({
        page: "2",
        pageSize: "50",
        docenteId: "1",
        utenteId: "u1",
        docenteNome: "Mar",
        docenteCognome: "Ros",
        copieEffettuate: "0",
        utenteIdentifier: "mario",
        sortField: "copieEffettuate",
        sortOrder: "desc",
      })
    })
  })

  it("useCreateRegistrazione registra le copie e aggiorna registrazioni e docenti", async () => {
    let corpo: unknown
    server.use(
      http.post(`${API}/registrazioni-copie/new-registrazione`, async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ docenteId: 1, copieEffettuate: 30 }, { status: 201 })
      }),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useCreateRegistrazione(), { wrapper })
    await act(() => result.current.mutateAsync({ docenteId: 1, copieEffettuate: 30, utenteId: "u1" }))

    expect(corpo).toEqual({ docenteId: 1, copieEffettuate: 30, utenteId: "u1" })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["registrazioni"] })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
    expect(toast.success).toHaveBeenCalledWith("30 copie registrate con successo")
  })

  it("useCreateRegistrazione oltre il limite mostra l'errore del backend", async () => {
    server.use(
      http.post(`${API}/registrazioni-copie/new-registrazione`, () =>
        HttpResponse.json({ error: "Limite di copie superato" }, { status: 400 }),
      ),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useCreateRegistrazione(), { wrapper })
    await act(async () => {
      await result.current.mutateAsync({ docenteId: 1, copieEffettuate: 999, utenteId: "u1" }).catch(() => {})
    })

    expect(toast.error).toHaveBeenCalledWith("Limite di copie superato")
    expect(invalidate).not.toHaveBeenCalled()
  })

  it("useUpdateRegistrazione usa l'id nell'URL e aggiorna anche i docenti", async () => {
    let corpo: unknown
    server.use(
      http.put(`${API}/registrazioni-copie/update-registrazione/5`, async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ message: "ok", data: registrazione })
      }),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useUpdateRegistrazione(), { wrapper })
    await act(() => result.current.mutateAsync({ id: 5, data: { copieEffettuate: 10 } }))

    expect(corpo).toEqual({ copieEffettuate: 10 })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["registrazioni"] })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
    expect(toast.success).toHaveBeenCalledWith("Registrazione aggiornata con successo")
  })

  it("useDeleteRegistrazione elimina la registrazione indicata", async () => {
    let chiamato = false
    server.use(
      http.delete(`${API}/registrazioni-copie/delete-registrazione/5`, () => {
        chiamato = true
        return HttpResponse.json({ message: "ok", id: 5 })
      }),
    )
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useDeleteRegistrazione(), { wrapper })
    await act(() => result.current.mutateAsync(5))

    expect(chiamato).toBe(true)
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
    expect(toast.success).toHaveBeenCalledWith("Registrazione eliminata con successo")
  })

  it("useDeleteAllRegistrazioni azzera il conteggio", async () => {
    server.use(http.delete(`${API}/registrazioni-copie/delete-all`, () => HttpResponse.json({ message: "ok" })))
    const { queryClient, wrapper } = createWrapper()
    const invalidate = vi.spyOn(queryClient, "invalidateQueries")

    const { result } = renderHook(() => useDeleteAllRegistrazioni(), { wrapper })
    await act(() => result.current.mutateAsync())

    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["registrazioni"] })
    expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
    expect(toast.success).toHaveBeenCalledWith("Tutte le registrazioni sono state eliminate con successo")
  })

  it("useDeleteAllRegistrazioni mostra l'errore", async () => {
    server.use(
      http.delete(`${API}/registrazioni-copie/delete-all`, () => HttpResponse.json({ error: "Solo gli admin" }, { status: 403 })),
    )
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useDeleteAllRegistrazioni(), { wrapper })
    await act(async () => {
      await result.current.mutateAsync().catch(() => {})
    })

    expect(toast.error).toHaveBeenCalledWith("Solo gli admin")
  })

  it("useExportRegistrazioni invia ordinamento e filtri (anche 0 copie), senza paginazione", async () => {
    const download = stubDownload()
    let parametri: URLSearchParams | undefined
    server.use(
      http.get(`${API}/registrazioni-copie/export`, ({ request }) => {
        parametri = new URL(request.url).searchParams
        return new HttpResponse("docente;copie\nRossi;30", { headers: { "Content-Type": "text/csv" } })
      }),
    )
    const { wrapper } = createWrapper()

    const { result } = renderHook(() => useExportRegistrazioni(), { wrapper })
    await act(() => result.current.mutateAsync({ page: 3, pageSize: 20, sortField: "createdAt", sortOrder: "desc", docenteCognome: "Ros", copieEffettuate: 0 }))

    expect(Object.fromEntries(parametri!)).toEqual({
      sortField: "createdAt",
      sortOrder: "desc",
      docenteCognome: "Ros",
      copieEffettuate: "0",
    })
    expect(download.nomiScaricati[0]).toMatch(/^registrazioni-\d{4}-\d{2}-\d{2}\.csv$/)
    expect(toast.success).toHaveBeenCalledWith("Esportazione registrazioni completata")
  })
})
