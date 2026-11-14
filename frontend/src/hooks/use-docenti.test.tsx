import { act, renderHook, waitFor } from "@testing-library/react"
import { http, HttpResponse, type JsonBodyType } from "msw"
import { toast } from "sonner"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { API, server } from "@/test/server"
import { stubDownload } from "@/test/download"
import { createSuspenseWrapper } from "@/test/suspense"
import { createWrapper, loginAs, paginated, resetAuthStore } from "@/test/utils"
import {
  useBulkImportDocenti,
  useCreateDocente,
  useDeleteAllDocenti,
  useDeleteDocente,
  useDocenti,
  useDocentiInfinite,
  useDocentiSuspense,
  useExportDocenti,
  useUpdateDocente,
} from "./use-docenti"

vi.mock("sonner", () => ({ toast: { success: vi.fn(), error: vi.fn() } }))

const docente = {
  id: 1,
  nome: "Mario",
  cognome: "Rossi",
  limiteCopie: 100,
  istitutoId: 1,
  createdAt: "2026-01-01T00:00:00.000Z",
  updatedAt: "2026-01-01T00:00:00.000Z",
  copieEffettuate: 10,
  copieRimanenti: 90,
}

/** Registra un handler GET /api/docenti e restituisce i parametri di ogni richiesta ricevuta. */
function listaDocenti(risposta: (url: URL) => JsonBodyType = () => paginated([docente])) {
  const richieste: URLSearchParams[] = []
  server.use(
    http.get(`${API}/docenti`, ({ request }) => {
      const url = new URL(request.url)
      richieste.push(url.searchParams)
      return HttpResponse.json(risposta(url))
    }),
  )
  return richieste
}

describe("use-docenti", () => {
  beforeEach(() => {
    resetAuthStore()
    loginAs("admin")
  })

  describe("useDocenti", () => {
    it("senza parametri chiede la prima pagina ordinata per data di creazione", async () => {
      const richieste = listaDocenti()
      const { wrapper } = createWrapper()

      const { result } = renderHook(() => useDocenti(), { wrapper })

      await waitFor(() => expect(result.current.isSuccess).toBe(true))
      expect(result.current.data?.data).toEqual([docente])
      expect(Object.fromEntries(richieste[0])).toEqual({
        page: "1",
        pageSize: "20",
        sortField: "createdAt",
        sortOrder: "asc",
      })
    })

    it("invia pagina, filtri e ordinamento scelti", async () => {
      const richieste = listaDocenti()
      const { wrapper } = createWrapper()

      const { result } = renderHook(
        () => useDocenti({ page: 3, pageSize: 50, nome: "Mar", cognome: "Ros", sortField: "cognome", sortOrder: "desc" }),
        { wrapper },
      )

      await waitFor(() => expect(result.current.isSuccess).toBe(true))
      expect(Object.fromEntries(richieste[0])).toEqual({
        page: "3",
        pageSize: "50",
        nome: "Mar",
        cognome: "Ros",
        sortField: "cognome",
        sortOrder: "desc",
      })
    })

    it("al cambio pagina mostra i dati precedenti finché arrivano i nuovi", async () => {
      listaDocenti((url) => paginated([{ ...docente, id: Number(url.searchParams.get("page")) }]))
      const { wrapper } = createWrapper()

      const { result, rerender } = renderHook(({ page }) => useDocenti({ page, pageSize: 20, sortField: "createdAt", sortOrder: "asc" }), {
        wrapper,
        initialProps: { page: 1 },
      })
      await waitFor(() => expect(result.current.data?.data[0].id).toBe(1))

      rerender({ page: 2 })
      expect(result.current.isPlaceholderData).toBe(true)
      expect(result.current.data?.data[0].id).toBe(1)

      await waitFor(() => expect(result.current.data?.data[0].id).toBe(2))
    })
  })

  describe("useDocentiSuspense", () => {
    it("carica la pagina con gli stessi parametri", async () => {
      const richieste = listaDocenti()
      const { wrapper } = createSuspenseWrapper()

      const { result } = renderHook(() => useDocentiSuspense({ page: 2, pageSize: 10, nome: "Anna", sortField: "nome", sortOrder: "asc" }), { wrapper })

      await waitFor(() => expect(result.current?.data.data).toEqual([docente]))
      expect(Object.fromEntries(richieste[0])).toEqual({ page: "2", pageSize: "10", nome: "Anna", sortField: "nome", sortOrder: "asc" })
    })
  })

  describe("useDocentiInfinite", () => {
    it("accoda le pagine successive finché il backend ne ha", async () => {
      const richieste = listaDocenti((url) => {
        const page = Number(url.searchParams.get("page"))
        return paginated([{ ...docente, id: page }], page, page < 2)
      })
      const { wrapper } = createSuspenseWrapper()

      const { result } = renderHook(() => useDocentiInfinite({ pageSize: 20, cognome: "Ros", sortField: "createdAt", sortOrder: "asc" }), { wrapper })
      await waitFor(() => expect(result.current?.data.pages).toHaveLength(1))
      expect(result.current.hasNextPage).toBe(true)

      await act(() => result.current.fetchNextPage())

      await waitFor(() => expect(result.current.data.pages).toHaveLength(2))
      expect(result.current.data.pages.map((p) => p.data[0].id)).toEqual([1, 2])
      expect(result.current.hasNextPage).toBe(false)
      expect(richieste.map((r) => r.get("page"))).toEqual(["1", "2"])
      expect(richieste[1].get("cognome")).toBe("Ros")
      expect(richieste[1].get("sortOrder")).toBe("asc")
    })
  })

  describe("mutazioni", () => {
    it("useCreateDocente invia il docente, aggiorna la lista e avvisa", async () => {
      let corpo: unknown
      server.use(
        http.post(`${API}/docenti/new-docente`, async ({ request }) => {
          corpo = await request.json()
          return HttpResponse.json({ message: "ok", data: docente }, { status: 201 })
        }),
      )
      const { queryClient, wrapper } = createWrapper()
      const invalidate = vi.spyOn(queryClient, "invalidateQueries")

      const { result } = renderHook(() => useCreateDocente(), { wrapper })
      await act(() => result.current.mutateAsync({ nome: "Mario", cognome: "Rossi", limiteCopie: 100 }))

      expect(corpo).toEqual({ nome: "Mario", cognome: "Rossi", limiteCopie: 100 })
      expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
      expect(toast.success).toHaveBeenCalledWith("Docente creato con successo")
    })

    it("useCreateDocente mostra l'errore del backend", async () => {
      server.use(
        http.post(`${API}/docenti/new-docente`, () => HttpResponse.json({ error: "Docente già presente" }, { status: 409 })),
      )
      const { queryClient, wrapper } = createWrapper()
      const invalidate = vi.spyOn(queryClient, "invalidateQueries")

      const { result } = renderHook(() => useCreateDocente(), { wrapper })
      await act(async () => {
        await result.current.mutateAsync({ nome: "Mario", cognome: "Rossi", limiteCopie: 100 }).catch(() => {})
      })

      expect(toast.error).toHaveBeenCalledWith("Docente già presente")
      expect(toast.success).not.toHaveBeenCalled()
      expect(invalidate).not.toHaveBeenCalled()
    })

    it("useUpdateDocente usa l'id nell'URL", async () => {
      let corpo: unknown
      server.use(
        http.put(`${API}/docenti/update-docente/7`, async ({ request }) => {
          corpo = await request.json()
          return HttpResponse.json({ message: "ok", data: docente })
        }),
      )
      const { queryClient, wrapper } = createWrapper()
      const invalidate = vi.spyOn(queryClient, "invalidateQueries")

      const { result } = renderHook(() => useUpdateDocente(), { wrapper })
      await act(() => result.current.mutateAsync({ id: 7, data: { limiteCopie: 200 } }))

      expect(corpo).toEqual({ limiteCopie: 200 })
      expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
      expect(toast.success).toHaveBeenCalledWith("Docente aggiornato con successo")
    })

    it("useUpdateDocente con un errore senza messaggio usa quello predefinito", async () => {
      server.use(http.put(`${API}/docenti/update-docente/7`, () => HttpResponse.error()))
      const { wrapper } = createWrapper()

      const { result } = renderHook(() => useUpdateDocente(), { wrapper })
      await act(async () => {
        await result.current.mutateAsync({ id: 7, data: { limiteCopie: 200 } }).catch(() => {})
      })

      expect(toast.error).toHaveBeenCalledTimes(1)
      expect(toast.error).toHaveBeenCalledWith("Network Error")
    })

    it("useDeleteDocente elimina il docente indicato", async () => {
      let chiamato = false
      server.use(
        http.delete(`${API}/docenti/delete-docente/7`, () => {
          chiamato = true
          return HttpResponse.json({ message: "ok", id: 7 })
        }),
      )
      const { queryClient, wrapper } = createWrapper()
      const invalidate = vi.spyOn(queryClient, "invalidateQueries")

      const { result } = renderHook(() => useDeleteDocente(), { wrapper })
      await act(() => result.current.mutateAsync(7))

      expect(chiamato).toBe(true)
      expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
      expect(toast.success).toHaveBeenCalledWith("Docente eliminato con successo")
    })

    it("useDeleteAllDocenti aggiorna anche le registrazioni", async () => {
      server.use(http.delete(`${API}/docenti/delete-all`, () => HttpResponse.json({ message: "ok", deletedCount: 3 })))
      const { queryClient, wrapper } = createWrapper()
      const invalidate = vi.spyOn(queryClient, "invalidateQueries")

      const { result } = renderHook(() => useDeleteAllDocenti(), { wrapper })
      await act(() => result.current.mutateAsync())

      expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
      expect(invalidate).toHaveBeenCalledWith({ queryKey: ["registrazioni"] })
      expect(toast.success).toHaveBeenCalledWith("Tutti i docenti sono stati eliminati con successo")
    })

    it("useBulkImportDocenti invia la lista e aggiorna docenti e registrazioni", async () => {
      let corpo: unknown
      server.use(
        http.post(`${API}/docenti/bulk-import`, async ({ request }) => {
          corpo = await request.json()
          return HttpResponse.json({ message: "ok", data: { docenti: [], totaleCreati: 1, totaleConCopie: 1 } })
        }),
      )
      const { queryClient, wrapper } = createWrapper()
      const invalidate = vi.spyOn(queryClient, "invalidateQueries")
      const docenti = [{ nome: "Mario", cognome: "Rossi", limiteCopie: 100, copieEffettuate: 5 }]

      const { result } = renderHook(() => useBulkImportDocenti(), { wrapper })
      await act(() => result.current.mutateAsync({ docenti }))

      expect(corpo).toEqual({ docenti })
      expect(invalidate).toHaveBeenCalledWith({ queryKey: ["docenti"] })
      expect(invalidate).toHaveBeenCalledWith({ queryKey: ["registrazioni"] })
      expect(toast.success).toHaveBeenCalledWith("Docenti importati con successo")
    })

    it("useBulkImportDocenti mostra l'errore del backend", async () => {
      server.use(
        http.post(`${API}/docenti/bulk-import`, () => HttpResponse.json({ error: "Riga 2 non valida" }, { status: 400 })),
      )
      const { wrapper } = createWrapper()

      const { result } = renderHook(() => useBulkImportDocenti(), { wrapper })
      await act(async () => {
        await result.current.mutateAsync({ docenti: [] }).catch(() => {})
      })

      expect(toast.error).toHaveBeenCalledWith("Riga 2 non valida")
    })
  })

  describe("useExportDocenti", () => {
    it("scarica il CSV con i filtri attivi, senza paginazione", async () => {
      const download = stubDownload()
      let parametri: URLSearchParams | undefined
      server.use(
        http.get(`${API}/docenti/export`, ({ request }) => {
          parametri = new URL(request.url).searchParams
          return new HttpResponse("nome;cognome\nMario;Rossi", { headers: { "Content-Type": "text/csv" } })
        }),
      )
      const { wrapper } = createWrapper()

      const { result } = renderHook(() => useExportDocenti(), { wrapper })
      await act(() => result.current.mutateAsync({ page: 4, pageSize: 20, nome: "Mar", sortField: "cognome", sortOrder: "desc" }))

      expect(Object.fromEntries(parametri!)).toEqual({ nome: "Mar", sortField: "cognome", sortOrder: "desc" })
      expect(download.nomiScaricati).toHaveLength(1)
      expect(download.nomiScaricati[0]).toMatch(/^docenti-\d{4}-\d{2}-\d{2}\.csv$/)
      expect(download.revokeObjectURL).toHaveBeenCalledWith("blob:finto")
      expect(toast.success).toHaveBeenCalledWith("Esportazione docenti completata")
    })

    it("se l'export fallisce avvisa senza scaricare nulla", async () => {
      const download = stubDownload()
      server.use(http.get(`${API}/docenti/export`, () => new HttpResponse(null, { status: 500 })))
      const { wrapper } = createWrapper()

      const { result } = renderHook(() => useExportDocenti(), { wrapper })
      await act(async () => {
        await result.current.mutateAsync({ page: 1, pageSize: 20, sortField: "createdAt", sortOrder: "asc" }).catch(() => {})
      })

      expect(download.nomiScaricati).toHaveLength(0)
      expect(toast.error).toHaveBeenCalledTimes(1)
    })
  })
})
