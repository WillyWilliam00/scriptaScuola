import { useQuery, useSuspenseQuery, useSuspenseInfiniteQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query';
import { api } from '../lib/api.js';
import type { DocentiPaginatedResponse, DocenteConRegistrazioni } from '../../../shared/types.js';
import type { BulkImportDocenti, DocentiQuery, InsertDocente, ModifyDocente } from '../../../shared/validation.js';
import { toast } from 'sonner';
import { formatError } from '@/lib/utils.js';

/**
 * Hook per gestire i docenti con TanStack Query
 *
 * Due pattern di paginazione:
 *
 * 1) useQuery (useDocenti)
 *    - Paginazione "a pagine": pulsanti "Pagina 1, 2, 3..." o "Precedente / Successivo"
 *    - Ogni cambio pagina sostituisce i dati (una pagina alla volta in cache)
 *    - Usa quando: tabella con selettore di pagina
 *
 * 2) useSuspenseQuery (useDocentiSuspense)
 *    - Come useDocenti ma sospende fino al primo caricamento: usare il componente che lo chiama dentro <Suspense>.
 *    - Restituisce data e isFetching (per indicatore di refetch in background). Con placeholderData: keepPreviousData
 *      al cambio pagina/filtri non sospende di nuovo e mostra i dati precedenti mentre arriva la nuova risposta.
 *
 * 3) useSuspenseInfiniteQuery (useDocentiInfinite)
 *    - Scroll infinito / "Carica altri": le pagine si accumulano; sospende fino alla prima pagina (usare dentro Suspense)
 *    - fetchNextPage() aggiunge la pagina successiva; hasNextPage, isFetchingNextPage
 *    - Usa quando: lista che cresce scrollando o bottone "Carica altri"
 *
 * Endpoint backend:
 * - GET /api/docenti?page=1&pageSize=20&...
 * - POST /api/docenti/new-docente
 * - PUT /api/docenti/update-docente/:id
 * - DELETE /api/docenti/delete-docente/:id
 * - DELETE /api/docenti/delete-all
 */

/**
 * Hook per ottenere la lista paginata dei docenti (paginazione a pagine).
 * Usa quando mostri "Pagina 1, 2, 3..." o Precedente/Successivo.
 *
 * Stati di loading (da usare nell'UI):
 * - isLoading: true = primo caricamento, nessun dato in cache → mostra skeleton/placeholder pieno
 * - isFetching: true = una richiesta è in corso (anche refetch in background) → se !isLoading puoi mostrare un piccolo indicatore
 * - isPlaceholderData: true = stai vedendo dati in cache mentre carichi un'altra pagina (es. keepPreviousData)
 *
 * Filtri/paginazione in URL: questo hook NON legge né scrive l'URL. Il componente che lo
 * usa può passare query da useState (solo in memoria) o da useSearchParams (React Router)
 * per avere page/filtri nell'URL (es. /docenti?page=2&nome=Mario) e permettere condivisione/refresh.
 *
 * @param query - Parametri di query per paginazione, filtri e ordinamento
 * @returns Query result con data, pagination, isLoading, isFetching, error, refetch, ecc.
 */
export function useDocenti(query: DocentiQuery = { page: 1, pageSize: 20, sortField: 'createdAt', sortOrder: 'asc' }) {
  return useQuery<DocentiPaginatedResponse>({
    // --- queryKey: identifica questa query in modo univoco.
    // La cache e il refetch si basano su questa chiave. Se query cambia (es. page 1 → 2),
    // la chiave cambia → TanStack Query tratta come richiesta diversa e può tenere in cache
    // sia la pagina 1 che la 2. Deve essere serializzabile (niente funzioni).
    queryKey: ['docenti', query],

    // --- queryFn: la funzione che fa la richiesta HTTP e restituisce i dati.
    // Viene chiamata quando: non c'è dato in cache, i dati sono "stale", o refetch manuale.
    // Il valore di ritorno viene messo in cache e restituito come data.
    queryFn: async () => {
      const params = new URLSearchParams();
      params.append('page', String(query.page ?? 1));
      params.append('pageSize', String(query.pageSize ?? 20));
      if (query.nome) params.append('nome', query.nome);
      if (query.cognome) params.append('cognome', query.cognome);
      if (query.sortField) params.append('sortField', query.sortField);
      if (query.sortOrder) params.append('sortOrder', query.sortOrder ?? 'asc');
      const response = await api.get<DocentiPaginatedResponse>(`/docenti?${params.toString()}`);
      return response.data;
    },

    // --- staleTime (ms): per quanto tempo i dati sono considerati "freschi".
    // Durante questo periodo NON parte un refetch automatico (es. al focus della finestra).
    // Dopo 2 minuti i dati diventano "stale" e la prossima volta che il componente monta
    // o la finestra riprende focus può scattare un refetch in background.
    staleTime: 2 * 60 * 1000,

    // --- placeholderData: cosa mostrare mentre si carica una nuova pagina.
    // keepPreviousData = mantieni i dati della richiesta precedente visibili finché
    // non arrivano i nuovi (evita flash bianco al cambio pagina).
    placeholderData: keepPreviousData,
  });
}

/**
 * Hook per ottenere la lista paginata dei docenti sospendendo fino al primo caricamento.
 * Usa useSuspenseQuery: il componente che lo chiama deve essere wrappato in <Suspense fallback={...}>.
 *
 * Stati dopo il primo load:
 * - isFetching: true quando è in corso un refetch (cambio pagina/filtri/sort o invalidazione).
 *   Puoi mostrare un indicatore discreto (es. barra o icona) senza nascondere la tabella.
 * - Con placeholderData: keepPreviousData, al cambio query (es. pagina 2) non sospende di nuovo:
 *   restano visibili i dati precedenti finché non arrivano i nuovi.
 *
 * @param query - Parametri di query per paginazione, filtri e ordinamento (stesso di useDocenti)
 * @returns Query result con data, pagination, isFetching, refetch, ecc. (niente isLoading: il loading è gestito da Suspense)
 */
export function useDocentiSuspense(query: DocentiQuery = { page: 1, pageSize: 20, sortField: 'createdAt', sortOrder: 'asc' }) {
  return useSuspenseQuery<DocentiPaginatedResponse>({
    queryKey: ['docenti', query],
    queryFn: async () => {
      const params = new URLSearchParams();
      params.append('page', String(query.page ?? 1));
      params.append('pageSize', String(query.pageSize ?? 20));
      if (query.nome) params.append('nome', query.nome);
      if (query.cognome) params.append('cognome', query.cognome);
      if (query.sortField) params.append('sortField', query.sortField);
      if (query.sortOrder) params.append('sortOrder', query.sortOrder ?? 'asc');
      const response = await api.get<DocentiPaginatedResponse>(`/docenti?${params.toString()}`);
      return response.data;
    },
    staleTime: 2 * 60 * 1000,
    // Nota: useSuspenseQuery in v5 non supporta placeholderData; al cambio query (es. pagina) può sospendere di nuovo.
    // Per evitarlo si può usare useQuery (useDocenti) con keepPreviousData; qui privilegiamo il pattern Suspense.
  });
}

/**
 * Hook per lista docenti con scroll infinito / "Carica altri".
 * Usa useSuspenseInfiniteQuery: sospende fino al caricamento della prima pagina (il componente che lo usa deve stare dentro <Suspense>).
 * Le pagine successive si caricano con fetchNextPage(); restano disponibili hasNextPage, isFetchingNextPage.
 *
 * @param baseQuery - Filtri e ordinamento (page non usato; la pagina è gestita internamente)
 * @returns Infinite query con data.pages, fetchNextPage, hasNextPage, isFetchingNextPage, ecc.
 */
export function useDocentiInfinite(
  baseQuery: Omit<DocentiQuery, 'page'> = { pageSize: 20, sortField: 'createdAt', sortOrder: 'asc' }
) {
  const pageSize = baseQuery.pageSize ?? 20;

  return useSuspenseInfiniteQuery({
    queryKey: ['docenti', 'infinite', { ...baseQuery, pageSize }],
    queryFn: async ({ pageParam = 1 }) => {
      const params = new URLSearchParams();
      params.append('page', String(pageParam));
      params.append('pageSize', String(pageSize));
      if (baseQuery.nome) params.append('nome', baseQuery.nome);
      if (baseQuery.cognome) params.append('cognome', baseQuery.cognome);
      if (baseQuery.sortField) params.append('sortField', baseQuery.sortField);
      params.append('sortOrder', baseQuery.sortOrder ?? 'asc');
      const response = await api.get<DocentiPaginatedResponse>(`/docenti?${params.toString()}`);
      return response.data;
    },
    getNextPageParam: (lastPage) => {
      const { page, hasNextPage } = lastPage.pagination;
      return hasNextPage ? page + 1 : undefined;
    },
    initialPageParam: 1,
    staleTime: 2 * 60 * 1000,
  });
}

/**
 * Hook per creare un nuovo docente
 *
 * Stati di loading: isPending (v5) / isLoading (v4) = true mentre la mutation è in corso → disabilita pulsante / mostra spinner
 *
 * @returns Mutation con mutate, mutateAsync, isPending (o isLoading), error, isError, ecc.
 */
export function useCreateDocente() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: Omit<InsertDocente, 'istitutoId'>) => {
      const response = await api.post<{ message: string; data: DocenteConRegistrazioni }>(
        '/docenti/new-docente',
        data
      );
      return response.data;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['docenti'] });
      toast.success("Docente creato con successo");
    },
    onError: (err) => {
      toast.error(formatError(err, "Errore durante la creazione del docente."));
    },
  });
}

/**
 * Hook per aggiornare un docente esistente
 *
 * Stati: isPending (o isLoading) = true durante l'update → disabilita form / mostra saving
 *
 * @returns Mutation con mutate, mutateAsync, isPending, error, ecc.
 */
export function useUpdateDocente() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, data }: { id: number; data: Omit<ModifyDocente, 'istitutoId'> }) => {
      const response = await api.put<{ message: string; data: DocenteConRegistrazioni }>(
        `/docenti/update-docente/${id}`,
        data
      );
      return response.data;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['docenti'] });
      toast.success("Docente aggiornato con successo");
    },
    onError: (err) => {
      toast.error(formatError(err, "Errore durante l'aggiornamento del docente."));
    },
  });
}

/**
 * Hook per eliminare un docente
 *
 * Stati: isPending (o isLoading) = true durante la delete → disabilita pulsante / conferma
 *
 * @returns Mutation con mutate, mutateAsync, isPending, error, ecc.
 */
export function useDeleteDocente() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (id: number) => {
      const response = await api.delete<{ message: string; id: number }>(
        `/docenti/delete-docente/${id}`
      );
      return response.data;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['docenti'] });
      toast.success("Docente eliminato con successo");
    },
    onError: (err) => {
      toast.error(formatError(err, "Errore durante l'eliminazione del docente."));
    },
  });
}

/**
 * Hook per eliminare tutti i docenti del tenant
 *
 * Stati: isPending = true durante la delete → disabilita pulsante / conferma
 *
 * @returns Mutation con mutate, mutateAsync, isPending, error, ecc.
 */
export function useDeleteAllDocenti() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async () => {
      const response = await api.delete<{ message: string; deletedCount: number }>(
        '/docenti/delete-all'
      );
      return response.data;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['docenti'] });
      queryClient.invalidateQueries({ queryKey: ['registrazioni'] });
      toast.success("Tutti i docenti sono stati eliminati con successo");
    },
    onError: (err) => {
      toast.error(formatError(err, "Errore durante l'eliminazione dei docenti."));
    },
  });
}

/**
 * Hook per importare multipli docenti con copie già effettuate
 *
 * Stati: isPending = true durante l'import → disabilita pulsante / mostra loading
 *
 * @returns Mutation con mutate, mutateAsync, isPending, error, ecc.
 */
export function useBulkImportDocenti() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: BulkImportDocenti) => {
      const response = await api.post<{ message: string; data: { docenti: DocenteConRegistrazioni[]; totaleCreati: number; totaleConCopie: number } }>(
        '/docenti/bulk-import',
        data
      );
      return response.data;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['docenti'] });
      queryClient.invalidateQueries({ queryKey: ['registrazioni'] });
      toast.success("Docenti importati con successo");
    },
    onError: (err) => {
      toast.error(formatError(err, "Errore durante l'importazione dei docenti."));
    },
  });
}


export function useExportDocenti() {
  return useMutation({
    mutationFn: async (query: DocentiQuery) => {
      const params = new URLSearchParams();
      if (query.nome) params.append('nome', query.nome);
      if (query.cognome) params.append('cognome', query.cognome);
      if (query.sortField) params.append('sortField', query.sortField);
      if (query.sortOrder) params.append('sortOrder', query.sortOrder ?? 'asc');
      const response = await api.get<Blob>(`/docenti/export?${params.toString()}`, {
        responseType: 'blob',
      });
      return response.data;
    },
    onSuccess: (data) => {
      const blob = new Blob([data], { type: 'text/csv;charset=utf-8;' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a')
      link.href = url
      link.setAttribute('download', `docenti-${new Date().toISOString().slice(0, 10)}.csv`);
      document.body.appendChild(link)
      link.click()
      link.remove()
      window.URL.revokeObjectURL(url)
      toast.success("Esportazione docenti completata");
    },
    onError: (err) => {
      toast.error(formatError(err, "Errore durante l'esportazione dei docenti."));
    },
  })
}
