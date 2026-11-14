import type { ReactNode } from "react"
import { QueryClient, QueryClientProvider } from "@tanstack/react-query"
import type { JwtPayload } from "@shared/types"
import { useAuthStore } from "@/store/auth-store"

/** QueryClient per i test: nessun retry (un errore arriva subito) e nessuna cache tra un test e l'altro. */
export function createTestQueryClient() {
  return new QueryClient({
    defaultOptions: {
      queries: { retry: false, gcTime: Infinity },
      mutations: { retry: false },
    },
  })
}

/** Wrapper per renderHook/render con il provider di TanStack Query. */
export function createWrapper(queryClient = createTestQueryClient()) {
  function Wrapper({ children }: { children: ReactNode }) {
    return <QueryClientProvider client={queryClient}>{children}</QueryClientProvider>
  }
  return { queryClient, wrapper: Wrapper }
}

function base64Url(value: object): string {
  return btoa(JSON.stringify(value)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "")
}

/** Crea un JWT finto (firma non valida, il frontend non la verifica) con il payload dato. */
export function makeJwt(payload: Partial<JwtPayload> = {}): string {
  const full: JwtPayload = {
    userId: "utente-1",
    ruolo: "admin",
    istitutoId: 1,
    exp: Math.floor(Date.now() / 1000) + 15 * 60,
    ...payload,
  }
  return `${base64Url({ alg: "HS256", typ: "JWT" })}.${base64Url(full)}.firma`
}

/** Riporta lo store di autenticazione allo stato iniziale (utente non collegato). */
export function resetAuthStore() {
  useAuthStore.setState({
    token: null,
    utente: null,
    istituto: null,
    jwtPayload: null,
    isAuthenticated: false,
    isInitializing: true,
  })
}

/** Simula un utente collegato con un access token valido. */
export function loginAs(ruolo: JwtPayload["ruolo"] = "admin", token = makeJwt({ ruolo })) {
  useAuthStore.getState().login({
    token,
    utente: {
      id: "utente-1",
      email: ruolo === "admin" ? "admin@scuola.it" : null,
      username: ruolo === "collaboratore" ? "mario" : null,
      ruolo,
      deveCambiarePassword: false,
    },
    istituto: { id: 1, nome: "Istituto di prova", codiceIstituto: "ABC123" },
  })
  return token
}

/** Risposta paginata vuota o con i dati dati, nel formato del backend. */
export function paginated<T>(data: T[], page = 1, hasNextPage = false) {
  return {
    data,
    pagination: {
      page,
      pageSize: 20,
      totalItems: data.length,
      totalPages: hasNextPage ? page + 1 : page,
      hasNextPage,
      hasPreviousPage: page > 1,
    },
  }
}
