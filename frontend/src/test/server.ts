import { setupServer } from "msw/node"

/**
 * Server MSW condiviso da tutti i test: intercetta le chiamate HTTP (axios compreso)
 * e risponde con gli handler che ogni test registra con `server.use(...)`.
 * Nessun handler di default: una chiamata non prevista fa fallire il test (vedi setup.ts).
 */
export const server = setupServer()

/** Prefisso delle chiamate API: in jsdom `/api` diventa http://localhost:3000/api. */
export const API = "*/api"
