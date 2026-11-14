import { afterEach, describe, expect, it, vi } from "vitest"
import { makeJwt } from "@/test/utils"
import { decodeJWT, isTokenExpired } from "./jwt-utils"

describe("decodeJWT", () => {
  it("legge il payload di un token valido", () => {
    const token = makeJwt({ userId: "abc", ruolo: "collaboratore", istitutoId: 7, exp: 2000000000 })
    expect(decodeJWT(token)).toEqual({ userId: "abc", ruolo: "collaboratore", istitutoId: 7, exp: 2000000000 })
  })

  it("decodifica correttamente i caratteri accentati", () => {
    const payload = { nome: "Città" }
    const parte = btoa(String.fromCharCode(...new TextEncoder().encode(JSON.stringify(payload))))
    expect(decodeJWT(`x.${parte}.y`)).toEqual(payload)
  })

  it("restituisce null per un token senza payload", () => {
    expect(decodeJWT("solo-una-parte")).toBeNull()
  })

  it("restituisce null per un payload che non è JSON", () => {
    vi.spyOn(console, "error").mockImplementation(() => {})
    expect(decodeJWT("a.bm9uLWpzb24.c")).toBeNull()
  })
})

describe("isTokenExpired", () => {
  afterEach(() => {
    vi.useRealTimers()
  })

  it("è false per un token che scade nel futuro", () => {
    expect(isTokenExpired(makeJwt({ exp: Math.floor(Date.now() / 1000) + 60 }))).toBe(false)
  })

  it("è true per un token già scaduto", () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date("2026-01-01T12:00:00Z"))
    const exp = Math.floor(new Date("2026-01-01T11:59:00Z").getTime() / 1000)
    expect(isTokenExpired(makeJwt({ exp }))).toBe(true)
  })

  it("è true per un token senza scadenza o malformato", () => {
    vi.spyOn(console, "error").mockImplementation(() => {})
    expect(isTokenExpired(makeJwt({ exp: undefined }))).toBe(true)
    expect(isTokenExpired("non-un-token")).toBe(true)
  })
})
