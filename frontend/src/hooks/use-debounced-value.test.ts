import { act, renderHook } from "@testing-library/react"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { useDebouncedValue } from "./use-debounced-value"

describe("useDebouncedValue", () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })
  afterEach(() => {
    vi.useRealTimers()
  })

  it("parte dal valore iniziale", () => {
    const { result } = renderHook(() => useDebouncedValue("ciao", 300))
    expect(result.current).toBe("ciao")
  })

  it("si aggiorna solo dopo il ritardo dall'ultimo cambio", () => {
    const { result, rerender } = renderHook(({ value }) => useDebouncedValue(value, 300), {
      initialProps: { value: "m" },
    })

    rerender({ value: "ma" })
    act(() => vi.advanceTimersByTime(200))
    rerender({ value: "mar" })
    act(() => vi.advanceTimersByTime(200))
    // Sono passati 400 ms in tutto, ma solo 200 dall'ultimo cambio
    expect(result.current).toBe("m")

    act(() => vi.advanceTimersByTime(100))
    expect(result.current).toBe("mar")
  })
})
