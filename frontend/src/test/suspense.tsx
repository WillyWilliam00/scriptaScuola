import { Suspense, type ReactNode } from "react"
import { createWrapper } from "./utils"

/** Come createWrapper, ma con un <Suspense> per gli hook useSuspenseQuery. */
export function createSuspenseWrapper() {
  const { queryClient, wrapper: Base } = createWrapper()
  function Wrapper({ children }: { children: ReactNode }) {
    return (
      <Base>
        <Suspense fallback={null}>{children}</Suspense>
      </Base>
    )
  }
  return { queryClient, wrapper: Wrapper }
}
