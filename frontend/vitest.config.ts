import { defineConfig, mergeConfig } from "vitest/config"
import viteConfig from "./vite.config"

// Riusa la configurazione di Vite (plugin e alias `@` e `@shared`) e aggiunge solo la parte dei test.
export default mergeConfig(
  viteConfig,
  defineConfig({
    test: {
      // jsdom simula il browser: DOM, localStorage, window.location
      environment: "jsdom",
      setupFiles: ["./src/test/setup.ts"],
      include: ["src/**/*.test.{ts,tsx}"],
      // Ogni test parte con mock e spie puliti
      clearMocks: true,
      restoreMocks: true,
      unstubGlobals: true,
      coverage: {
        provider: "v8",
        // Per ora solo il report, senza soglie: si fissano quando lib/, store/ e hooks/ sono coperti
        reporter: ["text", "html"],
        include: ["src/lib/**", "src/store/**", "src/hooks/**", "src/components/**"],
        exclude: ["src/components/ui/**", "src/test/**", "**/*.test.{ts,tsx}"],
      },
    },
  }),
)
