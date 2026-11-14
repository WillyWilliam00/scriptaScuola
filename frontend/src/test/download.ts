import { Blob as NodeBlob } from "node:buffer"
import { vi } from "vitest"

/**
 * Prepara jsdom per i download (export CSV):
 * - il Blob di jsdom non ha `stream()`, che serve a MSW per le risposte con responseType "blob":
 *   per questi test usiamo il Blob di Node (vi.unstubAllGlobals lo rimette a posto)
 * - jsdom non ha URL.createObjectURL e non sa aprire un link a un blob
 * Restituisce le spie per controllare il file scaricato.
 */
export function stubDownload() {
  vi.stubGlobal("Blob", NodeBlob)
  if (!URL.createObjectURL) URL.createObjectURL = () => ""
  if (!URL.revokeObjectURL) URL.revokeObjectURL = () => {}
  const createObjectURL = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:finto")
  const revokeObjectURL = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {})
  const nomiScaricati: (string | null)[] = []
  vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
    nomiScaricati.push(this.getAttribute("download"))
  })
  return { createObjectURL, revokeObjectURL, nomiScaricati }
}
