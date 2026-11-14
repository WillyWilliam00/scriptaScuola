import ExcelJS from "exceljs"
import { describe, expect, it } from "vitest"
import { parseExcelFile } from "./excel-utils"

/** Crea un vero file .xlsx in memoria con le righe date (la prima è l'intestazione). */
async function creaFileExcel(righe: unknown[][], nome = "docenti.xlsx"): Promise<File> {
  const workbook = new ExcelJS.Workbook()
  const foglio = workbook.addWorksheet("Docenti")
  righe.forEach((riga) => foglio.addRow(riga))
  const buffer = await workbook.xlsx.writeBuffer()
  return new File([buffer], nome, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  })
}

describe("parseExcelFile", () => {
  it("legge un file valido e normalizza i valori", async () => {
    const file = await creaFileExcel([
      ["Nome", "Cognome", "Limite Copie", "Copie Effettuate", "Note"],
      ["  Mario ", "Rossi", 100, 20, " primo anno "],
      ["Anna", "Bianchi", 50, 0, null],
    ])

    expect(await parseExcelFile(file)).toEqual([
      { nome: "Mario", cognome: "Rossi", limiteCopie: 100, copieEffettuate: 20, note: "primo anno" },
      { nome: "Anna", cognome: "Bianchi", limiteCopie: 50, copieEffettuate: 0, note: undefined },
    ])
  })

  it("accetta nomi di colonna alternativi (limite, copie, con trattini bassi)", async () => {
    const file = await creaFileExcel([
      ["NOME", "cognome", "limite", "copie_effettuate"],
      ["Luca", "Verdi", 30, 5],
    ])

    expect(await parseExcelFile(file)).toEqual([
      { nome: "Luca", cognome: "Verdi", limiteCopie: 30, copieEffettuate: 5, note: undefined },
    ])
  })

  it("senza colonne limite e copie le considera 0", async () => {
    const file = await creaFileExcel([
      ["Nome", "Cognome"],
      ["Luca", "Verdi"],
    ])

    expect(await parseExcelFile(file)).toEqual([
      { nome: "Luca", cognome: "Verdi", limiteCopie: 0, copieEffettuate: 0, note: undefined },
    ])
  })

  it("salta le righe completamente vuote", async () => {
    const file = await creaFileExcel([
      ["Nome", "Cognome", "Limite"],
      ["Luca", "Verdi", 10],
      [],
      ["Anna", "Neri", 20],
    ])

    expect((await parseExcelFile(file)).map((d) => d.nome)).toEqual(["Luca", "Anna"])
  })

  it("rifiuta un file con la sola intestazione", async () => {
    const file = await creaFileExcel([["Nome", "Cognome"]])
    await expect(parseExcelFile(file)).rejects.toThrow("Il file Excel è vuoto")
  })

  it("rifiuta una riga senza cognome, indicando il numero di riga", async () => {
    const file = await creaFileExcel([
      ["Nome", "Cognome", "Limite"],
      ["Luca", "Verdi", 10],
      ["Anna", "", 20],
    ])
    await expect(parseExcelFile(file)).rejects.toThrow("Riga 3: nome e cognome sono obbligatori")
  })

  it("rifiuta un file senza la colonna cognome", async () => {
    const file = await creaFileExcel([
      ["Nome", "Limite"],
      ["Luca", 10],
    ])
    await expect(parseExcelFile(file)).rejects.toThrow("Riga 2: nome e cognome sono obbligatori")
  })

  it("rifiuta un limite non numerico o negativo", async () => {
    const testo = await creaFileExcel([
      ["Nome", "Cognome", "Limite"],
      ["Luca", "Verdi", "tanti"],
    ])
    await expect(parseExcelFile(testo)).rejects.toThrow("Riga 2: limiteCopie deve essere un numero >= 0")

    const negativo = await creaFileExcel([
      ["Nome", "Cognome", "Limite"],
      ["Luca", "Verdi", -1],
    ])
    await expect(parseExcelFile(negativo)).rejects.toThrow("Riga 2: limiteCopie deve essere un numero >= 0")
  })

  it("rifiuta copie effettuate negative", async () => {
    const file = await creaFileExcel([
      ["Nome", "Cognome", "Limite", "Copie"],
      ["Luca", "Verdi", 10, -3],
    ])
    await expect(parseExcelFile(file)).rejects.toThrow("Riga 2: copieEffettuate deve essere un numero >= 0")
  })

  it("rifiuta copie effettuate oltre il limite", async () => {
    const file = await creaFileExcel([
      ["Nome", "Cognome", "Limite", "Copie"],
      ["Luca", "Verdi", 10, 11],
    ])
    await expect(parseExcelFile(file)).rejects.toThrow("Riga 2: copieEffettuate (11) supera limiteCopie (10)")
  })

  it("legge il risultato delle formule", async () => {
    const workbook = new ExcelJS.Workbook()
    const foglio = workbook.addWorksheet("Docenti")
    foglio.addRow(["Nome", "Cognome", "Limite"])
    foglio.addRow(["Luca", "Verdi", { formula: "10*2", result: 20 }])
    const file = new File([await workbook.xlsx.writeBuffer()], "formule.xlsx")

    expect((await parseExcelFile(file))[0].limiteCopie).toBe(20)
  })

  it("rifiuta un file che non è un Excel", async () => {
    const file = new File(["nome,cognome\nLuca,Verdi"], "docenti.csv", { type: "text/csv" })
    await expect(parseExcelFile(file)).rejects.toThrow()
  })
})
