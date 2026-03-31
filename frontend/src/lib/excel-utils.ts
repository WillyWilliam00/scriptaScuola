import ExcelJS from "exceljs";
import type { BulkImportDocenti } from "../../../shared/validation.js";

/**
 * Legge un file Excel e restituisce un array di docenti nel formato atteso.
 * Accetta colonne con nomi flessibili (es. "limite", "copie effettuate").
 * @throws Error se il file è vuoto, formato non valido o validazione fallisce
 */
export async function parseExcelFile(file: File): Promise<BulkImportDocenti["docenti"]> {
  try {
    const data = await file.arrayBuffer();
    if (!data) {
      throw new Error("Nessun dato letto dal file");
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);

    if (workbook.worksheets.length === 0) {
      throw new Error("Il file Excel è vuoto (nessun foglio)");
    }

    const worksheet = workbook.worksheets[0];
    const jsonData: Record<string, unknown>[] = [];
    const headers: Record<number, string> = {};

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        row.eachCell((cell, colNumber) => {
          headers[colNumber] = cell.value?.toString() || `Column${colNumber}`;
        });
      } else {
        const rowData: Record<string, unknown> = {};
        let hasValue = false;

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const header = headers[colNumber];
          if (header) {
            let value = cell.value;
            
            // Resolve formula and rich text
            if (value && typeof value === 'object') {
              if ('result' in value) {
                value = (value as ExcelJS.CellFormulaValue).result;
              } else if ('richText' in value) {
                value = (value as ExcelJS.CellRichTextValue).richText.map((rt: any) => rt.text).join('');
              } else if ('text' in value) {
                value = (value as ExcelJS.CellHyperlinkValue).text;
              }
            }
            
            rowData[header] = value;
            if (value !== null && value !== undefined && value !== "") {
              hasValue = true;
            }
          }
        });

        if (hasValue) {
          jsonData.push(rowData);
        }
      }
    });

    if (jsonData.length === 0) {
      throw new Error("Il file Excel è vuoto");
    }

    const getColumnValue = (row: Record<string, unknown>, candidates: string[]) => {
      for (const [key, value] of Object.entries(row)) {
        const normalized = key.toLowerCase().replace(/\s|_/g, "");
        if (candidates.includes(normalized)) return value;
      }
      return undefined;
    };

    const docenti = jsonData.map((row, index) => {
      const nomeRaw = getColumnValue(row, ["nome"]);
      const cognomeRaw = getColumnValue(row, ["cognome"]);
      const limiteRaw = getColumnValue(row, ["limitecopie", "limite"]);
      const copieRaw = getColumnValue(row, ["copieeffettuate", "copieeff", "copie"]);
      const noteRaw = getColumnValue(row, ["note"]);

      const nome =
        typeof nomeRaw === "string" ? nomeRaw.trim() : String(nomeRaw ?? "").trim();
      const cognome =
        typeof cognomeRaw === "string"
          ? cognomeRaw.trim()
          : String(cognomeRaw ?? "").trim();
      const limiteCopie = Number(limiteRaw ?? 0);
      const copieEffettuate = Number(copieRaw ?? 0);
      const note =
        typeof noteRaw === "string" ? noteRaw : noteRaw != null ? String(noteRaw) : "";

      if (!nome || !cognome) {
        throw new Error(`Riga ${index + 2}: nome e cognome sono obbligatori`);
      }
      if (isNaN(limiteCopie) || limiteCopie < 0) {
        throw new Error(`Riga ${index + 2}: limiteCopie deve essere un numero >= 0`);
      }
      if (isNaN(copieEffettuate) || copieEffettuate < 0) {
        throw new Error(
          `Riga ${index + 2}: copieEffettuate deve essere un numero >= 0`
        );
      }
      if (copieEffettuate > limiteCopie) {
        throw new Error(
          `Riga ${index + 2}: copieEffettuate (${copieEffettuate}) supera limiteCopie (${limiteCopie})`
        );
      }

      return {
        nome,
        cognome,
        limiteCopie,
        copieEffettuate,
        note: note ? String(note).trim() : undefined,
      };
    });

    return docenti;
  } catch (error) {
    if (error instanceof Error) {
      throw error;
    }
    throw new Error("Errore sconosciuto nella lettura del file");
  }
}
