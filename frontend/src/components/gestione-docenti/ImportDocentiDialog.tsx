import { useState, useRef } from "react";
import { formatError } from "@/lib/utils";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
  DialogClose,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Field, FieldLabel, FieldContent } from "@/components/ui/field";
import { Button } from "@/components/ui/button";
import { useBulkImportDocenti } from "@/hooks/use-docenti";
import { bulkImportDocentiSchema, type BulkImportDocenti } from "@shared/validation";
import { parseExcelFile } from "@/lib/excel-utils.js";

interface ImportDocentiDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function ImportDocentiDialog({ open, onOpenChange }: ImportDocentiDialogProps) {
  const [importPreview, setImportPreview] = useState<BulkImportDocenti["docenti"] | null>(null);
  const [parsingError, setParsingError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const bulkImport = useBulkImportDocenti();
  

  const handleClose = () => {
    setImportPreview(null);
    setParsingError(null);
    bulkImport.reset();
    if (fileInputRef.current) fileInputRef.current.value = "";
    onOpenChange(false);
  };

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setParsingError(null);
    setImportPreview(null);

    const ext = "." + file.name.split(".").pop()?.toLowerCase();
    if (ext !== ".xlsx") {
      setParsingError("Formato file non supportato. Usa un file .xlsx");
      return;
    }

    try {
      const docenti = await parseExcelFile(file);
      const validated = bulkImportDocentiSchema.parse({ docenti });
      setImportPreview(validated.docenti);
    } catch (error) {
      setParsingError(
        error instanceof Error ? error.message : "Errore durante la lettura del file Excel"
      );
    }
  };

  const handleConfirmImport = () => {
    if (!importPreview) return;
    bulkImport.mutate(
      { docenti: importPreview },
      {
        onSuccess: () => {
          setImportPreview(null);
          setParsingError(null);
          if (fileInputRef.current) fileInputRef.current.value = "";
          onOpenChange(false);
        },
      }
    );
  };

  const getImportErrorMessage = (): string | null => {
    if (parsingError) return parsingError;
    if (!bulkImport.isError || !bulkImport.error) return null;
    return formatError(bulkImport.error, "Errore durante l'import.");
  };
  const importErrorMessage = getImportErrorMessage();

  return (
    <Dialog open={open} onOpenChange={(o) => !o && handleClose()}>
      <DialogContent className="max-h-[90vh] md:max-w-7xl overflow-y-auto min-h-[500px]">
        <DialogHeader>
          <DialogTitle>Importa docenti da Excel</DialogTitle>
          <DialogDescription>
            Importa multipli docenti da un file Excel. Le copie già effettuate verranno
            registrate automaticamente.
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-4 py-4">
          <div className="bg-muted/50 p-4 rounded-md space-y-2">
            <h3 className="font-semibold text-sm">Formato file Excel richiesto:</h3>
            <ul className="text-sm space-y-1 list-disc list-inside text-muted-foreground">
              <li>La prima riga deve contenere le intestazioni delle colonne</li>
              <li>
                <strong>Colonne obbligatorie:</strong> nome, cognome, limiteCopie
              </li>
              <li>
                <strong>Colonne opzionali:</strong> copieEffettuate (default: 0), note
              </li>
              <li>
                Le colonne possono essere scritte in minuscolo, maiuscolo o con spazi (es.
                &quot;Limite Copie&quot;, &quot;Copie Effettuate&quot;)
              </li>
              <li>limiteCopie e copieEffettuate devono essere numeri</li>
              <li>Le copie effettuate non possono superare il limite di copie</li>
            </ul>
            <div className="mt-3 text-xs bg-background p-2 rounded border">
              <strong>Esempio:</strong>
              <table className="mt-2 w-full text-xs border-collapse">
                <thead>
                  <tr className="border-b">
                    <th className="text-left p-1">nome</th>
                    <th className="text-left p-1">cognome</th>
                    <th className="text-left p-1">limiteCopie</th>
                    <th className="text-left p-1">copieEffettuate</th>
                    <th className="text-left p-1">note</th>
                  </tr>
                </thead>
                <tbody>
                  <tr className="border-b">
                    <td className="p-1">Mario</td>
                    <td className="p-1">Rossi</td>
                    <td className="p-1">100</td>
                    <td className="p-1">50</td>
                    <td className="p-1">Copie già effettuate</td>
                  </tr>
                  <tr>
                    <td className="p-1">Luigi</td>
                    <td className="p-1">Verdi</td>
                    <td className="p-1">200</td>
                    <td className="p-1">0</td>
                    <td className="p-1"></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <Field>
            <FieldLabel htmlFor="excel-file">Seleziona file Excel</FieldLabel>
            <FieldContent>
              <Input
                id="excel-file"
                type="file"
                accept=".xlsx"
                ref={fileInputRef}
                onChange={handleFileSelect}
                disabled={bulkImport.isPending}
              />
            </FieldContent>
          </Field>

          {importErrorMessage && (
            <div className="p-3 bg-destructive/10 border border-destructive/20 rounded-md">
              <p className="text-sm text-destructive">{importErrorMessage}</p>
            </div>
          )}

          {importPreview && importPreview.length > 0 && (
            <div className="space-y-2">
              <h3 className="font-semibold text-sm">
                Anteprima dati ({importPreview.length} docente/i):
              </h3>
              <div className="border rounded-md max-h-64 overflow-y-auto">
                <table className="w-full text-sm">
                  <thead className="bg-muted sticky top-0">
                    <tr>
                      <th className="text-left p-2 border-r">Nome</th>
                      <th className="text-left p-2 border-r">Cognome</th>
                      <th className="text-left p-2 border-r">Limite Copie</th>
                      <th className="text-left p-2 border-r">Copie Effettuate</th>
                      <th className="text-left p-2">Note</th>
                    </tr>
                  </thead>
                  <tbody>
                    {importPreview.map((docente: BulkImportDocenti["docenti"][number], index: number) => (
                      <tr key={index} className="border-b">
                        <td className="p-2 border-r">{docente.nome}</td>
                        <td className="p-2 border-r">{docente.cognome}</td>
                        <td className="p-2 border-r">{docente.limiteCopie}</td>
                        <td className="p-2 border-r">{docente.copieEffettuate}</td>
                        <td className="p-2">{docente.note || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        <DialogFooter>
          <DialogClose asChild>
            <Button variant="outline" onClick={handleClose} disabled={bulkImport.isPending}>
              Annulla
            </Button>
          </DialogClose>
          <Button
            variant="default"
            onClick={handleConfirmImport}
            disabled={!importPreview || importPreview.length === 0 || bulkImport.isPending}
          >
            {bulkImport.isPending ? "Import in corso…" : "Conferma import"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
