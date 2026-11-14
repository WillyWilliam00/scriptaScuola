import { FileIcon } from "@hugeicons/core-free-icons";
import HeaderSection from "@/components/layout/HeaderSection";
import { GestioneDocentiContent } from "./GestioneDocentiContent";
import { useState, Suspense, useCallback, useMemo } from "react";
import { type DocentiQuery } from "@shared/validation";
import { ImportDocentiDialog } from "./ImportDocentiDialog";
import { ServerFiltersBar } from "@/components/common/ServerFiltersBar";
import { EliminaTuttiDocentiDialog } from "./EliminaTuttiDocentiDialog";
import { ResetConteggioDialog } from "./ResetConteggioDialog";
import NuovoDocenteModal from "./NuovoDocenteModal";
import EditDocenteModal from "./EditDocenteModal";
import { ViewDocenteModal } from "./ViewDocenteModal";
import { EliminaDocenteDialog } from "./EliminaDocenteDialog";
import { type Docenti } from "@/components/table/columns";
import { Button } from "@/components/ui/button";
import { HugeiconsIcon } from "@hugeicons/react";
import ErrorFallback from "../common/ErrorFallback";
import ErrorBoundary from "../common/ErrorBoundary";
import { useQueryClient } from "@tanstack/react-query";
import { useExportDocenti } from "@/hooks/use-docenti";

type TypeForm = "add" | "edit" | "view" | "delete" | null;

const defaultQuery: DocentiQuery = { page: 1, pageSize: 10, sortOrder: "desc", sortField: "createdAt" };

export default function GestioneDocenti() {
  const [typeForm, setTypeForm] = useState<TypeForm>(null);
  const [errorBoundaryKey, setErrorBoundaryKey] = useState(0);
  const [selectedDocente, setSelectedDocente] = useState<Docenti | null>(null);
  const [isImportDialogOpen, setIsImportDialogOpen] = useState(false);
  const [docentiQuery, setDocentiQuery] = useState<DocentiQuery>(defaultQuery);
  const queryClient = useQueryClient();
  const { mutate: exportDocenti, isPending: isExporting } = useExportDocenti();
  const handleRetry = () => {
    queryClient.invalidateQueries({ queryKey: ["docenti"] });
    setErrorBoundaryKey((k) => k + 1);
  };

  const handlePageChange = useCallback((page: number) => {
    setDocentiQuery((prev: DocentiQuery) => ({ ...prev, page }));
  }, []);

  const handleFilterChange = useCallback((filters: Record<string, string | undefined>) => {
    setDocentiQuery((prev: DocentiQuery) => ({
      ...prev,
      nome: filters.nome,
      cognome: filters.cognome,
      page: 1,
    }));
  }, []);

  const handlePageSizeChange = useCallback((pageSize: string) => {
    setDocentiQuery((prev: DocentiQuery) => ({
      ...prev,
      pageSize: Number(pageSize),
      page: 1,
    }));
  }, []);

  const handleSortChange = useCallback(
    (sortField: DocentiQuery['sortField'], sortOrder: 'asc' | 'desc') => {
      setDocentiQuery((prev: DocentiQuery) => ({
        ...prev,
        sortField,
        sortOrder,
        page: 1,
      }));
    },
    []
  );

  const handleCloseForm = useCallback(() => {
    setSelectedDocente(null);
    setTypeForm(null);
  }, []);

  const handleOpenForm = useCallback((type: "view" | "edit" | "delete", docente: Docenti) => {
    setSelectedDocente(docente);
    setTypeForm(type);
  }, []);

  const handleOpenAddForm = useCallback(() => {
    setTypeForm("add");
    setSelectedDocente(null);
  }, []);

  const handleImportClick = useCallback(() => setIsImportDialogOpen(true), []);

  const handleExport = useCallback(() => {
    exportDocenti(docentiQuery);
  }, [exportDocenti, docentiQuery]);

  const filterValues = useMemo(
    () => ({ nome: docentiQuery.nome, cognome: docentiQuery.cognome }),
    [docentiQuery.nome, docentiQuery.cognome]
  );

  return (
    <div>
      <div className="flex justify-end pt-5 pr-5 w-full gap-2">
        <EliminaTuttiDocentiDialog />
        <ResetConteggioDialog />
      </div>
      <HeaderSection title="Gestione Docenti" icon={FileIcon} />

      <div className="px-4">
        <div className="flex flex-row gap-2 w-full justify-between items-end">
          <ServerFiltersBar
            type="docenti"
            filterValues={filterValues}
            onFilterChange={handleFilterChange}
          />

          <div className="flex flex-row gap-2">
            <NuovoDocenteModal
              typeForm={typeForm}
              onClose={handleCloseForm}
              onOpenAddForm={handleOpenAddForm}
            />

            <Button className="" variant="outline" onClick={handleImportClick}>
              Importa file
              <HugeiconsIcon icon={FileIcon} strokeWidth={2} className="size-4" />
            </Button>

            <Button className="" variant="default" onClick={handleExport} disabled={isExporting}>
              {isExporting ? "Esportazione..." : "Esporta in Excel"}
              <HugeiconsIcon icon={FileIcon} strokeWidth={2} className="size-4" />
            </Button>
          </div>
        </div>
        <ErrorBoundary
          key={errorBoundaryKey}
          fallback={
            <ErrorFallback onRetry={handleRetry} />
          }
        >
        <Suspense fallback={<div className="w-full mt-4 p-8 text-center text-muted-foreground">Caricamento docenti…</div>}>
          <GestioneDocentiContent
            query={docentiQuery}
            onView={(docente) => handleOpenForm("view", docente)}
            onEdit={(docente) => handleOpenForm("edit", docente)}
            onDelete={(docente) => handleOpenForm("delete", docente)}
            onPageChange={handlePageChange}
            onPageSizeChange={handlePageSizeChange}
            onSortChange={handleSortChange}
          />
        </Suspense>
        </ErrorBoundary>
      </div>

      <ViewDocenteModal selectedDocente={selectedDocente} typeForm={typeForm} onClose={handleCloseForm} />
      <EditDocenteModal selectedDocente={selectedDocente} typeForm={typeForm} onClose={handleCloseForm} />
      <EliminaDocenteDialog selectedDocente={selectedDocente} typeForm={typeForm} onClose={handleCloseForm} />
      <ImportDocentiDialog open={isImportDialogOpen} onOpenChange={setIsImportDialogOpen} />
    </div>
  );
}
