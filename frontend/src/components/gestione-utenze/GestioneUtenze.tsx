import { Suspense, useState, useCallback, useMemo } from "react";
import HeaderSection from "@/components/layout/HeaderSection";
import { GestioneUtenzeContent } from "./GestioneUtenzeContent";
import { KeyIcon, File } from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";
import type { Utente } from "@shared/types";
import type { UtentiQuery } from "@shared/validation";
import { ServerFiltersBar } from "@/components/common/ServerFiltersBar";
import NuovoUtenteModal from "./NuovoUtenteModal";
import { ViewUtenteModal } from "./ViewUtenteModal";
import { DeleteUtenteAlertDialog } from "./DeleteUtenteAlertDialog";
import EditUtenteModal from "./EditUtenteModal";
import { RichiesteResetCard } from "./RichiesteResetCard";
import { Button } from "@/components/ui/button";
import ErrorBoundary from "../common/ErrorBoundary";
import ErrorFallback from "../common/ErrorFallback";
import { useQueryClient } from "@tanstack/react-query";
import { useExportUtenti } from "@/hooks/use-utenti.js";









export default function GestioneUtenze() {
  const defaultQuery: UtentiQuery = {
    page: 1,
    pageSize: 10,
    sortOrder: "asc",
    sortField: "ruolo",
  };
  const [utentiQuery, setUtentiQuery] = useState<UtentiQuery>(defaultQuery);
  const {mutate: exportUtenti, isPending: isExporting} = useExportUtenti();
  const [selectedUtente, setSelectedUtente] = useState<Utente | null>(null);
  const [typeForm, setTypeForm] = useState<"delete" | "edit" | "view" | "add" | null>(null);
  const [errorBoundaryKey, setErrorBoundaryKey] = useState(0);
  const queryClient = useQueryClient();
  const handleRetry = () => {
    queryClient.invalidateQueries({ queryKey: ["utenti"] });
    setErrorBoundaryKey((k) => k + 1);
  };


  const handlePageChange = useCallback((page: number) => {
    setUtentiQuery((prev: UtentiQuery) => ({ ...prev, page }));
  }, []);

  const handleFilterChange = useCallback((filters: Record<string, string | undefined>) => {
    setUtentiQuery((prev: UtentiQuery) => ({ ...prev, ...filters, page: 1 }));
  }, []);

  const handlePageSizeChange = useCallback((pageSize: string) => {
    setUtentiQuery((prev: UtentiQuery) => ({ ...prev, pageSize: Number(pageSize), page: 1 }));
  }, []);

  const handleSortChange = useCallback(
    (sortField: UtentiQuery['sortField'], sortOrder: 'asc' | 'desc') => {
      setUtentiQuery((prev: UtentiQuery) => ({
        ...prev,
        sortField,
        sortOrder,
        page: 1,
      }));
    },
    []
  );

  const handleCloseForm = useCallback(() => {
    setSelectedUtente(null);
    setTypeForm(null);
  }, []);
  const handleOpenForm = (type: "view" | "edit" | "delete", utente: Utente) => {
    setSelectedUtente(utente);
    setTypeForm(type);
  }
  const handleOpenAddForm = () => {
    setTypeForm("add");
    setSelectedUtente(null);
  }

  const handleExport = useCallback(() => {
    exportUtenti(utentiQuery)
  }, [exportUtenti, utentiQuery])

  const filterValues = useMemo(
    () => ({ identifier: utentiQuery.identifier }),
    [utentiQuery.identifier]
  );

  return (
    <div>
      <HeaderSection title="Gestione Utenze" icon={KeyIcon} />

      <div className="w-full  xl:mt-10 p-2 lg:p-4">
        <div className="flex flex-row gap-2 w-full justify-between items-end">
          <ServerFiltersBar
            type="utenze"
            filterValues={filterValues}
            onFilterChange={handleFilterChange}
          />
          <Button
            className="mt-4"
            variant="default"
            onClick={handleExport}
            disabled={isExporting}
          >
            {isExporting ? "Esportazione..." : "Esporta in Excel"}
            <HugeiconsIcon icon={File} strokeWidth={2} className="size-4" />
          </Button>
          <NuovoUtenteModal typeForm={typeForm} onClose={handleCloseForm} onOpenAddForm={handleOpenAddForm} />
        </div>
        <RichiesteResetCard />
        <ErrorBoundary
          key={errorBoundaryKey}
          fallback={  
            <ErrorFallback onRetry={handleRetry} />
          }
        >
        <Suspense
          fallback={
            <div className="overflow-hidden rounded-md border mt-6 p-8 text-center text-muted-foreground">
              Caricamento utenze…
            </div>
          }
        >
          <GestioneUtenzeContent
            query={utentiQuery}
            onPageChange={handlePageChange}
            onPageSizeChange={handlePageSizeChange}
            onSortChange={handleSortChange}
            onView={(utente) => {
              handleOpenForm("view", utente);
            }}
            onEdit={(utente) => {
              handleOpenForm("edit", utente);
            }}
            onDelete={(utente) => {
              handleOpenForm("delete", utente);
            }}
          />
        </Suspense>
        </ErrorBoundary>
      </div>

      <ViewUtenteModal selectedUtente={selectedUtente} typeForm={typeForm} onClose={handleCloseForm} />
      <EditUtenteModal selectedUtente={selectedUtente} typeForm={typeForm} onClose={handleCloseForm} />
      <DeleteUtenteAlertDialog selectedUtente={selectedUtente} typeForm={typeForm} onClose={handleCloseForm} />
    </div>
  );
}
