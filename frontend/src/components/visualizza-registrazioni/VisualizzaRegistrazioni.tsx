import { FileIcon } from "@hugeicons/core-free-icons";
import HeaderSection from "@/components/layout/HeaderSection";
import { VisualizzaRegistrazioniContent } from "./VisualizzaRegistrazioniContent";
import { type Registrazioni } from "@/components/table/columns";
import { useState, Suspense, useCallback, useMemo } from "react";
import { ViewRegistrazioneModal } from "./ViewRegistrazioneModal";
import EditRegistrazioneModal from "./EditRegistrazioneModal";
import { DeleteRegistrazioneAlertDialog } from "./DeleteRegistrazioneAlertDialog";
import ErrorBoundary from "@/components/common/ErrorBoundary";
import ErrorFallback from "@/components/common/ErrorFallback";
import { ServerFiltersBar } from "@/components/common/ServerFiltersBar";
import { useQueryClient } from "@tanstack/react-query";
import type { RegistrazioniCopieQuery } from "@shared/validation";
import type { RegistrazioniCopieSort } from "@shared/types";
import { Button } from "@/components/ui/button";
import { HugeiconsIcon } from "@hugeicons/react";
import { useExportRegistrazioni } from "@/hooks/use-registrazioni";


type TypeForm = "view" | "edit" | "delete" | null;

const defaultQuery: RegistrazioniCopieQuery = {
  page: 1,
  pageSize: 10,
  sortOrder: "desc",
  sortField: "createdAt",
};

export default function VisualizzaRegistrazioni() {
  const queryClient = useQueryClient();
  const {mutate: exportRegistrazioni, isPending: isExporting} = useExportRegistrazioni();
  const [typeForm, setTypeForm] = useState<TypeForm>(null);
  const [selectedRegistrazione, setSelectedRegistrazione] = useState<Registrazioni | null>(null);
  const [registrazioniQuery, setRegistrazioniQuery] = useState<RegistrazioniCopieQuery>(defaultQuery);
  const [errorBoundaryKey, setErrorBoundaryKey] = useState(0);

  const handleRetry = useCallback(() => {
    queryClient.invalidateQueries({ queryKey: ["registrazioni"] });
    setErrorBoundaryKey((k) => k + 1);
  }, [queryClient]);

  const handleOpenForm = useCallback((type: "view" | "edit" | "delete", registrazione: Registrazioni) => {
    setSelectedRegistrazione(registrazione);
    setTypeForm(type);
  }, []);

  const handleCloseForm = useCallback(() => {
    setSelectedRegistrazione(null);
    setTypeForm(null);
  }, []);

  const handlePageChange = useCallback((page: number) => {
    setRegistrazioniQuery((prev) => ({ ...prev, page }));
  }, []);

  const handlePageSizeChange = useCallback((pageSize: string) => {
    setRegistrazioniQuery((prev) => ({
      ...prev,
      pageSize: Number(pageSize),
      page: 1,
    }));
  }, []);

  const handleSortChange = useCallback((sortField: RegistrazioniCopieSort['field'], sortOrder: RegistrazioniCopieSort['order']) => {
    setRegistrazioniQuery((prev) => ({
      ...prev,
      sortField,
      sortOrder,
      page: 1,
    }));
  }, []);

  const handleFilterChange = useCallback((filters: Record<string, string | undefined>) => {
    setRegistrazioniQuery((prev) => {
      const next: RegistrazioniCopieQuery = {
        ...prev,
        page: 1,
        docenteNome: filters.docenteNome?.trim() || undefined,
        docenteCognome: filters.docenteCognome?.trim() || undefined,
        utenteIdentifier: filters.utenteIdentifier?.trim() || undefined,
      };
      const copieStr = filters.copieEffettuate?.trim();
      next.copieEffettuate =
        copieStr === "" ? undefined : (Number(copieStr) as number);
      if (next.copieEffettuate !== undefined && Number.isNaN(next.copieEffettuate)) {
        next.copieEffettuate = undefined;
      }
      return next;
    });
  }, []);

 const handleExport = useCallback(() => {
  exportRegistrazioni(registrazioniQuery)
 }, [exportRegistrazioni, registrazioniQuery])

  const filterValues = useMemo(
    () => ({
      docenteNome: registrazioniQuery.docenteNome ?? "",
      docenteCognome: registrazioniQuery.docenteCognome ?? "",
      copieEffettuate:
        registrazioniQuery.copieEffettuate !== undefined
          ? String(registrazioniQuery.copieEffettuate)
          : "",
      utenteIdentifier: registrazioniQuery.utenteIdentifier ?? "",
    }),
    [
      registrazioniQuery.docenteNome,
      registrazioniQuery.docenteCognome,
      registrazioniQuery.copieEffettuate,
      registrazioniQuery.utenteIdentifier,
    ]
  );

  return (
    <div>
      <HeaderSection title="Visualizza Registrazioni Copie" icon={FileIcon} />

      <div className="w-full xl:mt-10 p-2 lg:p-4">
        <div className="flex flex-row gap-2 w-full justify-between items-end">
          <ServerFiltersBar
            type="registrazioni"
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
            <HugeiconsIcon icon={FileIcon} strokeWidth={2} className="size-4" />
          </Button>
        </div>
        <ErrorBoundary
          key={errorBoundaryKey}
          fallback={<ErrorFallback onRetry={handleRetry} />}
        >
          <Suspense
            fallback={
              <div className="w-full mt-4 p-8 text-center text-muted-foreground">
                Caricamento registrazioni…
              </div>
            }
          >
            <VisualizzaRegistrazioniContent
              query={registrazioniQuery}
              onPageChange={handlePageChange}
              onPageSizeChange={handlePageSizeChange}
              onView={(r) => handleOpenForm("view", r)}
              onEdit={(r) => handleOpenForm("edit", r)}
              onDelete={(r) => handleOpenForm("delete", r)}
              handleSortChange={handleSortChange}
            />
          </Suspense>
        </ErrorBoundary>
      </div>

      <ViewRegistrazioneModal
        selectedRegistrazione={selectedRegistrazione}
        typeForm={typeForm}
        onClose={handleCloseForm}
      />
      <EditRegistrazioneModal
        selectedRegistrazione={selectedRegistrazione}
        typeForm={typeForm}
        onClose={handleCloseForm}
      />
      <DeleteRegistrazioneAlertDialog
        selectedRegistrazione={selectedRegistrazione}
        typeForm={typeForm}
        onClose={handleCloseForm}
      />
    </div>
  );
}
