import { useState, useRef, useCallback } from "react";
import { InputGroup, InputGroupAddon, InputGroupInput } from "@/components/ui/input-group";
import { Field, FieldGroup, FieldLabel } from "@/components/ui/field";
import { FilterIcon, SearchIcon } from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";
import { Select, SelectContent, SelectGroup, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Button } from "@/components/ui/button";

const DEFAULT_FILTER_DEBOUNCE_MS = 300;

export interface ServerFiltersBarProps {
  type: "docenti" | "utenze" | "registrazioni";
  filterValues: Record<string, string | undefined>;
  onFilterChange: (filters: Record<string, string | undefined>) => void;
  filterDebounceMs?: number;
}

export function ServerFiltersBar({
  type,
  filterValues,
  onFilterChange,
  filterDebounceMs = DEFAULT_FILTER_DEBOUNCE_MS,
}: ServerFiltersBarProps) {
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [localNome, setLocalNome] = useState(filterValues.nome ?? "");
  const [localCognome, setLocalCognome] = useState(filterValues.cognome ?? "");
  const [localIdentifier, setLocalIdentifier] = useState(filterValues.identifier ?? "");
  const [localRuolo, setLocalRuolo] = useState(filterValues.ruolo ?? "");
  const [localDocenteNome, setLocalDocenteNome] = useState(filterValues.docenteNome ?? "");
  const [localDocenteCognome, setLocalDocenteCognome] = useState(filterValues.docenteCognome ?? "");
  const [localCopieEffettuate, setLocalCopieEffettuate] = useState(filterValues.copieEffettuate ?? "");
  const [localUtenteIdentifier, setLocalUtenteIdentifier] = useState(filterValues.utenteIdentifier ?? "");
  // Quando i filtri arrivano cambiati dall'esterno (es. reset o navigazione), riallinea i
  // campi locali. Si fa durante il render confrontando con i valori precedenti, come
  // suggerisce React, invece che con un useEffect che causerebbe un render in più.
  const [prevFilterValues, setPrevFilterValues] = useState(filterValues);
  if (prevFilterValues !== filterValues) {
    setPrevFilterValues(filterValues);
    if (prevFilterValues.nome !== filterValues.nome) setLocalNome(filterValues.nome ?? "");
    if (prevFilterValues.cognome !== filterValues.cognome) setLocalCognome(filterValues.cognome ?? "");
    if (prevFilterValues.identifier !== filterValues.identifier) setLocalIdentifier(filterValues.identifier ?? "");
    if (prevFilterValues.ruolo !== filterValues.ruolo) setLocalRuolo(filterValues.ruolo ?? "");
    if (prevFilterValues.docenteNome !== filterValues.docenteNome) setLocalDocenteNome(filterValues.docenteNome ?? "");
    if (prevFilterValues.docenteCognome !== filterValues.docenteCognome) setLocalDocenteCognome(filterValues.docenteCognome ?? "");
    if (prevFilterValues.copieEffettuate !== filterValues.copieEffettuate) setLocalCopieEffettuate(filterValues.copieEffettuate ?? "");
    if (prevFilterValues.utenteIdentifier !== filterValues.utenteIdentifier) setLocalUtenteIdentifier(filterValues.utenteIdentifier ?? "");
  }

  const emitFilterChange = useCallback(
    (filters: Record<string, string | undefined>) => {
      if (debounceRef.current) clearTimeout(debounceRef.current);
      debounceRef.current = setTimeout(() => {
        debounceRef.current = null;
        onFilterChange(filters);
      }, filterDebounceMs);
    },
    [onFilterChange, filterDebounceMs]
  );

  const handleDocentiNomeChange = useCallback(
    (value: string) => {
      setLocalNome(value);
      emitFilterChange({ nome: value || undefined, cognome: localCognome || undefined });
    },
    [localCognome, emitFilterChange]
  );
  const handleDocentiCognomeChange = useCallback(
    (value: string) => {
      setLocalCognome(value);
      emitFilterChange({ nome: localNome || undefined, cognome: value || undefined });
    },
    [localNome, emitFilterChange]
  );
  const handleUtenzeIdentifierChange = useCallback(
    (value: string) => {
      setLocalIdentifier(value);
      emitFilterChange({ identifier: value.trim() || undefined });
    },
    [emitFilterChange]
  );
  const handleUtenzeRuoloChange = useCallback(
    (value: string | undefined) => {
      if (value === "all") {
        value = undefined;
      }
      setLocalRuolo(value ?? "");
      emitFilterChange({ ruolo: value || undefined });
    },
    [emitFilterChange]
  );

  const handleRegistrazioniDocenteNomeChange = useCallback(
    (value: string) => {
      setLocalDocenteNome(value);
      emitFilterChange({
        docenteNome: value.trim() || undefined,
        docenteCognome: localDocenteCognome.trim() || undefined,
        copieEffettuate: localCopieEffettuate.trim() || undefined,
        utenteIdentifier: localUtenteIdentifier.trim() || undefined,
      });
    },
    [localDocenteCognome, localCopieEffettuate, localUtenteIdentifier, emitFilterChange]
  );
  const handleRegistrazioniDocenteCognomeChange = useCallback(
    (value: string) => {
      setLocalDocenteCognome(value);
      emitFilterChange({
        docenteNome: localDocenteNome.trim() || undefined,
        docenteCognome: value.trim() || undefined,
        copieEffettuate: localCopieEffettuate.trim() || undefined,
        utenteIdentifier: localUtenteIdentifier.trim() || undefined,
      });
    },
    [localDocenteNome, localCopieEffettuate, localUtenteIdentifier, emitFilterChange]
  );
  const handleRegistrazioniCopieChange = useCallback(
    (value: string) => {
      setLocalCopieEffettuate(value);
      emitFilterChange({
        docenteNome: localDocenteNome.trim() || undefined,
        docenteCognome: localDocenteCognome.trim() || undefined,
        copieEffettuate: value.trim() || undefined,
        utenteIdentifier: localUtenteIdentifier.trim() || undefined,
      });
    },
    [localDocenteNome, localDocenteCognome, localUtenteIdentifier, emitFilterChange]
  );
  const handleRegistrazioniUtenteChange = useCallback(
    (value: string) => {
      setLocalUtenteIdentifier(value);
      emitFilterChange({
        docenteNome: localDocenteNome.trim() || undefined,
        docenteCognome: localDocenteCognome.trim() || undefined,
        copieEffettuate: localCopieEffettuate.trim() || undefined,
        utenteIdentifier: value.trim() || undefined,
      });
    },
    [localDocenteNome, localDocenteCognome, localCopieEffettuate, emitFilterChange]
  );

  return (
    <div className="w-full">
      <div className="hidden lg:block ">
        {type === "docenti" && (
          <FieldGroup className="flex flex-row gap-2 max-w-xl">
            <Field className="flex-1">
              <FieldLabel htmlFor="nome-input">Cerca docente per nome...</FieldLabel>
              <InputGroup>
                <InputGroupInput
                  id="nome-input"
                  value={localNome}
                  onChange={(e) => handleDocentiNomeChange(e.target.value)}
                  placeholder="Marco..."
                />
                <InputGroupAddon align="inline-start">
                  <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                </InputGroupAddon>
              </InputGroup>
            </Field>
            <Field className="flex-1">
              <FieldLabel htmlFor="cognome-input">Cerca docente per cognome...</FieldLabel>
              <InputGroup>
                <InputGroupInput
                  id="cognome-input"
                  value={localCognome}
                  onChange={(e) => handleDocentiCognomeChange(e.target.value)}
                  placeholder="Rossi..."
                />
                <InputGroupAddon align="inline-start">
                  <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                </InputGroupAddon>
              </InputGroup>
            </Field>
          </FieldGroup>
        )}
        {type === "utenze" && (
          <div className="flex flex-row gap-2 max-w-xl">
            <Field className="">
              <FieldLabel htmlFor="search-utenti">Cerca per email o username</FieldLabel>
              <InputGroup>
                <InputGroupInput
                  id="search-utenti"
                  value={localIdentifier}
                  onChange={(e) => handleUtenzeIdentifierChange(e.target.value)}
                  placeholder="email@esempio.it o username..."
                />
                <InputGroupAddon align="inline-start">
                  <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                </InputGroupAddon>
              </InputGroup>
            </Field>
            <Field>
              <FieldLabel htmlFor="ruolo-select">
                Ruolo
              </FieldLabel>
              <Select defaultValue={localRuolo} onValueChange={(value) => handleUtenzeRuoloChange(value)}>
                <SelectTrigger id="ruolo-select">
                  <SelectValue placeholder="Ruolo" />
                </SelectTrigger>
                <SelectContent>
                  <SelectGroup>
                    <SelectItem value="all">Tutti</SelectItem>
                    <SelectItem value="admin">Amministratore</SelectItem>
                    <SelectItem value="collaboratore">Collaboratore</SelectItem>
                  </SelectGroup>
                </SelectContent>
              </Select>
            </Field>
          </div>
        )}
        {type === "registrazioni" && (
          <FieldGroup className="flex flex-row  gap-4 max-w-7xl">
            <Field className="">
              <FieldLabel htmlFor="reg-docente-nome">Nome docente</FieldLabel>
              <InputGroup>
                <InputGroupInput
                  id="reg-docente-nome"
                  value={localDocenteNome}
                  onChange={(e) => handleRegistrazioniDocenteNomeChange(e.target.value)}
                  placeholder="Nome docente…"
                />
                <InputGroupAddon align="inline-start">
                  <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                </InputGroupAddon>
              </InputGroup>
            </Field>
            <Field className="">
              <FieldLabel htmlFor="reg-docente-cognome">Cognome docente</FieldLabel>
              <InputGroup>
                <InputGroupInput
                  id="reg-docente-cognome"
                  value={localDocenteCognome}
                  onChange={(e) => handleRegistrazioniDocenteCognomeChange(e.target.value)}
                  placeholder="Cognome docente…"
                />
                <InputGroupAddon align="inline-start">
                  <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                </InputGroupAddon>
              </InputGroup>
            </Field>
            <Field className="max-w-36">
              <FieldLabel htmlFor="reg-copie">Copie effettuate</FieldLabel>
              <InputGroup>
                <InputGroupInput
                  id="reg-copie"
                  type="number"
                  min={0}
                  value={localCopieEffettuate}
                  onChange={(e) => handleRegistrazioniCopieChange(e.target.value)}
                  placeholder="Copie"
                />
              </InputGroup>
            </Field>
            <Field className="">
              <FieldLabel htmlFor="reg-utente">Utente</FieldLabel>
              <InputGroup>
                <InputGroupInput
                  id="reg-utente"
                  value={localUtenteIdentifier}
                  onChange={(e) => handleRegistrazioniUtenteChange(e.target.value)}
                  placeholder="Username o email…"
                />
                <InputGroupAddon align="inline-start">
                  <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                </InputGroupAddon>
              </InputGroup>
            </Field>
          </FieldGroup>
        )}

      </div>
      <Popover>
        <PopoverTrigger asChild className="lg:hidden">
          <Button variant="outline">
            <HugeiconsIcon icon={FilterIcon} strokeWidth={2} className="size-4" />
          </Button>
        </PopoverTrigger>
        <PopoverContent className="flex flex-col gap-4" align="start">
          {type === "docenti" && (
            <FieldGroup className="flex flex-col gap-2">
              <Field className="flex flex-col gap-2">
                <FieldLabel htmlFor="nome-input" className="text-xs">Cerca docente per nome...</FieldLabel>
                <InputGroup>
                  <InputGroupInput
                    id="nome-input"
                    value={localNome}
                    onChange={(e) => handleDocentiNomeChange(e.target.value)}
                    placeholder="Marco..."
                  />
                  <InputGroupAddon align="inline-start">
                    <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                  </InputGroupAddon>
                </InputGroup>
              </Field>
              <Field className="flex flex-col gap-2">
                <FieldLabel htmlFor="cognome-input" className="text-xs">Cerca docente per cognome...</FieldLabel>
                <InputGroup>
                  <InputGroupInput
                    id="cognome-input"
                    value={localCognome}
                    onChange={(e) => handleDocentiCognomeChange(e.target.value)}
                    placeholder="Rossi..."
                  />
                  <InputGroupAddon align="inline-start">
                    <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                  </InputGroupAddon>
                </InputGroup>
              </Field>
            </FieldGroup>
          )}
          {type === "utenze" && (
            <div className="flex flex-col gap-2">
              <Field className="flex flex-col gap-2">
                <FieldLabel htmlFor="search-utenti" className="text-xs">Cerca per email o username</FieldLabel>
                <InputGroup>
                  <InputGroupInput
                    
                    id="search-utenti"
                    value={localIdentifier}
                    onChange={(e) => handleUtenzeIdentifierChange(e.target.value)}
                    placeholder="email@esempio.it o username..."
                  />
                  <InputGroupAddon align="inline-start">
                    <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                  </InputGroupAddon>
                </InputGroup>
              </Field>
              <Field className="flex flex-col gap-2">
                <FieldLabel htmlFor="ruolo-select" className="text-xs">
                  Ruolo
                </FieldLabel>
                <Select defaultValue={localRuolo} onValueChange={(value) => handleUtenzeRuoloChange(value)}>
                  <SelectTrigger id="ruolo-select" className="text-xs">
                    <SelectValue placeholder="Ruolo" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectGroup>
                      <SelectItem value="all">Tutti</SelectItem>
                      <SelectItem value="admin">Amministratore</SelectItem>
                      <SelectItem value="collaboratore">Collaboratore</SelectItem>
                    </SelectGroup>
                  </SelectContent>
                </Select>
              </Field>
            </div>
          )}
          {type === "registrazioni" && (
            <FieldGroup className="grid grid-cols-2 gap-2" >
              <Field className="flex flex-col gap-2">
                <FieldLabel htmlFor="reg-docente-nome" className="text-xs">Nome docente</FieldLabel>
                <InputGroup>
                  <InputGroupInput
                    id="reg-docente-nome"
                    value={localDocenteNome}
                    className="text-xs"
                    onChange={(e) => handleRegistrazioniDocenteNomeChange(e.target.value)}
                    placeholder="Nome..."
                  />
                  <InputGroupAddon align="inline-start">
                    <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                  </InputGroupAddon>
                </InputGroup>
              </Field>
              <Field className="flex flex-col gap-2">
                <FieldLabel htmlFor="reg-docente-cognome" className="text-xs">Cognome docente</FieldLabel>
                <InputGroup>
                  <InputGroupInput
                    id="reg-docente-cognome"
                    value={localDocenteCognome}
                    className="text-xs"
                    onChange={(e) => handleRegistrazioniDocenteCognomeChange(e.target.value)}
                    placeholder="Cognome..."
                  />
                  <InputGroupAddon align="inline-start">
                    <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                  </InputGroupAddon>
                </InputGroup>
              </Field>
              <Field className="flex flex-col gap-2 max-w-36">
                <FieldLabel htmlFor="reg-copie" className="text-xs">Copie effettuate</FieldLabel>
                <InputGroup>
                  <InputGroupInput
                    id="reg-copie"
                    type="number"
                    min={0}
                    value={localCopieEffettuate}
                    className="text-xs"
                    onChange={(e) => handleRegistrazioniCopieChange(e.target.value)}
                    placeholder="Copie..."
                  />
                </InputGroup>
              </Field>
              <Field className="flex flex-col gap-2 col-span-2">
                <FieldLabel htmlFor="reg-utente" className="text-xs">Utente</FieldLabel>
                <InputGroup>
                  <InputGroupInput
                    id="reg-utente"
                    value={localUtenteIdentifier}
                    className="text-xs"
                    onChange={(e) => handleRegistrazioniUtenteChange(e.target.value)}
                    placeholder="Username o email…"
                  />
                  <InputGroupAddon align="inline-start">
                    <HugeiconsIcon icon={SearchIcon} strokeWidth={2} className="size-4" />
                  </InputGroupAddon>
                </InputGroup>
              </Field>
            </FieldGroup>
          )}
        </PopoverContent>
      </Popover>
    </div>
  );
}
