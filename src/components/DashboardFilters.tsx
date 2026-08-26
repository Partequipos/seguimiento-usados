/**
 * Componente de filtros para el Dashboard
 * Implementa filtros en cascada (indexados) donde cada dropdown solo muestra
 * opciones relacionadas con los datos ya filtrados por otros filtros
 */

import React, { useMemo } from "react";
import { Filter, X } from "lucide-react";
import { SharePointListItem } from "../services/sharePointService";
import {
  AVANCE_FILTER_STEPS,
  getAvanceStepFilterKey,
  getFieldValue,
  matchesPorcentajeAvanceFilter,
  parsePorcentajeAvance,
  toDateOnlyString,
} from "../utils/sharePointFieldMapping";

interface DashboardFiltersProps {
  items: SharePointListItem[];
  filters: FilterState;
  onFilterChange: (filters: FilterState) => void;
}

export interface FilterState {
  sede: string;
  asesor: string;
  cliente: string;
  serie: string;
  fase: string;
  observaciones: string;
  ciclo: string[];
  fechaCompromisoDesde: string;
  fechaCompromisoHasta: string;
  fechaFinalDesde: string;
  fechaFinalHasta: string;
  /** "" | "100" | ">0" | "0" | "15" | "30" | … | "99" */
  porcentajeAvance: string;
}

const sortStrings = (a: string, b: string): number => a.localeCompare(b);

// Predicados por filtro (reducen complejidad cognitiva del filter principal)
function matchesSede(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  if (excludeField === "sede" || !filters.sede) return true;
  return getFieldValue(item.fields, "Sede") === filters.sede;
}

function matchesAsesor(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  if (excludeField === "asesor" || !filters.asesor) return true;
  return getFieldValue(item.fields, "Asesor") === filters.asesor;
}

function matchesCliente(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  if (excludeField === "cliente" || !filters.cliente) return true;
  return getFieldValue(item.fields, "Title") === filters.cliente;
}

function matchesSerie(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  if (excludeField === "serie" || !filters.serie) return true;
  return getFieldValue(item.fields, "Serie") === filters.serie;
}

function matchesObservaciones(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  if (excludeField === "observaciones" || !filters.observaciones) return true;
  return getFieldValue(item.fields, "Observaciones") === filters.observaciones;
}

function matchesCiclo(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  if (excludeField === "ciclo" || filters.ciclo.length === 0) return true;
  const raw = getFieldValue(item.fields, "Ciclo");
  let itemCicloStr = "";
  if (typeof raw === "string") itemCicloStr = raw;
  else if (raw != null && typeof raw === "number") itemCicloStr = String(raw);
  return filters.ciclo.includes(itemCicloStr);
}

function matchesFechaCompromiso(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  const skip =
    excludeField === "fechaCompromisoDesde" ||
    excludeField === "fechaCompromisoHasta";
  if (skip || (!filters.fechaCompromisoDesde && !filters.fechaCompromisoHasta))
    return true;
  const value = getFieldValue(item.fields, "FechaCompromisoComercial");
  if (!value) return false;
  const date = new Date(value as string | number | Date);
  if (Number.isNaN(date.getTime())) return false;
  const itemStr = toDateOnlyString(date);
  // Los inputs type="date" envían YYYY-MM-DD; usarlos directos evita desfase por zona horaria
  if (filters.fechaCompromisoDesde && itemStr < filters.fechaCompromisoDesde)
    return false;
  if (filters.fechaCompromisoHasta && itemStr > filters.fechaCompromisoHasta)
    return false;
  return true;
}

function matchesFechaFinalAlistamiento(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  const skip =
    excludeField === "fechaFinalDesde" || excludeField === "fechaFinalHasta";
  if (skip || (!filters.fechaFinalDesde && !filters.fechaFinalHasta))
    return true;
  const value = getFieldValue(item.fields, "FechaFinalAlistamiento");
  if (!value) return false;
  const date = new Date(value as string | number | Date);
  if (Number.isNaN(date.getTime())) return false;
  const itemStr = toDateOnlyString(date);
  // Los inputs type="date" envían YYYY-MM-DD; usarlos directos evita desfase por zona horaria
  if (filters.fechaFinalDesde && itemStr < filters.fechaFinalDesde)
    return false;
  if (filters.fechaFinalHasta && itemStr > filters.fechaFinalHasta)
    return false;
  return true;
}

function matchesPorcentajeAvance(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  if (excludeField === "porcentajeAvance" || !filters.porcentajeAvance)
    return true;
  const avance = parsePorcentajeAvance(item.fields as Record<string, unknown>);
  return matchesPorcentajeAvanceFilter(avance, filters.porcentajeAvance);
}

function formatAvanceStepLabel(step: number): string {
  const stepIndex = AVANCE_FILTER_STEPS.indexOf(
    step as (typeof AVANCE_FILTER_STEPS)[number]
  );
  if (stepIndex === -1) return `${step}%`;
  const upperExclusive =
    stepIndex < AVANCE_FILTER_STEPS.length - 1
      ? AVANCE_FILTER_STEPS[stepIndex + 1]
      : 100;
  if (step === 99) return "99% – <100%";
  return `${step}% – ${upperExclusive - 1}%`;
}

function itemMatchesAllFilters(
  item: SharePointListItem,
  filters: FilterState,
  excludeField: keyof FilterState
): boolean {
  return (
    matchesSede(item, filters, excludeField) &&
    matchesAsesor(item, filters, excludeField) &&
    matchesCliente(item, filters, excludeField) &&
    matchesSerie(item, filters, excludeField) &&
    matchesObservaciones(item, filters, excludeField) &&
    matchesCiclo(item, filters, excludeField) &&
    matchesFechaCompromiso(item, filters, excludeField) &&
    matchesFechaFinalAlistamiento(item, filters, excludeField) &&
    matchesPorcentajeAvance(item, filters, excludeField)
  );
}

const applyFiltersExcept = (
  items: SharePointListItem[],
  filters: FilterState,
  excludeField: keyof FilterState
): SharePointListItem[] =>
  items.filter((item) => itemMatchesAllFilters(item, filters, excludeField));

const DashboardFilters: React.FC<DashboardFiltersProps> = ({
  items,
  filters,
  onFilterChange,
}) => {
  // Filtros en cascada: cada dropdown muestra solo opciones de datos ya filtrados por otros filtros

  // Sede - no depende de otros filtros categóricos
  const sedes = useMemo(() => {
    const filtered = applyFiltersExcept(items, filters, "sede");
    const values = [
      ...new Set(
        filtered
          .map((i) => {
            const sede = i.fields.field_28 || i.fields.Sede;
            return sede;
          })
          .filter(Boolean)
      ),
    ];
    return values.map(String).sort(sortStrings);
  }, [items, filters]);

  // Asesor - depende de sede, ciclo, observaciones, porcentajeAvance, fechas
  const asesores = useMemo(() => {
    const filtered = applyFiltersExcept(items, filters, "asesor");
    const values = [
      ...new Set(
        filtered
          .map((i) => {
            const asesor = i.fields.field_4 || i.fields.Asesor;
            return asesor;
          })
          .filter(Boolean)
      ),
    ];
    return values.map(String).sort(sortStrings);
  }, [items, filters]);

  // Cliente - depende de todos los demás filtros
  const clientes = useMemo(() => {
    const filtered = applyFiltersExcept(items, filters, "cliente");
    const values = [
      ...new Set(filtered.map((i) => i.fields.Title).filter(Boolean)),
    ];
    return values.map(String).sort(sortStrings);
  }, [items, filters]);

  // Serie - depende de todos los demás filtros
  const series = useMemo(() => {
    const filtered = applyFiltersExcept(items, filters, "serie");
    const values = [
      ...new Set(
        filtered
          .map((i) => {
            const serie = i.fields.field_0 || i.fields.Serie;
            return serie;
          })
          .filter(Boolean)
      ),
    ];
    return values.map(String).sort(sortStrings);
  }, [items, filters]);

  // Observaciones - depende de otros filtros
  const observaciones = useMemo(() => {
    const filtered = applyFiltersExcept(items, filters, "observaciones");
    const values = [
      ...new Set(
        filtered
          .map((i) => {
            const obs = i.fields.field_8 || i.fields.Observaciones;
            return obs;
          })
          .filter(Boolean)
      ),
    ];
    return values.map(String).sort(sortStrings);
  }, [items, filters]);

  // Ciclo - depende de otros filtros
  const ciclos = useMemo(() => {
    const filtered = applyFiltersExcept(items, filters, "ciclo");
    const values = [
      ...new Set(
        filtered
          .map((i) => {
            const ciclo = i.fields.field_29 || i.fields.Ciclo;
            return ciclo;
          })
          .filter(Boolean)
      ),
    ];
    return values.map(String).sort(sortStrings);
  }, [items, filters]);

  // % Avance - índice para cascada (afecta otros filtros); opciones fijas siempre visibles
  const porcentajeAvanceOptions = useMemo(() => {
    const filtered = applyFiltersExcept(items, filters, "porcentajeAvance");
    const available = new Set<string>();

    for (const item of filtered) {
      const avance = parsePorcentajeAvance(
        item.fields as Record<string, unknown>
      );
      if (avance === 0) {
        available.add("0");
        continue;
      }
      if (avance === 100) {
        available.add("100");
        continue;
      }
      available.add(">0");
      const stepKey = getAvanceStepFilterKey(avance);
      if (stepKey) available.add(stepKey);
    }

    return available;
  }, [items, filters]);

  const handleChange = (field: keyof FilterState, value: string) => {
    onFilterChange({
      ...filters,
      [field]: value,
    });
  };

  const handleCicloToggle = (cicloValue: string) => {
    const current = filters.ciclo;
    const next = current.includes(cicloValue)
      ? current.filter((c) => c !== cicloValue)
      : [...current, cicloValue];
    onFilterChange({ ...filters, ciclo: next });
  };

  const handleCicloSelectAll = () => {
    onFilterChange({ ...filters, ciclo: [...ciclos] });
  };

  const handleCicloClear = () => {
    onFilterChange({ ...filters, ciclo: [] });
  };

  const clearFilters = () => {
    onFilterChange({
      sede: "",
      asesor: "",
      cliente: "",
      serie: "",
      fase: "",
      observaciones: "",
      ciclo: [],
      fechaCompromisoDesde: "",
      fechaCompromisoHasta: "",
      fechaFinalDesde: "",
      fechaFinalHasta: "",
      porcentajeAvance: "",
    });
  };

  const hasActiveFilters = ((): boolean => {
    const entries = Object.entries(filters) as [
      keyof FilterState,
      FilterState[keyof FilterState],
    ][];
    return entries.some(([key, v]) =>
      key === "ciclo"
        ? Array.isArray(v) && v.length > 0
        : (v as string) !== ""
    );
  })();

  return (
    <div className="panel-brand p-6 mb-6 w-full border-l-4 border-l-brand-red">
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <Filter className="w-5 h-5 text-brand-red" />
          <h3 className="text-lg font-semibold text-gray-900">Filtros</h3>
          {hasActiveFilters && (
            <span className="ml-2 px-2 py-1 bg-brand-red-soft text-brand-red-dark text-xs font-medium rounded">
              Filtros activos
            </span>
          )}
        </div>
        {hasActiveFilters && (
          <button
            onClick={clearFilters}
            className="flex items-center gap-1 text-sm text-gray-600 hover:text-gray-900"
          >
            <X className="w-4 h-4" />
            Limpiar filtros
          </button>
        )}
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-4 w-full">
        {/* Sede */}
        <div>
          <label
            htmlFor="filter-sede"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Sede
          </label>
          <select
            id="filter-sede"
            value={filters.sede}
            onChange={(e) => handleChange("sede", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          >
            <option value="">Todas las sedes</option>
            {sedes.map((sede) => (
              <option key={sede} value={sede}>
                {sede}
              </option>
            ))}
          </select>
        </div>

        {/* Asesor */}
        <div>
          <label
            htmlFor="filter-asesor"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Asesor
          </label>
          <select
            id="filter-asesor"
            value={filters.asesor}
            onChange={(e) => handleChange("asesor", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          >
            <option value="">Todos los asesores</option>
            {asesores.map((asesor) => (
              <option key={asesor} value={asesor}>
                {asesor}
              </option>
            ))}
          </select>
        </div>

        {/* Cliente */}
        <div>
          <label
            htmlFor="filter-cliente"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Cliente
          </label>
          <select
            id="filter-cliente"
            value={filters.cliente}
            onChange={(e) => handleChange("cliente", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          >
            <option value="">Todos los clientes</option>
            {clientes.map((cliente) => (
              <option key={cliente} value={cliente}>
                {cliente}
              </option>
            ))}
          </select>
        </div>

        {/* Serie */}
        <div>
          <label
            htmlFor="filter-serie"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Serie
          </label>
          <select
            id="filter-serie"
            value={filters.serie}
            onChange={(e) => handleChange("serie", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          >
            <option value="">Todas las series</option>
            {series.map((serie) => (
              <option key={serie} value={serie}>
                {serie}
              </option>
            ))}
          </select>
        </div>

        {/* Fase */}
        <div>
          <label
            htmlFor="filter-fase"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Fase (F1-F16)
          </label>
          <select
            id="filter-fase"
            value={filters.fase}
            onChange={(e) => handleChange("fase", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          >
            <option value="">Todas las fases</option>
            {Array.from({ length: 16 }, (_, i) => i + 1).map((num) => (
              <option key={`F${num}`} value={`F${num}`}>
                Fase {num} (F{num})
              </option>
            ))}
          </select>
        </div>

        {/* Observaciones */}
        <div>
          <label
            htmlFor="filter-observaciones"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Observaciones
          </label>
          <select
            id="filter-observaciones"
            value={filters.observaciones}
            onChange={(e) => handleChange("observaciones", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          >
            <option value="">Todas</option>
            {observaciones.map((obs) => (
              <option key={obs} value={obs}>
                {obs}
              </option>
            ))}
          </select>
        </div>

        {/* Ciclo - selección múltiple (lista contraíble) */}
        <details className="group border border-gray-300 rounded-lg bg-white focus-within:ring-2 focus-within:ring-brand-red/40 focus-within:border-transparent">
          <summary className="text-sm font-medium text-gray-700 px-3 py-2 cursor-pointer list-none flex items-center justify-between gap-2 hover:bg-gray-50 rounded-lg [&::-webkit-details-marker]:hidden">
            <span>
              Ciclo
              {filters.ciclo.length > 0 && (
                <span className="ml-1 text-gray-500 font-normal">
                  ({filters.ciclo.length} seleccionados)
                </span>
              )}
            </span>
            <span className="text-gray-400 group-open:rotate-180 transition-transform shrink-0" aria-hidden>
              ▼
            </span>
          </summary>
          <div className="px-3 pb-3 pt-1 border-t border-gray-100 max-h-48 overflow-y-auto">
            <div className="flex flex-wrap gap-x-3 gap-y-1 mb-2 pt-2 border-b border-gray-100 pb-2">
              <button
                type="button"
                onClick={handleCicloSelectAll}
                className="text-xs text-brand-red hover:text-brand-red-dark font-medium"
              >
                Todos
              </button>
              <button
                type="button"
                onClick={handleCicloClear}
                className="text-xs text-gray-500 hover:text-gray-700 font-medium"
              >
                Ninguno
              </button>
            </div>
            <div className="flex flex-col gap-1">
              {ciclos.map((ciclo) => (
                <label
                  key={ciclo}
                  className="flex items-center gap-2 cursor-pointer hover:bg-gray-50 rounded px-1 py-0.5 text-sm"
                >
                  <input
                    type="checkbox"
                    checked={filters.ciclo.includes(ciclo)}
                    onChange={() => handleCicloToggle(ciclo)}
                    className="rounded border-gray-300 text-brand-red focus:ring-brand-red/40"
                  />
                  <span>{ciclo}</span>
                </label>
              ))}
            </div>
            {ciclos.length === 0 && (
              <p className="text-sm text-gray-500 py-1">Sin ciclos disponibles</p>
            )}
          </div>
        </details>

        {/* % de Avance */}
        <div>
          <label
            htmlFor="filter-porcentajeAvance"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            % de Avance
          </label>
          <select
            id="filter-porcentajeAvance"
            value={filters.porcentajeAvance}
            onChange={(e) => handleChange("porcentajeAvance", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          >
            <option value="">Todos</option>
            <option value="100">100% (Completados)</option>
            <option value=">0">{">0% y <100%"} (En Proceso)</option>
            {AVANCE_FILTER_STEPS.map((step) => (
              <option key={step} value={String(step)}>
                {formatAvanceStepLabel(step)}
                {!porcentajeAvanceOptions.has(String(step))
                  ? " — sin datos"
                  : ""}
              </option>
            ))}
            <option value="0">0% (Pendientes)</option>
          </select>
        </div>

        {/* Fecha Compromiso Desde */}
        <div>
          <label
            htmlFor="filter-fechaCompromisoDesde"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Fecha Compromiso (Desde)
          </label>
          <input
            id="filter-fechaCompromisoDesde"
            type="date"
            value={filters.fechaCompromisoDesde}
            onChange={(e) =>
              handleChange("fechaCompromisoDesde", e.target.value)
            }
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          />
        </div>

        {/* Fecha Compromiso Hasta */}
        <div>
          <label
            htmlFor="filter-fechaCompromisoHasta"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Fecha Compromiso (Hasta)
          </label>
          <input
            id="filter-fechaCompromisoHasta"
            type="date"
            value={filters.fechaCompromisoHasta}
            onChange={(e) =>
              handleChange("fechaCompromisoHasta", e.target.value)
            }
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          />
        </div>

        {/* Fecha Final Alistamiento Desde */}
        <div>
          <label
            htmlFor="filter-fechaFinalDesde"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Fecha Final Alistamiento (Desde)
          </label>
          <input
            id="filter-fechaFinalDesde"
            type="date"
            value={filters.fechaFinalDesde}
            onChange={(e) => handleChange("fechaFinalDesde", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          />
        </div>

        {/* Fecha Final Alistamiento Hasta */}
        <div>
          <label
            htmlFor="filter-fechaFinalHasta"
            className="block text-sm font-medium text-gray-700 mb-1"
          >
            Fecha Final Alistamiento (Hasta)
          </label>
          <input
            id="filter-fechaFinalHasta"
            type="date"
            value={filters.fechaFinalHasta}
            onChange={(e) => handleChange("fechaFinalHasta", e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
          />
        </div>
      </div>
    </div>
  );
};

export default DashboardFilters;
