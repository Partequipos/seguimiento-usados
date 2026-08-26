/**
 * Visualización Ciclo-Mes — tablero para proyección en TV
 * Tabla ordenable (prioridad / % avance / manual) con semáforo y KPIs.
 */

import React, { useCallback, useEffect, useMemo, useState } from "react";
import {
  PieChart,
  Pie,
  Cell,
  Tooltip,
  ResponsiveContainer,
  Legend,
} from "recharts";
import {
  Target,
  Award,
  CheckCircle2,
  Filter,
  GripVertical,
  ArrowUpDown,
} from "lucide-react";
import { SharePointListItem } from "../services/sharePointService";
import {
  getFieldValue,
  parsePorcentajeAvance,
} from "../utils/sharePointFieldMapping";

const META_MENSUAL = 12;
const COMISION_EXTRA_MIN = 10;

type SortMode = "prioridad" | "avance" | "manual";

interface CicloMesViewProps {
  items: SharePointListItem[];
}

function fieldStr(fields: Record<string, unknown>, key: string): string {
  const value = getFieldValue(fields, key);
  if (value == null) return "";
  if (typeof value === "string" || typeof value === "number") {
    return String(value);
  }
  return "";
}

function normalizeCicloValue(raw: unknown): string {
  if (raw == null || raw === "") return "";
  if (typeof raw === "number") return `Ciclo ${raw}`;
  if (typeof raw !== "string") return "";
  const str = raw.trim();
  if (/^Ciclo\s+/i.test(str)) return str;
  return `Ciclo ${str}`;
}

function getPrioridad(fields: Record<string, unknown>): number {
  return Number(getFieldValue(fields, "Prioridad")) || 0;
}

function getAvanceRowClass(avance: number): string {
  if (avance >= 90) return "bg-green-200 border-l-8 border-green-600";
  if (avance >= 50) return "bg-yellow-200 border-l-8 border-yellow-500";
  return "bg-red-200 border-l-8 border-red-600";
}

function getAvanceTextClass(avance: number): string {
  if (avance >= 90) return "text-green-900";
  if (avance >= 50) return "text-yellow-900";
  return "text-red-900";
}

function comparePrioridad(
  a: SharePointListItem,
  b: SharePointListItem
): number {
  const pa = getPrioridad(a.fields);
  const pb = getPrioridad(b.fields);
  const aInRange = pa >= 1 && pa <= 10;
  const bInRange = pb >= 1 && pb <= 10;

  if (aInRange && bInRange) return pa - pb;
  if (aInRange) return -1;
  if (bInRange) return 1;
  if (pa !== pb) return pa - pb;
  return fieldStr(a.fields, "Serie").localeCompare(fieldStr(b.fields, "Serie"));
}

function compareAvance(a: SharePointListItem, b: SharePointListItem): number {
  const avanceA = parsePorcentajeAvance(a.fields);
  const avanceB = parsePorcentajeAvance(b.fields);
  if (avanceA !== avanceB) return avanceA - avanceB;
  return comparePrioridad(a, b);
}

/** Las de 100% siempre van al final; el resto según el comparador. */
function sortWithCompletedLast(
  list: SharePointListItem[],
  compareActive: (a: SharePointListItem, b: SharePointListItem) => number
): SharePointListItem[] {
  const active = list.filter(
    (item) => parsePorcentajeAvance(item.fields) < 100
  );
  const completed = list.filter(
    (item) => parsePorcentajeAvance(item.fields) === 100
  );
  active.sort(compareActive);
  completed.sort(comparePrioridad);
  return [...active, ...completed];
}

function applyManualOrder(
  list: SharePointListItem[],
  orderIds: string[]
): SharePointListItem[] {
  const byId = new Map(list.map((item) => [item.id, item]));
  const ordered: SharePointListItem[] = [];

  for (const id of orderIds) {
    const item = byId.get(id);
    if (item) {
      ordered.push(item);
      byId.delete(id);
    }
  }
  for (const item of byId.values()) {
    ordered.push(item);
  }

  const active = ordered.filter(
    (item) => parsePorcentajeAvance(item.fields) < 100
  );
  const completed = ordered.filter(
    (item) => parsePorcentajeAvance(item.fields) === 100
  );
  return [...active, ...completed];
}

function reorderIds(
  ids: string[],
  fromId: string,
  toId: string
): string[] | null {
  if (fromId === toId) return null;
  const next = [...ids];
  const fromIndex = next.indexOf(fromId);
  const toIndex = next.indexOf(toId);
  if (fromIndex < 0 || toIndex < 0) return null;
  next.splice(fromIndex, 1);
  next.splice(toIndex, 0, fromId);
  return next;
}

const CicloMesView: React.FC<CicloMesViewProps> = ({ items }) => {
  const ciclosDisponibles = useMemo(() => {
    const set = new Set<string>();
    for (const item of items) {
      const ciclo = normalizeCicloValue(getFieldValue(item.fields, "Ciclo"));
      if (ciclo) set.add(ciclo);
    }
    return [...set].sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true })
    );
  }, [items]);

  const [selectedCiclo, setSelectedCiclo] = useState<string>("");
  const [sortMode, setSortMode] = useState<SortMode>("prioridad");
  const [manualOrderIds, setManualOrderIds] = useState<string[]>([]);
  const [draggedId, setDraggedId] = useState<string | null>(null);

  const cicloActivo = selectedCiclo || ciclosDisponibles[0] || "";

  const itemsDelCiclo = useMemo(() => {
    if (!cicloActivo) return [];
    return items.filter(
      (item) =>
        normalizeCicloValue(getFieldValue(item.fields, "Ciclo")) === cicloActivo
    );
  }, [items, cicloActivo]);

  // Al cambiar de ciclo, volver a orden por prioridad
  useEffect(() => {
    setSortMode("prioridad");
    setManualOrderIds([]);
    setDraggedId(null);
  }, [cicloActivo]);

  // Mantener el orden manual sincronizado con los ítems del ciclo
  useEffect(() => {
    if (sortMode !== "manual") return;
    setManualOrderIds((prev) => {
      const currentIds = new Set(itemsDelCiclo.map((i) => i.id));
      const kept = prev.filter((id) => currentIds.has(id));
      const missing = itemsDelCiclo
        .map((i) => i.id)
        .filter((id) => !kept.includes(id));
      if (kept.length === prev.length && missing.length === 0) return prev;
      return [...kept, ...missing];
    });
  }, [itemsDelCiclo, sortMode]);

  const displayItems = useMemo(() => {
    if (sortMode === "manual") {
      const baseOrder =
        manualOrderIds.length > 0
          ? manualOrderIds
          : itemsDelCiclo.map((i) => i.id);
      return applyManualOrder(itemsDelCiclo, baseOrder);
    }
    if (sortMode === "avance") {
      return sortWithCompletedLast(itemsDelCiclo, compareAvance);
    }
    return sortWithCompletedLast(itemsDelCiclo, comparePrioridad);
  }, [itemsDelCiclo, sortMode, manualOrderIds]);

  const handleSortModeChange = (mode: SortMode) => {
    if (mode === "manual") {
      const currentIds = displayItems.map((i) => i.id);
      setManualOrderIds(currentIds);
    }
    setSortMode(mode);
  };

  const handleDragStart = (itemId: string, avance: number) => {
    if (avance === 100) return;
    setDraggedId(itemId);
    if (sortMode !== "manual") {
      setManualOrderIds(displayItems.map((i) => i.id));
      setSortMode("manual");
    }
  };

  const handleDragOver = useCallback(
    (event: React.DragEvent<HTMLTableRowElement>, targetId: string) => {
      event.preventDefault();
      if (!draggedId || draggedId === targetId) return;
      const targetItem = itemsDelCiclo.find((i) => i.id === targetId);
      if (
        targetItem &&
        parsePorcentajeAvance(targetItem.fields) === 100
      ) {
        return;
      }
    },
    [draggedId, itemsDelCiclo]
  );

  const handleDrop = (targetId: string) => {
    if (!draggedId) return;
    const targetItem = itemsDelCiclo.find((i) => i.id === targetId);
    const draggedItem = itemsDelCiclo.find((i) => i.id === draggedId);

    if (
      !targetItem ||
      !draggedItem ||
      parsePorcentajeAvance(targetItem.fields) === 100 ||
      parsePorcentajeAvance(draggedItem.fields) === 100
    ) {
      setDraggedId(null);
      return;
    }

    const baseIds =
      manualOrderIds.length > 0
        ? manualOrderIds
        : displayItems.map((i) => i.id);
    const reordered = reorderIds(baseIds, draggedId, targetId);
    if (reordered) {
      setManualOrderIds(reordered);
      setSortMode("manual");
    }
    setDraggedId(null);
  };

  const alistadasCount = useMemo(
    () =>
      itemsDelCiclo.filter(
        (item) => parsePorcentajeAvance(item.fields) === 100
      ).length,
    [itemsDelCiclo]
  );

  const metaCumplida = alistadasCount >= META_MENSUAL;
  const comisionExtraCumplida = alistadasCount >= COMISION_EXTRA_MIN;

  const kpiPieData = useMemo(
    () => [
      {
        name: `Alistadas (${alistadasCount})`,
        value: Math.max(alistadasCount, 1),
        color: "#16a34a",
      },
      {
        name: `Meta (${META_MENSUAL})`,
        value: META_MENSUAL,
        color: metaCumplida ? "#16a34a" : "#ef4444",
      },
      {
        name: `Comisión extra (≥${COMISION_EXTRA_MIN})`,
        value: COMISION_EXTRA_MIN,
        color: comisionExtraCumplida ? "#16a34a" : "#ef4444",
      },
    ],
    [alistadasCount, metaCumplida, comisionExtraCumplida]
  );

  const sortButtonClass = (mode: SortMode) =>
    `px-4 py-2 rounded-lg text-sm font-semibold border transition-colors ${
      sortMode === mode
        ? "bg-brand-red text-white border-brand-red"
        : "bg-white text-gray-700 border-gray-300 hover:bg-gray-50"
    }`;

  return (
    <div className="w-full space-y-6">
      {/* Selector de ciclo */}
      <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-6 flex flex-wrap items-end gap-4">
        <div className="flex items-center gap-2 text-gray-800">
          <Filter className="w-6 h-6 text-brand-red" />
          <h2 className="text-2xl font-bold">Visualización Ciclo-Mes</h2>
        </div>
        <div className="flex-1 min-w-[220px] max-w-md">
          <label
            htmlFor="ciclo-mes-selector"
            className="block text-sm font-semibold text-gray-700 mb-1"
          >
            Ciclo
          </label>
          <select
            id="ciclo-mes-selector"
            value={cicloActivo}
            onChange={(e) => setSelectedCiclo(e.target.value)}
            className="w-full px-4 py-3 text-lg border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent bg-white"
          >
            {ciclosDisponibles.length === 0 && (
              <option value="">Sin ciclos disponibles</option>
            )}
            {ciclosDisponibles.map((ciclo) => (
              <option key={ciclo} value={ciclo}>
                {ciclo}
              </option>
            ))}
          </select>
        </div>
        <p className="text-lg text-gray-600">
          <span className="font-bold text-gray-900">{displayItems.length}</span>{" "}
          equipos en {cicloActivo || "—"}
        </p>
      </div>

      {/* KPIs */}
      <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
        <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-6">
          <h3 className="text-xl font-bold text-gray-900 mb-4 flex items-center gap-2">
            <CheckCircle2 className="w-6 h-6 text-green-600" />
            Alistadas (100%)
          </h3>
          <p className="text-5xl font-black text-gray-900">{alistadasCount}</p>
          <p className="text-sm text-gray-500 mt-2">
            Equipos completados en el ciclo seleccionado
          </p>
        </div>

        <div
          className={`rounded-xl shadow-lg border p-6 ${
            metaCumplida
              ? "bg-green-50 border-green-400"
              : "bg-white border-gray-200"
          }`}
        >
          <h3 className="text-xl font-bold text-gray-900 mb-2 flex items-center gap-2">
            <Target
              className={`w-6 h-6 ${metaCumplida ? "text-green-600" : "text-gray-500"}`}
            />
            Meta mensual
          </h3>
          <p
            className={`text-5xl font-black ${metaCumplida ? "text-green-700" : "text-gray-900"}`}
          >
            {alistadasCount} / {META_MENSUAL}
          </p>
          <p
            className={`text-base font-semibold mt-2 ${metaCumplida ? "text-green-700" : "text-red-600"}`}
          >
            {metaCumplida ? "✓ Meta cumplida" : "Meta pendiente"}
          </p>
        </div>

        <div
          className={`rounded-xl shadow-lg border p-6 ${
            comisionExtraCumplida
              ? "bg-green-50 border-green-400"
              : "bg-white border-gray-200"
          }`}
        >
          <h3 className="text-xl font-bold text-gray-900 mb-2 flex items-center gap-2">
            <Award
              className={`w-6 h-6 ${comisionExtraCumplida ? "text-green-600" : "text-gray-500"}`}
            />
            Comisión extra
          </h3>
          <p
            className={`text-5xl font-black ${comisionExtraCumplida ? "text-green-700" : "text-gray-900"}`}
          >
            ≥ {COMISION_EXTRA_MIN}
          </p>
          <p
            className={`text-base font-semibold mt-2 ${comisionExtraCumplida ? "text-green-700" : "text-red-600"}`}
          >
            {comisionExtraCumplida
              ? "✓ Comisión extra alcanzada"
              : `Faltan ${COMISION_EXTRA_MIN - alistadasCount} alistadas`}
          </p>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-6">
        <h3 className="text-xl font-bold text-gray-900 mb-2 text-center">
          Alistadas · Meta · Comisión extra
        </h3>
        <p className="text-center text-gray-600 mb-4">
          Meta y comisión en{" "}
          <span className="text-green-600 font-semibold">verde</span> si se
          cumple, en <span className="text-red-600 font-semibold">rojo</span> si
          no
        </p>
        <ResponsiveContainer width="100%" height={320}>
          <PieChart>
            <Pie
              data={kpiPieData}
              cx="50%"
              cy="50%"
              outerRadius={120}
              dataKey="value"
              label={({ name }) => name}
            >
              {kpiPieData.map((entry) => (
                <Cell key={entry.name} fill={entry.color} />
              ))}
            </Pie>
            <Tooltip
              formatter={(_value: number, name: string) => {
                if (name.startsWith("Alistadas")) {
                  return [alistadasCount, name];
                }
                if (name.startsWith("Meta")) {
                  return [
                    `${alistadasCount} / ${META_MENSUAL}`,
                    metaCumplida ? "Meta cumplida" : "Meta pendiente",
                  ];
                }
                return [
                  `${alistadasCount} alistadas (mín. ${COMISION_EXTRA_MIN})`,
                  comisionExtraCumplida
                    ? "Comisión extra alcanzada"
                    : "Comisión extra pendiente",
                ];
              }}
            />
            <Legend />
          </PieChart>
        </ResponsiveContainer>
      </div>

      {/* Leyenda semáforo */}
      <div className="flex flex-wrap items-center gap-4 text-base font-semibold">
        <span className="inline-flex items-center gap-2 px-4 py-2 bg-green-200 rounded-lg border border-green-600">
          <span className="w-4 h-4 bg-green-600 rounded" aria-hidden="true" />
          <span>90% – 100%</span>
        </span>
        <span className="inline-flex items-center gap-2 px-4 py-2 bg-yellow-200 rounded-lg border border-yellow-500">
          <span className="w-4 h-4 bg-yellow-500 rounded" aria-hidden="true" />
          <span>50% – 89%</span>
        </span>
        <span className="inline-flex items-center gap-2 px-4 py-2 bg-red-200 rounded-lg border border-red-600">
          <span className="w-4 h-4 bg-red-600 rounded" aria-hidden="true" />
          <span>0% – 49%</span>
        </span>
      </div>

      {/* Tabla grande para TV */}
      <div className="bg-white rounded-xl shadow-lg border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-200 bg-gray-50 flex flex-wrap items-center justify-between gap-4">
          <div>
            <h3 className="text-2xl font-bold text-gray-900">
              Tablero por Prioridad — {cicloActivo || "Sin ciclo"}
            </h3>
            <p className="text-gray-600 mt-1 flex items-center gap-2">
              <ArrowUpDown className="w-4 h-4" />
              Ordena por prioridad o % avance, o arrastra filas. Las de 100%
              quedan al final.
            </p>
          </div>
          <fieldset className="flex flex-wrap gap-2 border-0 p-0 m-0">
            <legend className="sr-only">Orden de tabla</legend>
            <button
              type="button"
              className={sortButtonClass("prioridad")}
              onClick={() => handleSortModeChange("prioridad")}
            >
              Por prioridad
            </button>
            <button
              type="button"
              className={sortButtonClass("avance")}
              onClick={() => handleSortModeChange("avance")}
            >
              Por % avance
            </button>
            <button
              type="button"
              className={sortButtonClass("manual")}
              onClick={() => handleSortModeChange("manual")}
            >
              Orden manual
            </button>
          </fieldset>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[900px]">
            <thead>
              <tr className="bg-gray-600 text-white text-left">
                <th className="px-3 py-5 text-xl font-bold w-14" aria-label="Mover" />
                <th className="px-6 py-5 text-xl font-bold w-28">Prioridad</th>
                <th className="px-6 py-5 text-xl font-bold">Serie</th>
                <th className="px-6 py-5 text-xl font-bold">Modelo</th>
                <th className="px-6 py-5 text-xl font-bold">Cliente</th>
                <th className="px-6 py-5 text-xl font-bold">Asesor</th>
                <th className="px-6 py-5 text-xl font-bold">OTT</th>
                <th className="px-6 py-5 text-xl font-bold text-right w-36">
                  % Avance
                </th>
              </tr>
            </thead>
            <tbody>
              {displayItems.length === 0 ? (
                <tr>
                  <td
                    colSpan={8}
                    className="px-6 py-12 text-center text-xl text-gray-500"
                  >
                    No hay registros para el ciclo seleccionado
                  </td>
                </tr>
              ) : (
                displayItems.map((item) => {
                  const avance = parsePorcentajeAvance(item.fields);
                  const prioridad = getPrioridad(item.fields);
                  const isCompleted = avance === 100;
                  const isDragging = draggedId === item.id;

                  return (
                    <tr
                      key={item.id}
                      draggable={!isCompleted}
                      onDragStart={() => handleDragStart(item.id, avance)}
                      onDragOver={(e) => handleDragOver(e, item.id)}
                      onDrop={() => handleDrop(item.id)}
                      onDragEnd={() => setDraggedId(null)}
                      className={`border-b border-gray-300 ${getAvanceRowClass(avance)} ${
                        isDragging ? "opacity-50" : ""
                      } ${isCompleted ? "cursor-default" : "cursor-grab active:cursor-grabbing"}`}
                    >
                      <td className="px-3 py-5 text-center">
                        {isCompleted ? (
                          <span className="text-gray-400 text-xs font-semibold">
                            100%
                          </span>
                        ) : (
                          <GripVertical
                            className="w-6 h-6 text-gray-500 mx-auto"
                            aria-hidden="true"
                          />
                        )}
                      </td>
                      <td className="px-6 py-5">
                        <span className="inline-flex items-center justify-center min-w-[3rem] px-4 py-2 text-2xl font-black bg-gray-500 text-white rounded-lg">
                          {prioridad || "—"}
                        </span>
                      </td>
                      <td className="px-6 py-5 text-2xl font-bold text-gray-900">
                        {fieldStr(item.fields, "Serie") || "—"}
                      </td>
                      <td className="px-6 py-5 text-2xl font-semibold text-gray-800">
                        {fieldStr(item.fields, "Modelo") || "—"}
                      </td>
                      <td className="px-6 py-5 text-xl text-gray-800">
                        {fieldStr(item.fields, "Title") || "—"}
                      </td>
                      <td className="px-6 py-5 text-xl text-gray-800">
                        {fieldStr(item.fields, "Asesor") || "—"}
                      </td>
                      <td className="px-6 py-5 text-xl text-gray-800">
                        {fieldStr(item.fields, "OTT") || "—"}
                      </td>
                      <td
                        className={`px-6 py-5 text-3xl font-black text-right ${getAvanceTextClass(avance)}`}
                      >
                        {avance}%
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default CicloMesView;
