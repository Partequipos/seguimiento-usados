/**
 * Visualización Ciclo-Mes — tablero para proyección en TV
 * Tabla por prioridad (1-10) con semáforo de avance y KPIs de meta/comisión.
 */

import React, { useMemo, useState } from "react";
import {
  PieChart,
  Pie,
  Cell,
  Tooltip,
  ResponsiveContainer,
  Legend,
} from "recharts";
import { Target, Award, CheckCircle2, Filter } from "lucide-react";
import { SharePointListItem } from "../services/sharePointService";
import { getFieldValue, parsePorcentajeAvance } from "../utils/sharePointFieldMapping";

const META_MENSUAL = 12;
const COMISION_EXTRA_MIN = 10;

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
  const value = Number(getFieldValue(fields, "Prioridad")) || 0;
  return value;
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

function sortByPrioridad(a: SharePointListItem, b: SharePointListItem): number {
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

const CicloMesView: React.FC<CicloMesViewProps> = ({ items }) => {
  const ciclosDisponibles = useMemo(() => {
    const set = new Set<string>();
    for (const item of items) {
      const ciclo = normalizeCicloValue(getFieldValue(item.fields, "Ciclo"));
      if (ciclo) set.add(ciclo);
    }
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }, [items]);

  const [selectedCiclo, setSelectedCiclo] = useState<string>("");

  const cicloActivo = selectedCiclo || ciclosDisponibles[0] || "";

  const itemsDelCiclo = useMemo(() => {
    if (!cicloActivo) return [];
    return items.filter(
      (item) =>
        normalizeCicloValue(getFieldValue(item.fields, "Ciclo")) === cicloActivo
    );
  }, [items, cicloActivo]);

  const sortedItems = useMemo(
    () => [...itemsDelCiclo].sort(sortByPrioridad),
    [itemsDelCiclo]
  );

  const alistadasCount = useMemo(
    () =>
      itemsDelCiclo.filter(
        (item) => parsePorcentajeAvance(item.fields) === 100
      ).length,
    [itemsDelCiclo]
  );

  const metaCumplida = alistadasCount >= META_MENSUAL;
  const comisionExtraCumplida = alistadasCount >= COMISION_EXTRA_MIN;

  const pieData = useMemo(
    () => [
      {
        name: "Alistadas (100%)",
        value: alistadasCount,
        color: "#16a34a",
      },
      {
        name: "Pendientes / En proceso",
        value: Math.max(itemsDelCiclo.length - alistadasCount, 0),
        color: "#9ca3af",
      },
    ].filter((d) => d.value > 0),
    [alistadasCount, itemsDelCiclo.length]
  );

  const metaPieData = useMemo(
    () => [
      {
        name: "Alistadas",
        value: alistadasCount,
        color: metaCumplida ? "#16a34a" : "#ef4444",
      },
      {
        name: "Faltan para meta",
        value: Math.max(META_MENSUAL - alistadasCount, 0),
        color: "#e5e7eb",
      },
    ].filter((d) => d.value > 0),
    [alistadasCount, metaCumplida]
  );

  return (
    <div className="w-full space-y-6">
      {/* Selector de ciclo */}
      <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-6 flex flex-wrap items-end gap-4">
        <div className="flex items-center gap-2 text-gray-800">
          <Filter className="w-6 h-6 text-blue-600" />
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
            className="w-full px-4 py-3 text-lg border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white"
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
          <span className="font-bold text-gray-900">{sortedItems.length}</span>{" "}
          equipos en {cicloActivo || "—"}
        </p>
      </div>

      {/* KPIs y gráficos */}
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

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-6">
          <h3 className="text-xl font-bold text-gray-900 mb-4">
            Distribución del ciclo
          </h3>
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie
                data={pieData}
                cx="50%"
                cy="50%"
                outerRadius={100}
                dataKey="value"
                label={({ name, value }) =>
                  `${name}: ${typeof value === "number" ? value : 0}`
                }
              >
                {pieData.map((entry) => (
                  <Cell key={entry.name} fill={entry.color} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-6">
          <h3 className="text-xl font-bold text-gray-900 mb-4">
            Avance vs Meta ({META_MENSUAL})
          </h3>
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie
                data={metaPieData}
                cx="50%"
                cy="50%"
                outerRadius={100}
                dataKey="value"
                label={({ name, value }) =>
                  `${name}: ${typeof value === "number" ? value : 0}`
                }
              >
                {metaPieData.map((entry) => (
                  <Cell key={entry.name} fill={entry.color} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </div>
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
        <div className="px-6 py-4 border-b border-gray-200 bg-gray-50">
          <h3 className="text-2xl font-bold text-gray-900">
            Tablero por Prioridad — {cicloActivo || "Sin ciclo"}
          </h3>
          <p className="text-gray-600 mt-1">
            Orden: prioridad 1 → 10 · Proyección para técnicos de alistamiento
          </p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[900px]">
            <thead>
              <tr className="bg-gray-600 text-white text-left">
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
              {sortedItems.length === 0 ? (
                <tr>
                  <td
                    colSpan={7}
                    className="px-6 py-12 text-center text-xl text-gray-500"
                  >
                    No hay registros para el ciclo seleccionado
                  </td>
                </tr>
              ) : (
                sortedItems.map((item) => {
                  const avance = parsePorcentajeAvance(item.fields);
                  const prioridad = getPrioridad(item.fields);
                  return (
                    <tr
                      key={item.id}
                      className={`border-b border-gray-300 ${getAvanceRowClass(avance)}`}
                    >
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
