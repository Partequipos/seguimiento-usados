/**
 * Tabla simplificada para mostrar datos de SharePoint
 */

import React from "react";
import { SharePointListItem } from "../services/sharePointService";
import { Edit2, Trash2, TestTube2 } from "lucide-react";
import { format } from "date-fns";
import { es } from "date-fns/locale";

interface SharePointDataTableSimpleProps {
  items: SharePointListItem[];
  columns: any[];
  useMockData: boolean;
  onEdit: (vehicle: SharePointListItem) => void;
  onDelete: (vehicleId: string) => void;
}

const SharePointDataTableSimple: React.FC<SharePointDataTableSimpleProps> = ({
  items,
  columns,
  useMockData,
  onEdit,
  onDelete,
}) => {
  const formatValue = (value: any, columnType?: string): string => {
    if (value === null || value === undefined) return "-";

    // Formatear fechas
    if (columnType === "dateTime" || value instanceof Date) {
      try {
        const date = typeof value === "string" ? new Date(value) : value;
        return format(date, "dd/MM/yyyy", { locale: es });
      } catch {
        return String(value);
      }
    }

    // Formatear números con separador de miles
    if (typeof value === "number" && !columnType) {
      return value.toLocaleString("es-CO");
    }

    // Formatear currency
    if (columnType === "currency") {
      return new Intl.NumberFormat("es-CO", {
        style: "currency",
        currency: "COP",
        minimumFractionDigits: 0,
      }).format(value);
    }

    // Formatear booleanos
    if (typeof value === "boolean") {
      return value ? "Sí" : "No";
    }

    // Formatear objetos (ej: lookup fields)
    if (typeof value === "object") {
      if (value.LookupValue) return value.LookupValue;
      return JSON.stringify(value);
    }

    return String(value);
  };

  const getEstadoBadge = (estado: string) => {
    const badges: Record<string, string> = {
      "En Proceso": "bg-brand-red-soft text-brand-red-dark",
      Completado: "bg-green-100 text-green-800",
      Pendiente: "bg-yellow-100 text-yellow-800",
    };
    return badges[estado] || "bg-gray-100 text-gray-800";
  };

  const getPrioridadBadge = (prioridad: string) => {
    const badges: Record<string, string> = {
      Alta: "bg-red-100 text-red-800",
      Media: "bg-yellow-100 text-yellow-800",
      Baja: "bg-green-100 text-green-800",
    };
    return badges[prioridad] || "bg-gray-100 text-gray-800";
  };

  return (
    <div className="space-y-4">
      {useMockData && (
        <div className="bg-brand-gray-soft border-l-4 border-brand-gray p-4 rounded-r-lg">
          <div className="flex items-center gap-2">
            <TestTube2 className="w-5 h-5 text-brand-gray" />
            <p className="text-brand-ink font-medium">Modo de Prueba Activo</p>
          </div>
          <p className="text-brand-gray text-sm mt-1">
            Puedes agregar, editar y eliminar registros. Los cambios se
            aplicarán localmente. Una vez que el administrador apruebe los
            permisos, los cambios se guardarán en SharePoint.
          </p>
        </div>
      )}

      <div className="bg-white rounded-lg shadow overflow-hidden">
        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                {columns.map((column) => (
                  <th
                    key={column.id}
                    className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider"
                  >
                    {column.displayName || column.name}
                  </th>
                ))}
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Acciones
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {items.map((item) => (
                <tr key={item.id} className="hover:bg-gray-50">
                  {columns.map((column) => (
                    <td
                      key={`${item.id}-${column.name}`}
                      className="px-6 py-4 text-sm text-gray-900"
                    >
                      {column.name === "Estado" ? (
                        <span
                          className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getEstadoBadge(
                            item.fields[column.name]
                          )}`}
                        >
                          {item.fields[column.name]}
                        </span>
                      ) : column.name === "Prioridad" ? (
                        <span
                          className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getPrioridadBadge(
                            item.fields[column.name]
                          )}`}
                        >
                          {item.fields[column.name]}
                        </span>
                      ) : column.name === "PorcentajeAvance" ? (
                        <div className="flex items-center gap-2">
                          <div className="w-full bg-gray-200 rounded-full h-2.5">
                            <div
                              className="bg-brand-red h-2.5 rounded-full"
                              style={{
                                width: `${item.fields[column.name]}%`,
                              }}
                            />
                          </div>
                          <span className="text-xs text-gray-600 whitespace-nowrap">
                            {item.fields[column.name]}%
                          </span>
                        </div>
                      ) : column.name === "Observaciones" ? (
                        <div
                          className="max-w-xs truncate"
                          title={item.fields[column.name]}
                        >
                          {formatValue(item.fields[column.name], column.type)}
                        </div>
                      ) : (
                        formatValue(item.fields[column.name], column.type)
                      )}
                    </td>
                  ))}
                  <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                    <button
                      onClick={() => onEdit(item)}
                      className="text-brand-red hover:text-brand-red-dark mr-4"
                      title="Editar"
                    >
                      <Edit2 className="w-5 h-5" />
                    </button>
                    <button
                      onClick={() => onDelete(item.id)}
                      className="text-red-600 hover:text-red-900"
                      title="Eliminar"
                    >
                      <Trash2 className="w-5 h-5" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {items.length === 0 && (
        <div className="text-center py-12 bg-white rounded-lg">
          <p className="text-gray-500">No hay equipos registrados</p>
        </div>
      )}
    </div>
  );
};

export default SharePointDataTableSimple;
