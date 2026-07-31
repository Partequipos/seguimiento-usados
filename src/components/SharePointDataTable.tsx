/**
 * Componente de tabla para mostrar datos de SharePoint
 */

import React, { useEffect, useState } from "react";
import {
  sharePointService,
  SharePointListItem,
} from "../services/sharePointService";
import { useAuth } from "../context/AuthContext";
import {
  Loader2,
  RefreshCw,
  AlertCircle,
  Table2,
  TestTube2,
} from "lucide-react";
import { format } from "date-fns";
import { es } from "date-fns/locale";
import { mockItems, mockColumns } from "../data/mockData";

const SharePointDataTable: React.FC = () => {
  const { isAuthenticated } = useAuth();
  const [items, setItems] = useState<SharePointListItem[]>([]);
  const [columns, setColumns] = useState<any[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [useMockData, setUseMockData] = useState(false);

  useEffect(() => {
    if (isAuthenticated && !useMockData) {
      loadData();
    } else if (!isAuthenticated) {
      // Si no está autenticado, cargar datos mock automáticamente
      loadMockData();
    }
  }, [isAuthenticated, useMockData]);

  const loadMockData = () => {
    setIsLoading(true);
    // Simular carga
    setTimeout(() => {
      setItems(mockItems);
      setColumns(mockColumns);
      setUseMockData(true);
      setIsLoading(false);
      setError(null);
    }, 500);
  };

  const loadData = async () => {
    try {
      setIsLoading(true);
      setError(null);

      // Cargar columnas y datos en paralelo
      const [itemsData, columnsData] = await Promise.all([
        sharePointService.getListItems(),
        sharePointService.getListColumns(),
      ]);

      setItems(itemsData);
      // Filtrar solo columnas visibles y no del sistema
      const visibleColumns = columnsData.filter(
        (col) => !col.hidden && !col.name.startsWith("_")
      );
      setColumns(visibleColumns);
      setUseMockData(false);
    } catch (err: any) {
      console.error("Error cargando datos:", err);
      setError(err.message || "Error al cargar los datos de SharePoint");
      // Si falla, ofrecer cargar datos mock
      if (isAuthenticated) {
        setError(
          err.message +
            " - Puedes usar datos de prueba mientras esperas la aprobación del administrador."
        );
      }
    } finally {
      setIsLoading(false);
    }
  };

  const formatValue = (value: any, columnType?: string): string => {
    if (value === null || value === undefined) return "-";

    // Formatear fechas
    if (columnType === "dateTime" || value instanceof Date) {
      try {
        const date = typeof value === "string" ? new Date(value) : value;
        return format(date, "dd/MM/yyyy HH:mm", { locale: es });
      } catch {
        return String(value);
      }
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

  if (!isAuthenticated && !useMockData && items.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center p-8 bg-brand-red-soft rounded-lg">
        <AlertCircle className="w-12 h-12 text-brand-red mb-4" />
        <p className="text-lg text-gray-700 mb-4">
          Por favor inicia sesión para ver los datos de SharePoint
        </p>
        <p className="text-sm text-gray-600 mb-4">
          O puedes usar datos de prueba mientras esperas la aprobación del
          administrador
        </p>
        <button
          onClick={loadMockData}
          className="flex items-center gap-2 px-4 py-2 bg-brand-gray hover:bg-brand-ink text-white rounded-lg transition-colors"
        >
          <TestTube2 className="w-4 h-4" />
          Ver Datos de Prueba
        </button>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="flex flex-col items-center justify-center p-12">
        <Loader2 className="w-12 h-12 text-brand-red animate-spin mb-4" />
        <p className="text-gray-600">Cargando datos de SharePoint...</p>
      </div>
    );
  }

  if (error && !useMockData) {
    return (
      <div className="bg-red-50 border border-red-200 rounded-lg p-6">
        <div className="flex items-start gap-3">
          <AlertCircle className="w-6 h-6 text-red-600 flex-shrink-0 mt-0.5" />
          <div className="flex-1">
            <h3 className="text-lg font-semibold text-red-900 mb-2">
              Error al cargar los datos
            </h3>
            <p className="text-red-700 mb-4">{error}</p>
            <div className="flex gap-3">
              <button
                onClick={loadData}
                className="flex items-center gap-2 px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg transition-colors"
              >
                <RefreshCw className="w-4 h-4" />
                Reintentar
              </button>
              <button
                onClick={loadMockData}
                className="flex items-center gap-2 px-4 py-2 bg-brand-gray hover:bg-brand-ink text-white rounded-lg transition-colors"
              >
                <TestTube2 className="w-4 h-4" />
                Usar Datos de Prueba
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (items.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center p-8 bg-gray-50 rounded-lg">
        <Table2 className="w-12 h-12 text-gray-400 mb-4" />
        <p className="text-lg text-gray-600 mb-4">
          No se encontraron datos en la lista de SharePoint
        </p>
        <button
          onClick={loadData}
          className="flex items-center gap-2 px-4 py-2 bg-brand-red hover:bg-brand-red-dark text-white rounded-lg transition-colors"
        >
          <RefreshCw className="w-4 h-4" />
          Recargar
        </button>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {useMockData && (
        <div className="bg-brand-gray-soft border-l-4 border-brand-gray p-4 rounded-r-lg">
          <div className="flex items-center gap-2">
            <TestTube2 className="w-5 h-5 text-brand-gray" />
            <p className="text-brand-ink font-medium">Modo de Prueba Activo</p>
          </div>
          <p className="text-brand-gray text-sm mt-1">
            Estás viendo datos de ejemplo. Una vez que el administrador apruebe
            los permisos, podrás ver los datos reales de SharePoint.
          </p>
        </div>
      )}

      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-gray-900">
            {useMockData ? "Datos de Prueba" : "Datos de SharePoint"}
          </h2>
          <p className="text-sm text-gray-500 mt-1">
            {items.length} registro{items.length !== 1 ? "s" : ""} encontrado
            {items.length !== 1 ? "s" : ""}
          </p>
        </div>
        <div className="flex gap-2">
          {useMockData && isAuthenticated && (
            <button
              onClick={loadData}
              disabled={isLoading}
              className="flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white rounded-lg transition-colors"
            >
              <RefreshCw
                className={`w-4 h-4 ${isLoading ? "animate-spin" : ""}`}
              />
              Cargar Datos Reales
            </button>
          )}
          <button
            onClick={useMockData ? loadMockData : loadData}
            disabled={isLoading}
            className="flex items-center gap-2 px-4 py-2 bg-brand-red hover:bg-brand-red-dark disabled:bg-gray-400 text-white rounded-lg transition-colors"
          >
            <RefreshCw
              className={`w-4 h-4 ${isLoading ? "animate-spin" : ""}`}
            />
            Recargar
          </button>
        </div>
      </div>

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
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {items.map((item) => (
                <tr key={item.id} className="hover:bg-gray-50">
                  {columns.map((column) => (
                    <td
                      key={`${item.id}-${column.name}`}
                      className="px-6 py-4 whitespace-nowrap text-sm text-gray-900"
                    >
                      {formatValue(item.fields[column.name], column.type)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default SharePointDataTable;
