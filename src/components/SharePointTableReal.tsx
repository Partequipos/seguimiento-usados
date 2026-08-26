/**
 * Tabla optimizada para datos reales de SharePoint
 * Muestra las columnas más importantes
 */

import React, { useState, useMemo, useEffect } from "react";
import {
  SharePointListItem,
  sharePointService,
} from "../services/sharePointService";
import { sharePointConfig } from "../config/authConfig";
import {
  Edit2,
  Trash2,
  TestTube2,
  ChevronDown,
  ChevronUp,
  Search,
  X,
  Paperclip,
  Download,
  FileIcon,
  Loader2,
} from "lucide-react";
import { format } from "date-fns";
import { es } from "date-fns/locale";
import {
  getFieldValue,
  calcularPorcentajeAvance,
} from "../utils/sharePointFieldMapping";

interface SharePointTableRealProps {
  items: SharePointListItem[];
  useMockData: boolean;
  onEdit: (vehicle: SharePointListItem) => void;
  onDelete: (vehicleId: string) => void;
  onRefresh: () => void;
}

const SharePointTableReal: React.FC<SharePointTableRealProps> = ({
  items,
  useMockData,
  onEdit,
  onDelete,
}) => {
  const [expandedRow, setExpandedRow] = useState<string | null>(null);
  const [serieFilter, setSerieFilter] = useState<string>("");
  const [attachments, setAttachments] = useState<Record<string, any[]>>({});
  const [loadingAttachments, setLoadingAttachments] = useState<
    Record<string, boolean>
  >({});

  // Función helper para obtener porcentaje de avance
  const getPorcentajeAvance = (fields: Record<string, any>) => {
    let porcentaje = getFieldValue(fields, "PorcentajeAvanceTotal");

    // Si es un string (ej: "97%"), extraer el número
    if (typeof porcentaje === "string") {
      porcentaje =
        parseFloat(porcentaje.replace("%", "").replace(/[^0-9.]/g, "")) || 0;
    } else {
      porcentaje = Number(porcentaje) || 0;
    }

    return porcentaje;
  };

  // Extraer series únicas para el filtro
  const seriesUnicas = useMemo(() => {
    return [
      ...new Set(
        items
          .map((i) => {
            const serie = getFieldValue(i.fields, "Serie");
            return serie;
          })
          .filter(Boolean)
      ),
    ].sort();
  }, [items]);

  // Filtrar items por serie si hay filtro activo (búsqueda por coincidencia parcial)
  const filteredItems = useMemo(() => {
    if (!serieFilter) return items;
    return items.filter((item) => {
      const serie = String(
        getFieldValue(item.fields, "Serie") || ""
      ).toLowerCase();
      const filtro = serieFilter.toLowerCase();
      return serie.includes(filtro);
    });
  }, [items, serieFilter]);

  // Ordenar items por % Avance (menor a mayor)
  const sortedItems = [...filteredItems].sort((a, b) => {
    const avanceA = getPorcentajeAvance(a.fields);
    const avanceB = getPorcentajeAvance(b.fields);
    return avanceA - avanceB;
  });

  const formatDate = (dateString: string): string => {
    try {
      const date = new Date(dateString);
      return format(date, "dd/MM/yyyy", { locale: es });
    } catch {
      return "-";
    }
  };

  // Prioridad es solo un número, sin colores ni texto especial

  // Cargar adjuntos cuando se expande una fila
  useEffect(() => {
    const loadAttachments = async (itemId: string) => {
      if (!useMockData && !attachments[itemId] && !loadingAttachments[itemId]) {
        setLoadingAttachments((prev) => ({ ...prev, [itemId]: true }));
        try {
          const files = await sharePointService.getItemAttachments(itemId);
          setAttachments((prev) => ({ ...prev, [itemId]: files }));
        } catch (error) {
          console.error(`Error cargando adjuntos para item ${itemId}:`, error);
          setAttachments((prev) => ({ ...prev, [itemId]: [] }));
        } finally {
          setLoadingAttachments((prev) => ({ ...prev, [itemId]: false }));
        }
      }
    };

    if (expandedRow) {
      loadAttachments(expandedRow);
    }
  }, [expandedRow, useMockData, attachments, loadingAttachments]);

  const toggleRow = (id: string) => {
    setExpandedRow(expandedRow === id ? null : id);
  };

  // Nombres de las fases
  const faseNombres: Record<number, string> = {
    1: "LAVADO INICIAL",
    2: "INSPECCION",
    3: "DESENSAMBLE / ENSAMBLE INICIAL",
    4: "INSPECCION AJUSTES",
    5: "DESINSTALACION COMPONENTES MAYORES",
    6: "SOLDADURA / TORNO",
    7: "CABINA",
    8: "RODAJE / CILINDROS",
    9: "REPARACION COMPONENTES MAYORES",
    10: "ELECTRICIDAD / MISCELANEOS / CORRECCION FUGAS",
    11: "INSTALACION COMPONENTES MAYORES",
    12: "ENSAMBLE FINAL",
    13: "MANTENIMIENTO PREVENTIVO",
    14: "LAVADO PARA PINTURA",
    15: "PINTURA",
    16: "PREENTREGA",
  };

  return (
    <div className="w-full space-y-4">
      {useMockData && (
        <div className="bg-purple-50 border-l-4 border-purple-500 p-4 rounded-r-lg w-full">
          <div className="flex items-center gap-2">
            <TestTube2 className="w-5 h-5 text-purple-600" />
            <p className="text-purple-800 font-medium">Modo de Prueba Activo</p>
          </div>
          <p className="text-purple-700 text-sm mt-1">
            Puedes agregar, editar y eliminar registros. Los cambios se
            aplicarán localmente. Una vez que el administrador apruebe los
            permisos, los cambios se guardarán en SharePoint.
          </p>
        </div>
      )}

      <div className="bg-white rounded-lg shadow overflow-hidden w-full">
        <div className="overflow-x-auto w-full">
          <table
            className="w-full divide-y divide-gray-200"
            style={{ minWidth: "100%" }}
          >
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Título
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  <div className="space-y-2">
                    <div>Serie</div>
                    <div className="relative">
                      <div className="flex items-center gap-1">
                        <Search className="w-3 h-3 text-gray-400 absolute left-2 pointer-events-none" />
                        <input
                          type="text"
                          value={serieFilter}
                          onChange={(e) => setSerieFilter(e.target.value)}
                          placeholder="Buscar serie..."
                          className="w-full pl-7 pr-6 py-1 text-xs border border-gray-300 rounded focus:ring-1 focus:ring-brand-red/40 focus:border-blue-500 bg-white"
                          onClick={(e) => e.stopPropagation()}
                        />
                        {serieFilter && (
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setSerieFilter("");
                            }}
                            className="absolute right-2 text-gray-400 hover:text-gray-600"
                            title="Limpiar filtro"
                          >
                            <X className="w-3 h-3" />
                          </button>
                        )}
                      </div>
                      {seriesUnicas.length > 0 && (
                        <details className="mt-1">
                          <summary className="text-xs text-gray-500 cursor-pointer hover:text-gray-700">
                            Ver todas ({seriesUnicas.length})
                          </summary>
                          <div className="mt-1 max-h-32 overflow-y-auto border border-gray-200 rounded bg-white shadow-sm">
                            {seriesUnicas.map((serie) => (
                              <button
                                key={serie}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setSerieFilter(serie);
                                }}
                                className="w-full text-left px-2 py-1 text-xs hover:bg-brand-red-soft text-gray-700"
                              >
                                {serie}
                              </button>
                            ))}
                          </div>
                        </details>
                      )}
                    </div>
                  </div>
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Modelo
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Asesor
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Sede
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Prioridad
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  % Avance
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Días Restantes
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Acciones
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {sortedItems.map((item) => (
                <React.Fragment key={item.id}>
                  <tr className="hover:bg-gray-50">
                    <td className="px-6 py-4 text-sm font-medium text-gray-900">
                      {getFieldValue(item.fields, "Title") || "-"}
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-900">
                      {getFieldValue(item.fields, "Serie") || "-"}
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-900">
                      {getFieldValue(item.fields, "Modelo") || "-"}
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-900">
                      {getFieldValue(item.fields, "Asesor") || "-"}
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-900">
                      {getFieldValue(item.fields, "Sede") || "-"}
                    </td>
                    <td className="px-6 py-4 text-sm">
                      <span className="inline-flex px-2 py-1 text-xs font-medium text-gray-900">
                        {Number(getFieldValue(item.fields, "Prioridad")) || 0}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-sm">
                      <div className="flex items-center gap-2">
                        {(() => {
                          // Usar el valor calculado de SharePoint directamente
                          let porcentaje = getFieldValue(
                            item.fields,
                            "PorcentajeAvanceTotal"
                          );

                          // Si es un string (ej: "97%"), extraer el número
                          if (typeof porcentaje === "string") {
                            porcentaje =
                              parseFloat(
                                porcentaje
                                  .replace("%", "")
                                  .replace(/[^0-9.]/g, "")
                              ) || 0;
                          } else {
                            porcentaje = Number(porcentaje) || 0;
                          }

                          return (
                            <>
                              <div className="w-24 bg-gray-200 rounded-full h-2">
                                <div
                                  className="bg-brand-red h-2 rounded-full"
                                  style={{
                                    width: `${porcentaje}%`,
                                  }}
                                />
                              </div>
                              <span className="text-xs text-gray-600 w-10 text-right">
                                {porcentaje}%
                              </span>
                            </>
                          );
                        })()}
                      </div>
                    </td>
                    <td className="px-6 py-4 text-sm">
                      <span
                        className={`font-medium ${
                          (Number(
                            getFieldValue(item.fields, "DiasRestantes")
                          ) || 0) < 5
                            ? "text-red-600"
                            : (Number(
                                getFieldValue(item.fields, "DiasRestantes")
                              ) || 0) < 10
                            ? "text-yellow-600"
                            : "text-green-600"
                        }`}
                      >
                        {Number(getFieldValue(item.fields, "DiasRestantes")) ||
                          0}{" "}
                        días
                      </span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                      <button
                        onClick={() => toggleRow(item.id)}
                        className="text-gray-600 hover:text-gray-900 mr-3"
                        title="Ver detalles"
                      >
                        {expandedRow === item.id ? (
                          <ChevronUp className="w-5 h-5" />
                        ) : (
                          <ChevronDown className="w-5 h-5" />
                        )}
                      </button>
                      <button
                        onClick={() => onEdit(item)}
                        className="text-brand-red hover:text-brand-red-dark mr-3"
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

                  {/* Fila expandible con detalles */}
                  {expandedRow === item.id && (
                    <tr className="bg-gray-50">
                      <td colSpan={9} className="px-6 py-4">
                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                          {/* Información General */}
                          <div className="space-y-2">
                            <h4 className="font-semibold text-gray-900 mb-3">
                              Información General
                            </h4>
                            <div>
                              <span className="text-xs text-gray-600">
                                OTT:
                              </span>
                              <p className="text-sm font-medium text-gray-900">
                                {getFieldValue(item.fields, "OTT") || "-"}
                              </p>
                            </div>
                            <div>
                              <span className="text-xs text-gray-600">
                                Ciclo:
                              </span>
                              <p className="text-sm font-medium text-gray-900">
                                {getFieldValue(item.fields, "Ciclo") || "-"}
                              </p>
                            </div>
                            <div>
                              <span className="text-xs text-gray-600">
                                Observaciones:
                              </span>
                              <p className="text-sm font-medium text-gray-900">
                                {getFieldValue(item.fields, "Observaciones") ||
                                  "-"}
                              </p>
                            </div>

                            {/* Adjuntos */}
                            <div className="mt-4 pt-4 border-t border-gray-200">
                              <div className="flex items-center gap-2 mb-2">
                                <Paperclip className="w-4 h-4 text-gray-600" />
                                <span className="text-xs font-semibold text-gray-700">
                                  Archivos Adjuntos
                                </span>
                              </div>
                              {loadingAttachments[item.id] ? (
                                <div className="flex items-center gap-2 text-sm text-gray-500">
                                  <Loader2 className="w-4 h-4 animate-spin" />
                                  <span>Cargando adjuntos...</span>
                                </div>
                              ) : attachments[item.id] &&
                                attachments[item.id].length > 0 ? (
                                <div className="space-y-1.5">
                                  {attachments[item.id].map(
                                    (file: any, index: number) => {
                                      // SharePoint REST API devuelve ServerRelativeUrl
                                      // Extraer el dominio de SharePoint desde siteUrl
                                      const getSharePointDomain = () => {
                                        try {
                                          const url = new URL(
                                            sharePointConfig.siteUrl
                                          );
                                          return `${url.protocol}//${url.hostname}`;
                                        } catch {
                                          return sharePointConfig.siteUrl.split(
                                            "/sites/"
                                          )[0];
                                        }
                                      };

                                      const fileUrl = file.ServerRelativeUrl
                                        ? `${getSharePointDomain()}${
                                            file.ServerRelativeUrl
                                          }`
                                        : file[
                                            "@microsoft.graph.downloadUrl"
                                          ] || file.webUrl;

                                      return (
                                        <a
                                          key={index}
                                          href={fileUrl}
                                          target="_blank"
                                          rel="noopener noreferrer"
                                          className="flex items-center gap-2 p-2 bg-white rounded border border-gray-200 hover:border-brand-red-muted hover:bg-brand-red-soft transition-colors group"
                                        >
                                          <FileIcon className="w-4 h-4 text-brand-red flex-shrink-0" />
                                          <span className="text-xs text-gray-900 truncate flex-1">
                                            {file.FileName || file.name}
                                          </span>
                                          <Download className="w-3.5 h-3.5 text-gray-400 group-hover:text-brand-red flex-shrink-0" />
                                        </a>
                                      );
                                    }
                                  )}
                                </div>
                              ) : (
                                <p className="text-xs text-gray-400 italic">
                                  Sin archivos adjuntos
                                </p>
                              )}
                            </div>
                          </div>

                          {/* Fechas */}
                          <div className="space-y-2">
                            <h4 className="font-semibold text-gray-900 mb-3">
                              Fechas
                            </h4>
                            <div>
                              <span className="text-xs text-gray-600">
                                Solicitud:
                              </span>
                              <p className="text-sm font-medium text-gray-900">
                                {formatDate(
                                  getFieldValue(
                                    item.fields,
                                    "FechaSolicitud"
                                  ) || ""
                                )}
                              </p>
                            </div>
                            <div>
                              <span className="text-xs text-gray-600">
                                Compromiso Comercial:
                              </span>
                              <p className="text-sm font-medium text-gray-900">
                                {formatDate(
                                  getFieldValue(
                                    item.fields,
                                    "FechaCompromisoComercial"
                                  ) || ""
                                )}
                              </p>
                            </div>
                            <div>
                              <span className="text-xs text-gray-600">
                                Inicio Ciclo:
                              </span>
                              <p className="text-sm font-medium text-gray-900">
                                {formatDate(
                                  getFieldValue(
                                    item.fields,
                                    "FechaInicioCiclo"
                                  ) || ""
                                )}
                              </p>
                            </div>
                            <div>
                              <span className="text-xs text-gray-600">
                                Final Alistamiento:
                              </span>
                              <p className="text-sm font-medium text-gray-900">
                                {getFieldValue(
                                  item.fields,
                                  "FechaFinalAlistamiento"
                                )
                                  ? formatDate(
                                      getFieldValue(
                                        item.fields,
                                        "FechaFinalAlistamiento"
                                      ) || ""
                                    )
                                  : "-"}
                              </p>
                            </div>
                          </div>

                          {/* Fases */}
                          <div className="space-y-2 md:col-span-2">
                            <h4 className="font-semibold text-gray-900 mb-3">
                              Estado de Fases (F1-F16)
                            </h4>
                            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                              {Array.from({ length: 16 }, (_, i) => i + 1).map(
                                (num) => {
                                  const porcentaje =
                                    getFieldValue(item.fields, `F${num}`) ||
                                    "0%";
                                  const bgColor =
                                    porcentaje === "100%"
                                      ? "bg-green-500"
                                      : porcentaje === "75%"
                                      ? "bg-lime-500"
                                      : porcentaje === "50%"
                                      ? "bg-yellow-500"
                                      : porcentaje === "25%"
                                      ? "bg-orange-500"
                                      : "bg-red-500";
                                  return (
                                    <div
                                      key={`F${num}`}
                                      className="bg-white rounded-lg border border-gray-200 p-2 hover:shadow-md transition-shadow"
                                      title={`${faseNombres[num]}: ${porcentaje}`}
                                    >
                                      <div className="flex items-center justify-between mb-1">
                                        <span className="text-xs font-bold text-gray-700">
                                          F{num}
                                        </span>
                                        <div
                                          className={`${bgColor} text-white text-xs font-bold rounded px-2 py-0.5`}
                                        >
                                          {porcentaje}
                                        </div>
                                      </div>
                                      <p className="text-[10px] text-gray-500 leading-tight">
                                        {faseNombres[num]}
                                      </p>
                                    </div>
                                  );
                                }
                              )}
                            </div>

                            {/* Leyenda actualizada */}
                            <div className="flex flex-wrap gap-3 text-xs mt-4 p-3 bg-gray-100 rounded-lg">
                              <span className="font-semibold text-gray-700">
                                Leyenda:
                              </span>
                              <div className="flex items-center gap-1">
                                <div className="w-4 h-4 bg-green-500 rounded shadow-sm"></div>
                                <span className="font-medium">
                                  100% - Completada
                                </span>
                              </div>
                              <div className="flex items-center gap-1">
                                <div className="w-4 h-4 bg-lime-500 rounded shadow-sm"></div>
                                <span className="font-medium">
                                  75% - Avanzada
                                </span>
                              </div>
                              <div className="flex items-center gap-1">
                                <div className="w-4 h-4 bg-yellow-500 rounded shadow-sm"></div>
                                <span className="font-medium">
                                  50% - En Progreso
                                </span>
                              </div>
                              <div className="flex items-center gap-1">
                                <div className="w-4 h-4 bg-orange-500 rounded shadow-sm"></div>
                                <span className="font-medium">
                                  25% - Iniciada
                                </span>
                              </div>
                              <div className="flex items-center gap-1">
                                <div className="w-4 h-4 bg-red-500 rounded shadow-sm"></div>
                                <span className="font-medium">
                                  0% - No Iniciada
                                </span>
                              </div>
                            </div>
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {sortedItems.length === 0 && (
        <div className="text-center py-12 bg-white rounded-lg">
          <p className="text-gray-500">No hay equipos registrados</p>
        </div>
      )}
    </div>
  );
};

export default SharePointTableReal;
