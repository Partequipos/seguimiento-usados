import { useState, useEffect, useMemo, useCallback } from "react";
import { AuthProvider, useAuth } from "./context/AuthContext";
import LoginButton from "./components/LoginButton";
import SharePointTableReal from "./components/SharePointTableReal";
import DashboardReal from "./components/DashboardReal";
import VehicleFormReal from "./components/VehicleFormReal";
import DashboardFilters, { FilterState } from "./components/DashboardFilters";
import {
  sharePointService,
  SharePointListItem,
} from "./services/sharePointService";
import { realItems } from "./data/mockDataReal";
import CicloMesView from "./components/CicloMesView";
import { LayoutDashboard, Table2, Plus, Loader2, Calendar, Monitor } from "lucide-react";
import {
  normalizeSharePointFields,
  getFieldValue,
  matchesPorcentajeAvanceFilter,
  parsePorcentajeAvance,
  toDateOnlyString,
} from "./utils/sharePointFieldMapping";

type View = "dashboard" | "table" | "ciclo-mes";

/** Payload del formulario de equipo (VehicleFormReal) */
interface VehicleFormPayload {
  Title?: string;
  Serie?: string;
  Prioridad?: number;
  Modelo?: string;
  OTT?: string;
  Asesor?: string;
  Sede?: string;
  FechaSolicitud?: string;
  FechaCompromisoComercial?: string;
  FechaInicioCiclo?: string;
  FechaFinalAlistamiento?: string;
  Observaciones?: string;
  Ciclo?: number;
  F1?: string;
  F2?: string;
  F3?: string;
  F4?: string;
  F5?: string;
  F6?: string;
  F7?: string;
  F8?: string;
  F9?: string;
  F10?: string;
  F11?: string;
  F12?: string;
  F13?: string;
  F14?: string;
  F15?: string;
  F16?: string;
}

function AppContent() {
  const { isAuthenticated } = useAuth();
  const [currentView, setCurrentView] = useState<View>("dashboard");
  const [items, setItems] = useState<SharePointListItem[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [useMockData, setUseMockData] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [editingVehicle, setEditingVehicle] =
    useState<SharePointListItem | null>(null);
  const [filters, setFilters] = useState<FilterState>({
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

  const loadMockData = useCallback(() => {
    setItems(realItems);
    setUseMockData(true);
  }, []);

  const loadRealData = useCallback(async () => {
    try {
      setIsLoading(true);
      const itemsData = await sharePointService.getListItems();
      console.log(`📊 Registros recibidos de SharePoint: ${itemsData.length}`);
      const normalizedItems = itemsData.map((item) => ({
        ...item,
        fields: normalizeSharePointFields(item.fields),
      }));
      console.log(
        `✅ Registros normalizados y listos para mostrar: ${normalizedItems.length}`
      );
      console.log(
        `📋 Primeros 5 registros normalizados:`,
        normalizedItems.slice(0, 5).map((i) => ({
          id: i.id,
          title: i.fields.Title,
          serie: getFieldValue(i.fields, "Serie"),
        }))
      );
      setItems(normalizedItems);
      setUseMockData(false);
    } catch (error) {
      console.error("Error loading data:", error);
      loadMockData();
    } finally {
      setIsLoading(false);
    }
  }, [loadMockData]);

  useEffect(() => {
    if (!isAuthenticated) {
      loadMockData();
    } else {
      loadRealData();
    }
  }, [isAuthenticated, loadMockData, loadRealData]);

  // Filtrar items
  const filteredItems = useMemo(() => {
    console.log(`🔍 Filtrando ${items.length} registros con filtros:`, filters);
    const filtered = items.filter((item) => {
      // Filtro por Sede
      if (filters.sede && getFieldValue(item.fields, "Sede") !== filters.sede)
        return false;

      // Filtro por Asesor
      if (
        filters.asesor &&
        getFieldValue(item.fields, "Asesor") !== filters.asesor
      )
        return false;

      // Filtro por Cliente
      if (
        filters.cliente &&
        getFieldValue(item.fields, "Title") !== filters.cliente
      )
        return false;

      // Filtro por Serie
      if (
        filters.serie &&
        getFieldValue(item.fields, "Serie") !== filters.serie
      )
        return false;

      // Filtro por Observaciones
      if (
        filters.observaciones &&
        getFieldValue(item.fields, "Observaciones") !== filters.observaciones
      )
        return false;

      // Filtro por Ciclo (selección múltiple)
      if (filters.ciclo.length > 0) {
        const raw = getFieldValue(item.fields, "Ciclo");
        let itemCicloStr = "";
        if (typeof raw === "string") itemCicloStr = raw;
        else if (raw != null && typeof raw === "number")
          itemCicloStr = String(raw);
        if (!filters.ciclo.includes(itemCicloStr)) return false;
      }

      // Filtro por Fecha de Compromiso (columna FechaCompromisoComercial) - comparación por día
      if (filters.fechaCompromisoDesde || filters.fechaCompromisoHasta) {
        const fechaCompromisoValue = getFieldValue(
          item.fields,
          "FechaCompromisoComercial"
        );
        if (!fechaCompromisoValue) return false;
        const fechaCompromiso = new Date(
          fechaCompromisoValue as string | number | Date
        );
        if (Number.isNaN(fechaCompromiso.getTime())) return false;
        const itemDateStr = toDateOnlyString(fechaCompromiso);
        // Los inputs type="date" ya envían YYYY-MM-DD; usarlos directos evita desfase por zona horaria
        if (filters.fechaCompromisoDesde) {
          if (itemDateStr < filters.fechaCompromisoDesde) return false;
        }
        if (filters.fechaCompromisoHasta) {
          if (itemDateStr > filters.fechaCompromisoHasta) return false;
        }
      }

      // Filtro por Fecha Final Alistamiento (columna FechaFinalAlistamiento en SharePoint) - comparación por día
      if (filters.fechaFinalDesde || filters.fechaFinalHasta) {
        const fechaFinalValue = getFieldValue(
          item.fields,
          "FechaFinalAlistamiento"
        );
        if (!fechaFinalValue) return false;
        const fechaFinal = new Date(
          fechaFinalValue as string | number | Date
        );
        if (Number.isNaN(fechaFinal.getTime())) return false;
        const itemDateStr = toDateOnlyString(fechaFinal);
        // Los inputs type="date" ya envían YYYY-MM-DD; usarlos directos evita desfase por zona horaria
        if (filters.fechaFinalDesde) {
          if (itemDateStr < filters.fechaFinalDesde) return false;
        }
        if (filters.fechaFinalHasta) {
          if (itemDateStr > filters.fechaFinalHasta) return false;
        }
      }

      // Filtro por % de Avance (incluye escalones 15%…99% indexados)
      if (filters.porcentajeAvance) {
        const avance = parsePorcentajeAvance(item.fields);
        if (
          !matchesPorcentajeAvanceFilter(avance, filters.porcentajeAvance)
        ) {
          return false;
        }
      }

      return true;
    });
    console.log(`✅ Registros después de filtrar: ${filtered.length}`);
    return filtered;
  }, [items, filters]);

  const handleAddVehicle = async (
    data: VehicleFormPayload,
    files?: File[]
  ) => {
    if (useMockData) {
      // Simular guardado en mock
      const newItem: SharePointListItem = {
        id: String(items.length + 1),
        fields: {
          ...data,
          FechaSolicitud: data.FechaSolicitud + "T00:00:00Z",
          FechaCompromisoComercial:
            data.FechaCompromisoComercial + "T00:00:00Z",
          FechaInicioCiclo: data.FechaInicioCiclo + "T00:00:00Z",
          FechaFinalAlistamiento: data.FechaFinalAlistamiento
            ? data.FechaFinalAlistamiento + "T00:00:00Z"
            : null,
          PorcentajeAvanceTotal: 0, // Se calculará en SharePoint
          DiasRestantes: 30, // Se calculará en SharePoint
        },
        createdDateTime: new Date().toISOString(),
        lastModifiedDateTime: new Date().toISOString(),
      };
      setItems([...items, newItem]);
      setShowForm(false);
      alert(
        "Equipo agregado (modo prueba). En producción se guardará en SharePoint."
      );
    } else {
      // Guardar en SharePoint - Convertir nombres amigables a nombres internos
      try {
        const sharePointFields: Record<string, unknown> = {
          Title: data.Title || "",
          field_0: data.Serie || "",
          field_1: data.Prioridad ?? 0,
          field_2: data.Modelo || "",
          field_3: data.OTT || "",
          field_4: data.Asesor || "",
          field_28: data.Sede || "",
        };

        if (data.FechaSolicitud) {
          sharePointFields.field_7 = new Date(
            data.FechaSolicitud + "T00:00:00Z"
          ).toISOString();
        }
        if (data.FechaCompromisoComercial) {
          sharePointFields.field_9 = new Date(
            data.FechaCompromisoComercial + "T00:00:00Z"
          ).toISOString();
        }
        if (data.FechaInicioCiclo) {
          sharePointFields.field_10 = new Date(
            data.FechaInicioCiclo + "T00:00:00Z"
          ).toISOString();
        }
        if (data.FechaFinalAlistamiento) {
          sharePointFields.FechaFinalAlistamiento = new Date(
            data.FechaFinalAlistamiento + "T00:00:00Z"
          ).toISOString();
        }

        if (data.Observaciones) {
          sharePointFields.field_8 = data.Observaciones;
        }
        if (data.Ciclo !== undefined && data.Ciclo !== null) {
          sharePointFields.field_29 = `Ciclo ${data.Ciclo}`;
        }

        for (let i = 1; i <= 16; i++) {
          const faseKey = `F${i}` as keyof VehicleFormPayload;
          const value = data[faseKey];
          if (value) {
            sharePointFields[`field_${10 + i}`] = value;
          }
        }

        const newItem = await sharePointService.createListItem(sharePointFields);
        console.log(`✅ Equipo creado con ID: ${newItem.id}`);

        let uploadedCount = 0;
        let failedCount = 0;
        if (files && files.length > 0) {
          console.log(`📎 Iniciando subida de ${files.length} archivo(s)...`);
          await new Promise((resolve) => setTimeout(resolve, 1000));

          for (const file of files) {
            try {
              console.log(`⬆️ Subiendo: ${file.name}...`);
              await sharePointService.uploadAttachment(newItem.id, file);
              uploadedCount++;
            } catch (fileError: unknown) {
              console.error(`❌ Error subiendo ${file.name}:`, fileError);
              const err = fileError as { response?: { data?: unknown } };
              console.error("Detalles:", err.response?.data);
              failedCount++;
            }
          }

          console.log(
            `📊 Resultado: ${uploadedCount} exitosos, ${failedCount} fallidos`
          );
        }

        await loadRealData();
        setShowForm(false);

        const message =
          files && files.length > 0
            ? `Equipo agregado exitosamente. Archivos: ${uploadedCount} subidos` +
              (failedCount > 0 ? `, ${failedCount} fallidos` : "")
            : "Equipo agregado exitosamente";

        alert(message);
      } catch (error: unknown) {
        console.error("Error adding vehicle:", error);
        const err = error as {
          response?: { data?: { error?: { message?: string } } };
          message?: string;
        };
        const errorMessage =
          err.response?.data?.error?.message ||
          err.message ||
          "Error desconocido";
        console.error("Error completo:", err.response?.data);
        alert(
          `Error al agregar equipo: ${errorMessage}. Revisa la consola para más detalles.`
        );
      }
    }
  };

  const handleEditVehicle = async (
    data: VehicleFormPayload,
    files?: File[]
  ) => {
    if (!editingVehicle) return;

    if (useMockData) {
      // Simular edición en mock
      const updatedItems = items.map((item) =>
        item.id === editingVehicle.id
          ? {
              ...item,
              fields: {
                ...data,
                FechaSolicitud: data.FechaSolicitud + "T00:00:00Z",
                FechaCompromisoComercial:
                  data.FechaCompromisoComercial + "T00:00:00Z",
                FechaInicioCiclo: data.FechaInicioCiclo + "T00:00:00Z",
                FechaFinalAlistamiento: data.FechaFinalAlistamiento
                  ? data.FechaFinalAlistamiento + "T00:00:00Z"
                  : null,
                PorcentajeAvanceTotal: item.fields.PorcentajeAvanceTotal, // Mantener calculado
                DiasRestantes: item.fields.DiasRestantes, // Mantener calculado
              },
              lastModifiedDateTime: new Date().toISOString(),
            }
          : item
      );
      setItems(updatedItems);
      setEditingVehicle(null);
      alert(
        "Equipo actualizado (modo prueba). En producción se guardará en SharePoint."
      );
    } else {
      try {
        const sharePointFields: Record<string, unknown> = {
          Title: data.Title || "",
          field_0: data.Serie || "",
          field_1: data.Prioridad ?? 0,
          field_2: data.Modelo || "",
          field_3: data.OTT || "",
          field_4: data.Asesor || "",
          field_28: data.Sede || "",
        };

        if (data.FechaSolicitud) {
          sharePointFields.field_7 = new Date(
            data.FechaSolicitud + "T00:00:00Z"
          ).toISOString();
        }
        if (data.FechaCompromisoComercial) {
          sharePointFields.field_9 = new Date(
            data.FechaCompromisoComercial + "T00:00:00Z"
          ).toISOString();
        }
        if (data.FechaInicioCiclo) {
          sharePointFields.field_10 = new Date(
            data.FechaInicioCiclo + "T00:00:00Z"
          ).toISOString();
        }
        if (data.FechaFinalAlistamiento) {
          sharePointFields.FechaFinalAlistamiento = new Date(
            data.FechaFinalAlistamiento + "T00:00:00Z"
          ).toISOString();
        }

        if (data.Observaciones) {
          sharePointFields.field_8 = data.Observaciones;
        }
        if (data.Ciclo !== undefined && data.Ciclo !== null) {
          sharePointFields.field_29 = `Ciclo ${data.Ciclo}`;
        }

        for (let i = 1; i <= 16; i++) {
          const faseKey = `F${i}` as keyof VehicleFormPayload;
          const value = data[faseKey];
          if (value) {
            sharePointFields[`field_${10 + i}`] = value;
          }
        }

        console.log(
          `📤 Enviando actualización para item ID: ${editingVehicle.id}`
        );
        console.log(`📋 Item completo:`, editingVehicle);
        console.log(`📋 Datos a enviar:`, sharePointFields);

        if (!editingVehicle.id || editingVehicle.id.trim() === "") {
          throw new Error("El ID del item no es válido");
        }

        await sharePointService.updateListItem(
          editingVehicle.id,
          sharePointFields
        );
        console.log(`✅ Equipo actualizado con ID: ${editingVehicle.id}`);

        let uploadedCount = 0;
        let failedCount = 0;
        if (files && files.length > 0) {
          console.log(`📎 Iniciando subida de ${files.length} archivo(s)...`);
          await new Promise((resolve) => setTimeout(resolve, 500));

          for (const file of files) {
            try {
              console.log(`⬆️ Subiendo: ${file.name}...`);
              await sharePointService.uploadAttachment(editingVehicle.id, file);
              uploadedCount++;
            } catch (fileError: unknown) {
              console.error(`❌ Error subiendo ${file.name}:`, fileError);
              const err = fileError as { response?: { data?: unknown } };
              console.error("Detalles:", err.response?.data);
              failedCount++;
            }
          }

          console.log(
            `📊 Resultado: ${uploadedCount} exitosos, ${failedCount} fallidos`
          );
        }

        await loadRealData();
        setEditingVehicle(null);

        const message =
          files && files.length > 0
            ? `Equipo actualizado exitosamente. Archivos: ${uploadedCount} subidos` +
              (failedCount > 0 ? `, ${failedCount} fallidos` : "")
            : "Equipo actualizado exitosamente";

        alert(message);
      } catch (error: unknown) {
        console.error("Error updating vehicle:", error);
        const err = error as {
          response?: { data?: { error?: { message?: string } } };
          message?: string;
        };
        const errorMessage =
          err.response?.data?.error?.message ||
          err.message ||
          "Error desconocido";
        console.error("Error completo:", err.response?.data);
        alert(
          `Error al actualizar equipo: ${errorMessage}. Revisa la consola para más detalles.`
        );
      }
    }
  };

  const handleDeleteVehicle = async (vehicleId: string) => {
    if (!confirm("¿Estás seguro de eliminar este equipo?")) return;

    if (useMockData) {
      // Simular eliminación en mock
      setItems(items.filter((item) => item.id !== vehicleId));
      alert(
        "Equipo eliminado (modo prueba). En producción se eliminará de SharePoint."
      );
    } else {
      // Eliminar de SharePoint
      try {
        await sharePointService.deleteListItem(vehicleId);
        await loadRealData();
        alert("Equipo eliminado exitosamente");
      } catch (error) {
        console.error("Error deleting vehicle:", error);
        alert("Error al eliminar equipo. Verifica los permisos.");
      }
    }
  };

  const viewPanelClass = (view: View) =>
    currentView === view ? "w-full" : "hidden";

  return (
    <div className="min-h-screen bg-gray-100">
      {/* Header */}
      <header className="bg-white shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <div className="flex items-center justify-between">
            <div className="flex-1 flex justify-start">
              <a
                href="https://calendarioservicios.vercel.app/login"
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors border border-gray-200"
              >
                <Calendar className="w-5 h-5 text-blue-600" />
                Calendario Servicio
              </a>
            </div>
            {/* Logo centrado */}
            <div className="flex-1 flex justify-center">
              <img
                src="https://res.cloudinary.com/dbufrzoda/image/upload/v1750457354/Captura_de_pantalla_2025-06-20_170819_wzmyli.png"
                alt="Logo Partequipos"
                className="h-16 object-contain"
              />
            </div>
            {/* Botón de login alineado a la derecha */}
            <div className="flex-1 flex justify-end">
              <LoginButton />
            </div>
          </div>
        </div>
      </header>

      {/* Navigation */}
      <div className="bg-white border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between">
            <div className="flex space-x-4">
              <button
                type="button"
                onClick={() => setCurrentView("dashboard")}
                className={`px-4 py-3 flex items-center gap-2 border-b-2 transition-colors ${
                  currentView === "dashboard"
                    ? "border-blue-600 text-blue-600 font-medium"
                    : "border-transparent text-gray-600 hover:text-gray-900"
                }`}
              >
                <LayoutDashboard className="w-5 h-5" />
                Dashboard
              </button>
              <button
                type="button"
                onClick={() => setCurrentView("table")}
                className={`px-4 py-3 flex items-center gap-2 border-b-2 transition-colors ${
                  currentView === "table"
                    ? "border-blue-600 text-blue-600 font-medium"
                    : "border-transparent text-gray-600 hover:text-gray-900"
                }`}
              >
                <Table2 className="w-5 h-5" />
                Tabla de Datos
              </button>
              <button
                type="button"
                onClick={() => setCurrentView("ciclo-mes")}
                className={`px-4 py-3 flex items-center gap-2 border-b-2 transition-colors ${
                  currentView === "ciclo-mes"
                    ? "border-blue-600 text-blue-600 font-medium"
                    : "border-transparent text-gray-600 hover:text-gray-900"
                }`}
              >
                <Monitor className="w-5 h-5" />
                Visualización Ciclo-Mes
              </button>
            </div>
            <button
              type="button"
              onClick={() => setShowForm(true)}
              className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors flex items-center gap-2"
            >
              <Plus className="w-5 h-5" />
              Agregar Equipo
            </button>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <main className="w-full px-4 sm:px-6 lg:px-8 py-8">
        {isLoading ? (
          <div className="flex flex-col items-center justify-center p-12">
            <Loader2 className="w-12 h-12 text-blue-600 animate-spin mb-4" />
            <p className="text-gray-600">Cargando datos...</p>
          </div>
        ) : (
          <>
            <div className={viewPanelClass("dashboard")}>
              <div className="w-full mb-6">
                <DashboardFilters
                  items={items}
                  filters={filters}
                  onFilterChange={setFilters}
                />
              </div>
              <DashboardReal items={filteredItems} />
            </div>
            <div className={viewPanelClass("table")}>
              <SharePointTableReal
                items={filteredItems}
                useMockData={useMockData}
                onEdit={setEditingVehicle}
                onDelete={handleDeleteVehicle}
                onRefresh={useMockData ? loadMockData : loadRealData}
              />
            </div>
            <div className={viewPanelClass("ciclo-mes")}>
              <CicloMesView items={items} />
            </div>
          </>
        )}
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-gray-200 mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <p className="text-center text-sm text-gray-500">
            © 2025 Seguimiento Usados - Desarrollado con React + Vite +
            TypeScript
          </p>
        </div>
      </footer>

      {/* Modal de Formulario */}
      {(showForm || editingVehicle) && (
        <VehicleFormReal
          vehicle={editingVehicle || undefined}
          onSubmit={editingVehicle ? handleEditVehicle : handleAddVehicle}
          onCancel={() => {
            setShowForm(false);
            setEditingVehicle(null);
          }}
          isEditing={!!editingVehicle}
        />
      )}
    </div>
  );
}

function App() {
  return (
    <AuthProvider>
      <AppContent />
    </AuthProvider>
  );
}

export default App;
