/**
 * Formulario para agregar/editar equipos
 * Basado en la estructura real de SharePoint
 */

import React, { useState } from "react";
import { useForm, useWatch } from "react-hook-form";
import { X, Save, Loader2, Upload, FileIcon, Trash2 } from "lucide-react";
import { SharePointListItem } from "../services/sharePointService";

interface VehicleFormProps {
  vehicle?: SharePointListItem;
  onSubmit: (data: any, files?: File[]) => Promise<void>;
  onCancel: () => void;
  isEditing?: boolean;
}

interface FormData {
  Title: string;
  Serie: string;
  Prioridad: number;
  Modelo: string;
  OTT: string;
  Asesor: string;
  FechaSolicitud: string;
  Observaciones: string;
  FechaCompromisoComercial: string;
  FechaInicioCiclo: string;
  F1: string;
  F2: string;
  F3: string;
  F4: string;
  F5: string;
  F6: string;
  F7: string;
  F8: string;
  F9: string;
  F10: string;
  F11: string;
  F12: string;
  F13: string;
  F14: string;
  F15: string;
  F16: string;
  Sede: string;
  Ciclo: number;
  FechaFinalAlistamiento: string;
}

const VehicleFormReal: React.FC<VehicleFormProps> = ({
  vehicle,
  onSubmit,
  onCancel,
  isEditing = false,
}) => {
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [currentTab, setCurrentTab] = useState<"general" | "fases">("general");
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);

  const {
    register,
    handleSubmit,
    formState: { errors },
    control,
  } = useForm<FormData>({
    defaultValues: vehicle
      ? {
          Title: vehicle.fields.Title,
          Serie: vehicle.fields.Serie,
          Prioridad: vehicle.fields.Prioridad,
          Modelo: vehicle.fields.Modelo,
          OTT: vehicle.fields.OTT,
          Asesor: vehicle.fields.Asesor,
          FechaSolicitud: vehicle.fields.FechaSolicitud?.split("T")[0] || "",
          Observaciones: vehicle.fields.Observaciones,
          FechaCompromisoComercial:
            vehicle.fields.FechaCompromisoComercial?.split("T")[0] || "",
          FechaInicioCiclo:
            vehicle.fields.FechaInicioCiclo?.split("T")[0] || "",
          F1: vehicle.fields.F1 || "0%",
          F2: vehicle.fields.F2 || "0%",
          F3: vehicle.fields.F3 || "0%",
          F4: vehicle.fields.F4 || "0%",
          F5: vehicle.fields.F5 || "0%",
          F6: vehicle.fields.F6 || "0%",
          F7: vehicle.fields.F7 || "0%",
          F8: vehicle.fields.F8 || "0%",
          F9: vehicle.fields.F9 || "0%",
          F10: vehicle.fields.F10 || "0%",
          F11: vehicle.fields.F11 || "0%",
          F12: vehicle.fields.F12 || "0%",
          F13: vehicle.fields.F13 || "0%",
          F14: vehicle.fields.F14 || "0%",
          F15: vehicle.fields.F15 || "0%",
          F16: vehicle.fields.F16 || "0%",
          Sede: vehicle.fields.Sede,
          Ciclo: vehicle.fields.Ciclo ? (typeof vehicle.fields.Ciclo === 'string' ? Number(vehicle.fields.Ciclo.replace('Ciclo ', '')) || 1 : Number(vehicle.fields.Ciclo) || 1) : 1,
          FechaFinalAlistamiento:
            vehicle.fields.FechaFinalAlistamiento?.split("T")[0] || "",
        }
      : {
          Prioridad: 0,
          F1: "0%",
          F2: "0%",
          F3: "0%",
          F4: "0%",
          F5: "0%",
          F6: "0%",
          F7: "0%",
          F8: "0%",
          F9: "0%",
          F10: "0%",
          F11: "0%",
          F12: "0%",
          F13: "0%",
          F14: "0%",
          F15: "0%",
          F16: "0%",
          Ciclo: "1",
        },
  });

  const onSubmitForm = async (data: FormData) => {
    setIsSubmitting(true);
    try {
      await onSubmit(data, selectedFiles);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const files = Array.from(e.target.files);
      setSelectedFiles((prev) => [...prev, ...files]);
    }
  };

  const removeFile = (index: number) => {
    setSelectedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  // Observar valores del formulario para mostrar en el header
  const watchedSerie = useWatch({ control, name: "Serie" });
  const watchedModelo = useWatch({ control, name: "Modelo" });
  const watchedAsesor = useWatch({ control, name: "Asesor" });

  const faseOptions = ["0%", "25%", "50%", "75%", "100%"];

  // Nombres descriptivos de las fases
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

  // Modelos reales de maquinaria de Partequipos
  const modeloOptions = [
    "ZX75US-3",
    "ZX200-3",
    "ZX120-3",
    "ZX75US-A",
    "ZX75USK-3",
    "SD45D",
    "ZX40U-5A",
    "ZX135US-3",
    "ZX200-5G",
    "D6M",
    "ZX120-5B",
    "SH200-5",
    "ZX140H",
    "SH75",
    "ZX70",
    "ZX 70-3",
    "ZX225US-3",
    "VIO70-3",
    "K120",
    "ZX210-5G",
    "ZX135US-5B",
    "ZX17U-2",
    "ZX130-5G",
    "K70-3",
    "ZX50U-3",
    "ZX40U-3",
    "580N",
    "SH120-5",
    "210X3E",
    "VIO17-1B",
    "130X3",
    "ZX30U-5A",
    "SR200B",
    "SH210-5",
    "ZX75US-5B",
    "ZX200-5B",
    "ZX225US-5B",
    "ZX210K-5B",
    "ZX40U-5B",
    "D3C",
    "ZX200-6",
    "ZX135US-6",
    "ZX75USK-5B",
    "ZX200X -5B",
    "856H",
    "ZX130L-5B",
    "ASC70D",
    "ZX120-6",
    "ZX50U-5B",
    "ZX130K-6",
    "ZX225USR-6",
    "ZX300-6N",
  ];

  // Asesores reales de Partequipos
  const asesorOptions = [
    "Juan Guerrero",
    "Claudia Bogotá",
    "Juan Suarez",
    "Luis Anchundia",
    "Felipe Moreno",
    "Claudia Echeverri",
    "Edgar Bustos",
    "Fabian Corrales",
    "Efrain Rua",
    "Cesar Zapata",
    "German Blanco",
    "Diego Rojas",
    "Gerencia",
    "Juliana Florez",
    "Jhon Ussa",
    "Leonardo Cardona",
    "Diana Ardila",
    "Yudi Ochoa",
    "Fabiola Hurtado",
    "Julian Ramirez",
    "Libre",
    "Ludwing Martínez",
  ];

  // Sedes de Partequipos
  const sedeOptions = [
    "Bogotá",
    "Medellín",
    "Cali",
    "Barranquilla",
    "Cartagena",
    "Guarne",
  ];

  // Observaciones reales
  const observacionesOptions = [
    "Ok dinero y OC",
    "Entregada al cliente",
    "Libre",
    "Lista, Pendiente Entrega",
    "Sin alistamiento",
    "Reservada",
  ];

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-6xl w-full my-8">
        {/* Header */}
        <div className="flex items-center justify-between p-6 border-b border-gray-200 bg-gradient-to-r from-brand-red-soft to-white">
          <div className="flex-1">
            <h2 className="text-2xl font-bold text-gray-900 mb-2">
              {isEditing ? "Editar Equipo" : "Agregar Nuevo Equipo"}
            </h2>
            {/* Información del equipo visible en modo edición */}
            {isEditing && (watchedSerie || watchedModelo || watchedAsesor) && (
              <div className="flex flex-wrap items-center gap-3 mt-3">
                {watchedSerie && (
                  <div className="inline-flex items-center gap-2 px-3 py-1.5 bg-brand-red-soft text-brand-red-dark rounded-lg border border-brand-red-muted">
                    <span className="text-xs font-semibold uppercase tracking-wide">Serie:</span>
                    <span className="font-bold">{watchedSerie}</span>
                  </div>
                )}
                {watchedModelo && (
                  <div className="inline-flex items-center gap-2 px-3 py-1.5 bg-green-100 text-green-800 rounded-lg border border-green-200">
                    <span className="text-xs font-semibold uppercase tracking-wide">Modelo:</span>
                    <span className="font-bold">{watchedModelo}</span>
                  </div>
                )}
                {watchedAsesor && (
                  <div className="inline-flex items-center gap-2 px-3 py-1.5 bg-purple-100 text-purple-800 rounded-lg border border-purple-200">
                    <span className="text-xs font-semibold uppercase tracking-wide">Asesor:</span>
                    <span className="font-bold">{watchedAsesor}</span>
                  </div>
                )}
              </div>
            )}
          </div>
          <button
            onClick={onCancel}
            className="text-gray-400 hover:text-gray-600 transition-colors ml-4"
          >
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Tabs */}
        <div className="border-b border-gray-200">
          <div className="flex px-6">
            <button
              type="button"
              onClick={() => setCurrentTab("general")}
              className={`px-4 py-3 border-b-2 transition-colors ${
                currentTab === "general"
                  ? "border-brand-red text-brand-red font-medium"
                  : "border-transparent text-gray-600 hover:text-gray-900"
              }`}
            >
              Información General
            </button>
            <button
              type="button"
              onClick={() => setCurrentTab("fases")}
              className={`px-4 py-3 border-b-2 transition-colors ${
                currentTab === "fases"
                  ? "border-brand-red text-brand-red font-medium"
                  : "border-transparent text-gray-600 hover:text-gray-900"
              }`}
            >
              Fases del Alistamiento (F1-F16)
            </button>
          </div>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit(onSubmitForm)}>
          <div className="p-6 max-h-[60vh] overflow-y-auto">
            {currentTab === "general" ? (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Título */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Cliente <span className="text-red-500">*</span>
                  </label>
                  <input
                    {...register("Title", {
                      required: "El cliente es obligatorio",
                    })}
                    type="text"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                    placeholder="Ingrese el nombre del cliente"
                  />
                  {errors.Title && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.Title.message}
                    </p>
                  )}
                </div>

                {/* Serie */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Serie <span className="text-red-500">*</span>
                  </label>
                  <input
                    {...register("Serie", {
                      required: "La serie es obligatoria",
                    })}
                    type="text"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                    placeholder="2024-001"
                  />
                  {errors.Serie && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.Serie.message}
                    </p>
                  )}
                </div>

                {/* Modelo */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Modelo <span className="text-red-500">*</span>
                  </label>
                  <select
                    {...register("Modelo", {
                      required: "El modelo es obligatorio",
                    })}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                  >
                    <option value="">Seleccione un modelo</option>
                    {modeloOptions.map((modelo) => (
                      <option key={modelo} value={modelo}>
                        {modelo}
                      </option>
                    ))}
                  </select>
                  {errors.Modelo && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.Modelo.message}
                    </p>
                  )}
                </div>

                {/* OTT */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    OTT <span className="text-red-500">*</span>
                  </label>
                  <input
                    {...register("OTT", { required: "El OTT es obligatorio" })}
                    type="text"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                    placeholder="OTT-2024-001"
                  />
                  {errors.OTT && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.OTT.message}
                    </p>
                  )}
                </div>

                {/* Asesor */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Asesor <span className="text-red-500">*</span>
                  </label>
                  <select
                    {...register("Asesor", {
                      required: "El asesor es obligatorio",
                    })}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                  >
                    <option value="">Seleccione un asesor</option>
                    {asesorOptions.map((asesor) => (
                      <option key={asesor} value={asesor}>
                        {asesor}
                      </option>
                    ))}
                  </select>
                  {errors.Asesor && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.Asesor.message}
                    </p>
                  )}
                </div>

                {/* Prioridad */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Prioridad <span className="text-red-500">*</span>
                  </label>
                  <input
                    {...register("Prioridad", {
                      required: "La prioridad es obligatoria",
                      valueAsNumber: true,
                    })}
                    type="number"
                    step="1"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                    placeholder="Ingrese un número"
                  />
                  {errors.Prioridad && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.Prioridad.message}
                    </p>
                  )}
                </div>

                {/* Sede */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Sede <span className="text-red-500">*</span>
                  </label>
                  <select
                    {...register("Sede", {
                      required: "La sede es obligatoria",
                    })}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                  >
                    <option value="">Seleccione una sede</option>
                    {sedeOptions.map((sede) => (
                      <option key={sede} value={sede}>
                        {sede}
                      </option>
                    ))}
                  </select>
                  {errors.Sede && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.Sede.message}
                    </p>
                  )}
                </div>

                {/* Ciclo */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Ciclo <span className="text-red-500">*</span>
                  </label>
                  <input
                    {...register("Ciclo", {
                      required: "El ciclo es obligatorio",
                      valueAsNumber: true,
                      min: { value: 1, message: "El ciclo debe ser mayor a 0" },
                    })}
                    type="number"
                    min="1"
                    step="1"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                    placeholder="Ingrese el número de ciclo"
                  />
                  {errors.Ciclo && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.Ciclo.message}
                    </p>
                  )}
                </div>

                {/* Fecha de Solicitud */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Fecha de Solicitud
                  </label>
                  <input
                    {...register("FechaSolicitud")}
                    type="date"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                  />
                  {errors.FechaSolicitud && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.FechaSolicitud.message}
                    </p>
                  )}
                </div>

                {/* Fecha de Compromiso Comercial */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Fecha de Compromiso Comercial
                  </label>
                  <input
                    {...register("FechaCompromisoComercial")}
                    type="date"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                  />
                  {errors.FechaCompromisoComercial && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.FechaCompromisoComercial.message}
                    </p>
                  )}
                </div>

                {/* Fecha de Inicio de Ciclo */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Fecha de Inicio de Ciclo{" "}
                    <span className="text-red-500">*</span>
                  </label>
                  <input
                    {...register("FechaInicioCiclo", {
                      required: "La fecha es obligatoria",
                    })}
                    type="date"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                  />
                  {errors.FechaInicioCiclo && (
                    <p className="mt-1 text-sm text-red-600">
                      {errors.FechaInicioCiclo.message}
                    </p>
                  )}
                </div>

                {/* Observaciones */}
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Observaciones
                  </label>
                  <select
                    {...register("Observaciones")}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                  >
                    <option value="">Seleccione una observación</option>
                    {observacionesOptions.map((obs) => (
                      <option key={obs} value={obs}>
                        {obs}
                      </option>
                    ))}
                  </select>
                </div>

                {/* Adjuntos / Archivos */}
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    📎 Adjuntos (Fotos/Archivos)
                  </label>
                  <div className="border-2 border-dashed border-gray-300 rounded-lg p-4 hover:border-brand-red-muted transition-colors">
                    <div className="flex items-center justify-center">
                      <label className="cursor-pointer flex flex-col items-center">
                        <Upload className="w-8 h-8 text-gray-400 mb-2" />
                        <span className="text-sm text-gray-600">
                          Haz clic para seleccionar archivos
                        </span>
                        <span className="text-xs text-gray-400 mt-1">
                          Imágenes, PDFs, documentos (máx. 10MB por archivo)
                        </span>
                        <input
                          type="file"
                          multiple
                          accept="image/*,.pdf,.doc,.docx,.xls,.xlsx"
                          onChange={handleFileChange}
                          className="hidden"
                        />
                      </label>
                    </div>
                  </div>

                  {/* Lista de archivos seleccionados */}
                  {selectedFiles.length > 0 && (
                    <div className="mt-3 space-y-2">
                      <p className="text-sm font-medium text-gray-700">
                        Archivos seleccionados ({selectedFiles.length}):
                      </p>
                      {selectedFiles.map((file, index) => (
                        <div
                          key={index}
                          className="flex items-center justify-between p-2 bg-brand-red-soft rounded-lg border border-brand-red-muted"
                        >
                          <div className="flex items-center gap-2 flex-1 min-w-0">
                            <FileIcon className="w-4 h-4 text-brand-red flex-shrink-0" />
                            <span className="text-sm text-gray-900 truncate">
                              {file.name}
                            </span>
                            <span className="text-xs text-gray-500">
                              ({(file.size / 1024).toFixed(0)} KB)
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => removeFile(index)}
                            className="ml-2 p-1 text-red-600 hover:bg-red-100 rounded transition-colors"
                            title="Eliminar archivo"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            ) : (
              /* Pestaña de Fases */
              <div>
                <div className="mb-6 p-4 bg-brand-red-soft rounded-lg border-l-4 border-blue-500">
                  <p className="text-sm text-blue-800">
                    <strong>📋 Instrucciones:</strong> Selecciona el porcentaje de cada
                    fase del alistamiento (0%, 25%, 50%, 75%, 100%). El % de avance total se calculará
                    automáticamente en SharePoint.
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  {Array.from({ length: 16 }, (_, i) => i + 1).map((num) => (
                    <div key={`F${num}`} className="bg-gray-50 p-3 rounded-lg border border-gray-200 hover:border-blue-300 transition-colors">
                      <label className="block text-sm font-semibold text-gray-800 mb-1">
                        F{num}
                      </label>
                      <p className="text-[10px] text-gray-500 mb-2 leading-tight uppercase">
                        {faseNombres[num]}
                      </p>
                      <select
                        {...register(`F${num}` as any)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent text-sm bg-white"
                      >
                        {faseOptions.map((option) => (
                          <option key={option} value={option}>
                            {option}
                          </option>
                        ))}
                      </select>
                    </div>
                  ))}
                </div>

                {/* Aviso y Campo de Fecha Final Alistamiento */}
                <div className="mt-6 grid grid-cols-1 md:grid-cols-2 gap-4 items-start">
                  {/* Aviso en rojo */}
                  <div className="p-4 bg-red-50 border-2 border-red-300 rounded-lg">
                    <p className="text-sm font-semibold text-red-700">
                      ⚠️ Asignar fecha final de alistamiento al quedar todo 100%
                    </p>
                  </div>
                  
                  {/* Campo Fecha Final Alistamiento */}
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Fecha Final Alistamiento
                    </label>
                    <input
                      {...register("FechaFinalAlistamiento")}
                      type="date"
                      className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                    />
                  </div>
                </div>

                <div className="mt-6 p-4 bg-gray-50 rounded-lg">
                  <h4 className="font-medium text-gray-900 mb-2">
                    Campos Calculados Automáticamente:
                  </h4>
                  <ul className="text-sm text-gray-600 space-y-1">
                    <li>
                      • <strong>% avance total:</strong> Se calcula basándose en
                      el estado de F1-F16
                    </li>
                    <li>
                      • <strong>Días faltantes:</strong> Se calcula
                      automáticamente según las fases
                    </li>
                  </ul>
                </div>
              </div>
            )}
          </div>

          {/* Footer */}
          <div className="flex items-center justify-end gap-4 px-6 py-4 border-t border-gray-200 bg-gray-50">
            <button
              type="button"
              onClick={onCancel}
              disabled={isSubmitting}
              className="px-6 py-2 border border-gray-300 text-gray-700 rounded-lg hover:bg-white transition-colors disabled:opacity-50"
            >
              Cancelar
            </button>
            <button
              type="submit"
              disabled={isSubmitting}
              className="px-6 py-2 bg-brand-red text-white rounded-lg hover:bg-brand-red-dark transition-colors disabled:opacity-50 flex items-center gap-2"
            >
              {isSubmitting ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" />
                  Guardando...
                </>
              ) : (
                <>
                  <Save className="w-5 h-5" />
                  {isEditing ? "Actualizar" : "Guardar"}
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default VehicleFormReal;
