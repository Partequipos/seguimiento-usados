/**
 * Formulario para agregar/editar equipos
 */

import React, { useState } from "react";
import { useForm } from "react-hook-form";
import { X, Save, Loader2 } from "lucide-react";
import { SharePointListItem } from "../services/sharePointService";

interface VehicleFormProps {
  vehicle?: SharePointListItem;
  onSubmit: (data: any) => Promise<void>;
  onCancel: () => void;
  isEditing?: boolean;
}

interface FormData {
  Title: string;
  Marca: string;
  Modelo: string;
  Año: number;
  Estado: string;
  TipoVehiculo: string;
  FechaIngreso: string;
  FechaEstimadaEntrega: string;
  Kilometraje: number;
  PrecioVenta: number;
  Responsable: string;
  Observaciones: string;
  Prioridad: string;
  PorcentajeAvance: number;
}

const VehicleForm: React.FC<VehicleFormProps> = ({
  vehicle,
  onSubmit,
  onCancel,
  isEditing = false,
}) => {
  const [isSubmitting, setIsSubmitting] = useState(false);

  const {
    register,
    handleSubmit,
    formState: { errors },
  } = useForm<FormData>({
    defaultValues: vehicle
      ? {
          Title: vehicle.fields.Title,
          Marca: vehicle.fields.Marca,
          Modelo: vehicle.fields.Modelo,
          Año: vehicle.fields.Año,
          Estado: vehicle.fields.Estado,
          TipoVehiculo: vehicle.fields.TipoVehiculo,
          FechaIngreso: vehicle.fields.FechaIngreso?.split("T")[0] || "",
          FechaEstimadaEntrega:
            vehicle.fields.FechaEstimadaEntrega?.split("T")[0] || "",
          Kilometraje: vehicle.fields.Kilometraje,
          PrecioVenta: vehicle.fields.PrecioVenta,
          Responsable: vehicle.fields.Responsable,
          Observaciones: vehicle.fields.Observaciones,
          Prioridad: vehicle.fields.Prioridad,
          PorcentajeAvance: vehicle.fields.PorcentajeAvance,
        }
      : {
          Estado: "Pendiente",
          TipoVehiculo: "Automóvil",
          Prioridad: "Media",
          PorcentajeAvance: 0,
        },
  });

  const onSubmitForm = async (data: FormData) => {
    setIsSubmitting(true);
    try {
      await onSubmit(data);
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full max-h-[90vh] overflow-y-auto">
        {/* Header */}
        <div className="flex items-center justify-between p-6 border-b border-gray-200">
          <h2 className="text-2xl font-bold text-gray-900">
            {isEditing ? "Editar Equipo" : "Agregar Nuevo Equipo"}
          </h2>
          <button
            onClick={onCancel}
            className="text-gray-400 hover:text-gray-600 transition-colors"
          >
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit(onSubmitForm)} className="p-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Placa */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Placa <span className="text-red-500">*</span>
              </label>
              <input
                {...register("Title", { required: "La placa es obligatoria" })}
                type="text"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="ABC123"
              />
              {errors.Title && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.Title.message}
                </p>
              )}
            </div>

            {/* Marca */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Marca <span className="text-red-500">*</span>
              </label>
              <input
                {...register("Marca", { required: "La marca es obligatoria" })}
                type="text"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="Toyota"
              />
              {errors.Marca && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.Marca.message}
                </p>
              )}
            </div>

            {/* Modelo */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Modelo <span className="text-red-500">*</span>
              </label>
              <input
                {...register("Modelo", {
                  required: "El modelo es obligatorio",
                })}
                type="text"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="Hilux"
              />
              {errors.Modelo && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.Modelo.message}
                </p>
              )}
            </div>

            {/* Año */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Año <span className="text-red-500">*</span>
              </label>
              <input
                {...register("Año", {
                  required: "El año es obligatorio",
                  min: { value: 1990, message: "Año mínimo: 1990" },
                  max: { value: 2030, message: "Año máximo: 2030" },
                })}
                type="number"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="2020"
              />
              {errors.Año && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.Año.message}
                </p>
              )}
            </div>

            {/* Tipo de Equipo */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Tipo de Equipo <span className="text-red-500">*</span>
              </label>
              <select
                {...register("TipoVehiculo", {
                  required: "El tipo es obligatorio",
                })}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
              >
                <option value="Automóvil">Automóvil</option>
                <option value="Camioneta">Camioneta</option>
                <option value="SUV">SUV</option>
                <option value="Camión">Camión</option>
                <option value="Motocicleta">Motocicleta</option>
              </select>
            </div>

            {/* Estado */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Estado <span className="text-red-500">*</span>
              </label>
              <select
                {...register("Estado", {
                  required: "El estado es obligatorio",
                })}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
              >
                <option value="Pendiente">Pendiente</option>
                <option value="En Proceso">En Proceso</option>
                <option value="Completado">Completado</option>
              </select>
            </div>

            {/* Kilometraje */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Kilometraje <span className="text-red-500">*</span>
              </label>
              <input
                {...register("Kilometraje", {
                  required: "El kilometraje es obligatorio",
                  min: {
                    value: 0,
                    message: "El kilometraje debe ser positivo",
                  },
                })}
                type="number"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="50000"
              />
              {errors.Kilometraje && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.Kilometraje.message}
                </p>
              )}
            </div>

            {/* Precio de Venta */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Precio de Venta (COP) <span className="text-red-500">*</span>
              </label>
              <input
                {...register("PrecioVenta", {
                  required: "El precio es obligatorio",
                  min: { value: 0, message: "El precio debe ser positivo" },
                })}
                type="number"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="50000000"
              />
              {errors.PrecioVenta && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.PrecioVenta.message}
                </p>
              )}
            </div>

            {/* Fecha de Ingreso */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Fecha de Ingreso <span className="text-red-500">*</span>
              </label>
              <input
                {...register("FechaIngreso", {
                  required: "La fecha es obligatoria",
                })}
                type="date"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
              />
              {errors.FechaIngreso && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.FechaIngreso.message}
                </p>
              )}
            </div>

            {/* Fecha Estimada de Entrega */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Fecha Estimada Entrega <span className="text-red-500">*</span>
              </label>
              <input
                {...register("FechaEstimadaEntrega", {
                  required: "La fecha es obligatoria",
                })}
                type="date"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
              />
              {errors.FechaEstimadaEntrega && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.FechaEstimadaEntrega.message}
                </p>
              )}
            </div>

            {/* Responsable */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Responsable <span className="text-red-500">*</span>
              </label>
              <input
                {...register("Responsable", {
                  required: "El responsable es obligatorio",
                })}
                type="text"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="Juan Pérez"
              />
              {errors.Responsable && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.Responsable.message}
                </p>
              )}
            </div>

            {/* Prioridad */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Prioridad <span className="text-red-500">*</span>
              </label>
              <select
                {...register("Prioridad", {
                  required: "La prioridad es obligatoria",
                })}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
              >
                <option value="Baja">Baja</option>
                <option value="Media">Media</option>
                <option value="Alta">Alta</option>
              </select>
            </div>

            {/* Porcentaje de Avance */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                % Avance <span className="text-red-500">*</span>
              </label>
              <input
                {...register("PorcentajeAvance", {
                  required: "El porcentaje es obligatorio",
                  min: { value: 0, message: "Mínimo 0%" },
                  max: { value: 100, message: "Máximo 100%" },
                })}
                type="number"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent"
                placeholder="0"
              />
              {errors.PorcentajeAvance && (
                <p className="mt-1 text-sm text-red-600">
                  {errors.PorcentajeAvance.message}
                </p>
              )}
            </div>

            {/* Observaciones - Ancho completo */}
            <div className="md:col-span-2">
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Observaciones
              </label>
              <textarea
                {...register("Observaciones")}
                rows={4}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-red/40 focus:border-transparent resize-none"
                placeholder="Observaciones adicionales sobre el equipo..."
              />
            </div>
          </div>

          {/* Footer */}
          <div className="flex items-center justify-end gap-4 mt-8 pt-6 border-t border-gray-200">
            <button
              type="button"
              onClick={onCancel}
              disabled={isSubmitting}
              className="px-6 py-2 border border-gray-300 text-gray-700 rounded-lg hover:bg-gray-50 transition-colors disabled:opacity-50"
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

export default VehicleForm;
