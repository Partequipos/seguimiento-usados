/**
 * Dashboard con estadísticas y gráficos
 */

import React from "react";
import {
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from "recharts";
import {
  TrendingUp,
  Clock,
  CheckCircle,
  AlertCircle,
  DollarSign,
  Car,
} from "lucide-react";
import { SharePointListItem } from "../services/sharePointService";

interface DashboardProps {
  items: SharePointListItem[];
}

const Dashboard: React.FC<DashboardProps> = ({ items }) => {
  // Calcular estadísticas
  const stats = {
    total: items.length,
    enProceso: items.filter((i) => i.fields.Estado === "En Proceso").length,
    completados: items.filter((i) => i.fields.Estado === "Completado").length,
    pendientes: items.filter((i) => i.fields.Estado === "Pendiente").length,
    precioTotal: items.reduce(
      (sum, i) => sum + (Number(i.fields.PrecioVenta) || 0),
      0
    ),
    promedioAvance:
      items.reduce(
        (sum, i) => sum + (Number(i.fields.PorcentajeAvance) || 0),
        0
      ) / items.length || 0,
  };

  // Datos para gráfico de estados
  const estadosData = [
    { name: "En Proceso", value: stats.enProceso, color: "#cf1b22" },
    { name: "Completado", value: stats.completados, color: "#10b981" },
    { name: "Pendiente", value: stats.pendientes, color: "#f59e0b" },
  ];

  // Datos para gráfico de equipos por responsable
  const equiposPorResponsable = items.reduce(
    (acc: Record<string, number>, item) => {
      const responsable =
        (item.fields.Responsable as string) || "Sin asignar";
      acc[responsable] = (acc[responsable] ?? 0) + 1;
      return acc;
    },
    {} as Record<string, number>
  );

  const responsablesData = Object.entries(equiposPorResponsable).map(
    ([name, value]) => ({
      name,
      cantidad: value,
    })
  );

  // Datos para gráfico de prioridad
  const prioridadData = items.reduce(
    (acc: Record<string, number>, item) => {
      const prioridad =
        String(item.fields.Prioridad ?? "") || "Sin prioridad";
      acc[prioridad] = (acc[prioridad] ?? 0) + 1;
      return acc;
    },
    {} as Record<string, number>
  );

  const getPrioridadColor = (name: string): string => {
    if (name === "Alta") return "#ef4444";
    if (name === "Media") return "#f59e0b";
    return "#10b981";
  };

  const prioridadChartData = Object.entries(prioridadData).map(
    ([name, value]) => ({
      name,
      value,
      color: getPrioridadColor(name),
    })
  );

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat("es-CO", {
      style: "currency",
      currency: "COP",
      minimumFractionDigits: 0,
    }).format(value);
  };

  return (
    <div className="space-y-6">
      {/* Tarjetas de estadísticas */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        <div className="bg-white rounded-lg shadow p-6">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-gray-600">
                Total Equipos
              </p>
              <p className="text-3xl font-bold text-gray-900 mt-2">
                {stats.total}
              </p>
            </div>
            <div className="bg-brand-red-soft rounded-full p-3">
              <Car className="w-6 h-6 text-brand-red" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow p-6">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-gray-600">En Proceso</p>
              <p className="text-3xl font-bold text-brand-red mt-2">
                {stats.enProceso}
              </p>
            </div>
            <div className="bg-brand-red-soft rounded-full p-3">
              <Clock className="w-6 h-6 text-brand-red" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow p-6">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-gray-600">Completados</p>
              <p className="text-3xl font-bold text-green-600 mt-2">
                {stats.completados}
              </p>
            </div>
            <div className="bg-green-100 rounded-full p-3">
              <CheckCircle className="w-6 h-6 text-green-600" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow p-6">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-gray-600">Pendientes</p>
              <p className="text-3xl font-bold text-yellow-600 mt-2">
                {stats.pendientes}
              </p>
            </div>
            <div className="bg-yellow-100 rounded-full p-3">
              <AlertCircle className="w-6 h-6 text-yellow-600" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow p-6">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-gray-600">Valor Total</p>
              <p className="text-2xl font-bold text-gray-900 mt-2">
                {formatCurrency(stats.precioTotal)}
              </p>
            </div>
            <div className="bg-green-100 rounded-full p-3">
              <DollarSign className="w-6 h-6 text-green-600" />
            </div>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow p-6">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-gray-600">
                Avance Promedio
              </p>
              <p className="text-3xl font-bold text-brand-red mt-2">
                {stats.promedioAvance.toFixed(0)}%
              </p>
            </div>
            <div className="bg-brand-red-soft rounded-full p-3">
              <TrendingUp className="w-6 h-6 text-brand-red" />
            </div>
          </div>
        </div>
      </div>

      {/* Gráficos */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Gráfico de Estados */}
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">
            Distribución por Estado
          </h3>
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie
                data={estadosData}
                cx="50%"
                cy="50%"
                labelLine={false}
                label={({
                  name,
                  percent,
                }: {
                  name?: string;
                  percent?: number;
                }) =>
                  `${name ?? ""}: ${((percent ?? 0) * 100).toFixed(0)}%`
                }
                outerRadius={100}
                fill="#8884d8"
                dataKey="value"
              >
                {estadosData.map((entry) => (
                  <Cell key={`cell-estado-${entry.name}`} fill={entry.color} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>

        {/* Gráfico de Responsables */}
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">
            Equipos por Responsable
          </h3>
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={responsablesData}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" angle={-45} textAnchor="end" height={100} />
              <YAxis />
              <Tooltip />
              <Bar dataKey="cantidad" fill="#cf1b22" />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Gráfico de Prioridad */}
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">
            Distribución por Prioridad
          </h3>
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie
                data={prioridadChartData}
                cx="50%"
                cy="50%"
                labelLine={false}
                label={({
                  name,
                  percent,
                }: {
                  name?: string;
                  percent?: number;
                }) =>
                  `${name ?? ""}: ${((percent ?? 0) * 100).toFixed(0)}%`
                }
                outerRadius={100}
                fill="#8884d8"
                dataKey="value"
              >
                {prioridadChartData.map((entry) => (
                  <Cell
                    key={`cell-prioridad-${entry.name}`}
                    fill={entry.color}
                  />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>

        {/* Tabla de resumen */}
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">
            Resumen de Estados
          </h3>
          <div className="space-y-4">
            {estadosData.map((estado) => (
              <div
                key={estado.name}
                className="flex items-center justify-between"
              >
                <div className="flex items-center gap-3">
                  <div
                    className="w-4 h-4 rounded"
                    style={{ backgroundColor: estado.color }}
                  />
                  <span className="text-gray-700">{estado.name}</span>
                </div>
                <div className="flex items-center gap-4">
                  <span className="text-2xl font-bold text-gray-900">
                    {estado.value}
                  </span>
                  <span className="text-sm text-gray-500">
                    ({((estado.value / stats.total) * 100).toFixed(0)}%)
                  </span>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
