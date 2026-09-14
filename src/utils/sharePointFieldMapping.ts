/**
 * Utilidad para mapear nombres internos de SharePoint a nombres amigables
 * y viceversa
 */

// Mapeo de nombres internos de SharePoint -> nombres amigables
export const SharePointFieldMapping: Record<string, string> = {
  // Campos estándar
  Title: "Title",
  
  // Campos numéricos (field_X)
  field_0: "Serie",
  field_1: "Prioridad",
  field_2: "Modelo",
  field_3: "OTT",
  field_4: "Asesor",
  field_7: "FechaSolicitud",
  field_8: "Observaciones",
  field_9: "FechaCompromisoComercial",
  field_10: "FechaInicioCiclo",
  
  // Fases del alistamiento (F1-F16)
  field_11: "F1",
  field_12: "F2",
  field_13: "F3",
  field_14: "F4",
  field_15: "F5",
  field_16: "F6",
  field_17: "F7",
  field_18: "F8",
  field_19: "F9",
  field_20: "F10",
  field_21: "F11",
  field_22: "F12",
  field_23: "F13",
  field_24: "F14",
  field_25: "F15",
  field_26: "F16",
  
  // Otros campos
  field_28: "Sede",
  field_29: "Ciclo",
  
  // Campos calculados con codificación Unicode
  // El campo de porcentaje puede venir con o sin guión bajo inicial
  "_x0025__x0020_avance_x0020_total": "PorcentajeAvanceTotal",
  "x0025__x0020_avance_x0020_total": "PorcentajeAvanceTotal",
  "D_x00ed_as_x0020_faltantes_x0020": "DiasRestantes",
  
  // Campos que mantienen su nombre
  FechaFinalAlistamiento: "FechaFinalAlistamiento",
  
  // Otros campos que pueden tener nombres diferentes
  // Agregar más según sea necesario
};

// Mapeo inverso: nombres amigables -> nombres internos de SharePoint
export const ReverseFieldMapping: Record<string, string> = Object.entries(
  SharePointFieldMapping
).reduce((acc, [key, value]) => {
  if (value !== key) {
    acc[value] = key;
  }
  return acc;
}, {} as Record<string, string>);

/**
 * Normaliza un objeto de campos de SharePoint
 * Convierte nombres internos a nombres amigables
 */
export function normalizeSharePointFields(
  fields: Record<string, unknown>
): Record<string, unknown> {
  const normalized: Record<string, unknown> = {};
  
  for (const [key, value] of Object.entries(fields)) {
    // Verificar si existe un mapeo para este campo
    const mappedKey = SharePointFieldMapping[key] || key;
    normalized[mappedKey] = value;
    
    // También mantener el campo original si es diferente
    if (mappedKey !== key && !normalized[key]) {
      normalized[key] = value;
    }
  }
  
  return normalized;
}

/**
 * Convierte campos normalizados de vuelta a nombres internos de SharePoint
 * para enviar datos a la API
 */
export function denormalizeSharePointFields(
  fields: Record<string, unknown>
): Record<string, unknown> {
  const denormalized: Record<string, unknown> = {};
  
  for (const [key, value] of Object.entries(fields)) {
    // Verificar si existe un mapeo inverso
    const mappedKey = ReverseFieldMapping[key] || key;
    denormalized[mappedKey] = value;
  }
  
  return denormalized;
}

/**
 * Retorna YYYY-MM-DD en hora local para comparar solo el día (alineado con filtros de SharePoint).
 */
export function toDateOnlyString(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * Obtiene el valor de un campo normalizado desde un objeto de campos
 * Intenta múltiples nombres posibles
 */
export function getFieldValue(
  fields: Record<string, unknown>,
  fieldName: string
): unknown {
  // Intentar con el nombre amigable primero
  if (fields[fieldName] !== undefined) {
    return fields[fieldName];
  }
  
  // Intentar con el nombre interno
  const internalName = ReverseFieldMapping[fieldName];
  if (internalName && fields[internalName] !== undefined) {
    return fields[internalName];
  }
  
  // Si no se encuentra, buscar por coincidencia parcial (útil para campos calculados)
  const fieldKey = Object.keys(fields).find((key) => {
    // Normalizar la clave (remover guiones bajos y códigos Unicode comunes)
    const normalizedKey = key
      .replaceAll(/_x0020_/g, " ")
      .replaceAll(/_x0025_/g, "%")
      .replaceAll(/x0025__x0020_/g, "% ")
      .replaceAll(/_x00ed_/g, "í")
      .toLowerCase()
      .replaceAll(/\s+/g, "");
    
    const normalizedFieldName = fieldName.toLowerCase().replaceAll(/\s+/g, "");
    
    // Buscar coincidencias parciales
    if (normalizedKey.includes(normalizedFieldName) || normalizedFieldName.includes(normalizedKey)) {
      return true;
    }
    
    // Buscar específicamente el campo de porcentaje de avance
    if (fieldName === "PorcentajeAvanceTotal") {
      return normalizedKey.includes("avance") && normalizedKey.includes("total") ||
             key.includes("x0025") && key.includes("avance");
    }
    
    return false;
  });
  
  return fieldKey ? fields[fieldKey] : undefined;
}

/**
 * Valores válidos de cada fase F1–F16 (incluye NA = no aplica)
 */
export const FASE_OPTIONS = ["0%", "25%", "50%", "75%", "100%", "NA"] as const;
export type FaseOption = (typeof FASE_OPTIONS)[number];

const FASE_WEIGHTS = [
  0.0227272727272727, // F1
  0.0227272727272727, // F2
  0.0454545454545455, // F3
  0.0227272727272727, // F4
  0.0454545454545455, // F5
  0.113636363636364, // F6
  0.136363636363636, // F7
  0.0909090909090909, // F8
  0.136363636363636, // F9
  0.0681818181818182, // F10
  0.0454545454545455, // F11
  0.0454545454545455, // F12
  0.0227272727272727, // F13
  0.0227272727272727, // F14
  0.136363636363636, // F15
  0.0227272727272727, // F16
] as const;

export function isFaseNA(faseValue: unknown): boolean {
  if (faseValue == null) return false;
  if (typeof faseValue === "string") {
    return faseValue.trim().toUpperCase() === "NA";
  }
  return false;
}

/**
 * Convierte el valor de una fase a decimal 0–1.
 * "NA" y valores no numéricos aportan 0 (y además se excluyen del peso total).
 */
export function parseFaseToDecimal(faseValue: unknown): number {
  if (faseValue == null || isFaseNA(faseValue)) return 0;
  const rawStr =
    typeof faseValue === "string"
      ? faseValue.trim()
      : typeof faseValue === "number"
        ? String(faseValue)
        : "";
  if (!rawStr) return 0;
  const porcentajeNum = Number.parseFloat(rawStr.replaceAll("%", "")) || 0;
  return porcentajeNum / 100;
}

export function hasAnyFaseNA(fields: Record<string, unknown>): boolean {
  for (let i = 1; i <= 16; i++) {
    if (isFaseNA(getFieldValue(fields, `F${i}`))) return true;
  }
  return false;
}

function tryParseAvanceNumber(raw: unknown): number | null {
  if (raw == null || raw === "") return null;
  if (typeof raw === "number") {
    return Number.isFinite(raw) ? raw : null;
  }
  if (typeof raw !== "string") return null;

  const trimmed = raw.trim();
  if (!trimmed || trimmed.startsWith("#")) return null;

  const upper = trimmed.toUpperCase();
  if (
    upper === "NA" ||
    upper === "N/A" ||
    upper.includes("VALUE") ||
    upper.includes("ERROR")
  ) {
    return null;
  }

  const cleaned = trimmed.replaceAll("%", "").replaceAll(/[^0-9.]/g, "");
  if (!cleaned) return null;
  const num = Number.parseFloat(cleaned);
  return Number.isFinite(num) ? num : null;
}

/**
 * Calcula el % de avance total a partir de F1–F16.
 * Las fases en "NA" se excluyen y su peso se redistribuye entre las aplicables,
 * para que el avance llegue a 100% cuando todas las fases aplicables estén al 100%.
 */
export function calcularPorcentajeAvance(
  fields: Record<string, unknown>
): number {
  let pesoAplicable = 0;
  let totalPonderado = 0;

  for (let i = 1; i <= 16; i++) {
    const faseValue = getFieldValue(fields, `F${i}`);
    if (isFaseNA(faseValue)) continue;

    const peso = FASE_WEIGHTS[i - 1];
    pesoAplicable += peso;
    totalPonderado += parseFaseToDecimal(faseValue) * peso;
  }

  if (pesoAplicable <= 0) return 0;

  const porcentaje = (totalPonderado / pesoAplicable) * 100;
  return Math.round(porcentaje * 100) / 100;
}

/** Escalones de % avance para filtros indexados (15 → 99) */
export const AVANCE_FILTER_STEPS = [15, 30, 45, 60, 75, 90, 99] as const;

/**
 * Obtiene el % de avance total.
 * Si SharePoint no midió (por NA u otro error) o hay fases NA, usa cálculo local.
 */
export function parsePorcentajeAvance(
  fields: Record<string, unknown>
): number {
  const fromSharePoint = tryParseAvanceNumber(
    getFieldValue(fields, "PorcentajeAvanceTotal")
  );

  // Con NA, la columna calculada de SharePoint suele fallar → siempre recalcular
  if (hasAnyFaseNA(fields) || fromSharePoint == null) {
    return calcularPorcentajeAvance(fields);
  }

  return fromSharePoint;
}

/**
 * Indica si un % de avance cumple el filtro seleccionado.
 * Escalones 15/30/…/99 = rangos [step, nextStep); 99 = [99, 100).
 */
export function matchesPorcentajeAvanceFilter(
  avance: number,
  filterValue: string
): boolean {
  if (!filterValue) return true;
  if (filterValue === "100") return avance === 100;
  if (filterValue === "0") return avance === 0;
  if (filterValue === ">0") return avance > 0 && avance < 100;

  const step = Number(filterValue);
  if (!Number.isFinite(step)) return true;

  const stepIndex = AVANCE_FILTER_STEPS.indexOf(
    step as (typeof AVANCE_FILTER_STEPS)[number]
  );
  if (stepIndex === -1) return false;

  const lower = step;
  const upper =
    stepIndex < AVANCE_FILTER_STEPS.length - 1
      ? AVANCE_FILTER_STEPS[stepIndex + 1]
      : 100;
  return avance >= lower && avance < upper;
}

/**
 * Devuelve la clave de filtro de escalón (15|30|…|99) para un avance, o null.
 */
export function getAvanceStepFilterKey(avance: number): string | null {
  for (const step of AVANCE_FILTER_STEPS) {
    if (matchesPorcentajeAvanceFilter(avance, String(step))) {
      return String(step);
    }
  }
  return null;
}
