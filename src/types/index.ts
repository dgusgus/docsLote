// ====================== TIPOS PRINCIPALES ======================
export type TipoSalida = "solo-pdf" | "solo-originals" | "ambos";
export type TipoPlantilla = "word" | "excel";
export type ModoEjecucion = "todos" | "rango" | "especifico";

// ====================== INTERFACES ======================
export interface Persona {
  apellido2: string;
  nombre: string;
  apellido1: string;
  email?: string;
  telefono?: string;
  documento?: string;
  curso?: string;
  fecha_inicio?: string;
  fecha_fin?: string;
  [key: string]: any; // Para campos adicionales dinámicos
}

export interface PlantillaConfig {
  nombre: string;
  archivo: string;
  tipo: TipoPlantilla;
  campos?: string[]; // Campos específicos que usa esta plantilla
  descripcion?: string;
}

export interface ConfiguracionApp {
  SPREADSHEET_ID: string;
  RANGE: string;
  LIMITE_PERSONAS: number;
  SOFFICE_PATH: string;
  PLANTILLAS: PlantillaConfig[];
  OUTPUT_DIR: string;
  TIPO_SALIDA: TipoSalida;
  LIMPIAR_TEMPORALES: boolean;
  PROCESO_SECUENCIAL: boolean;
}

export interface OpcionesEjecucion {
  modo: ModoEjecucion;
  tipoSalida: TipoSalida;
  rangoInicio?: number;
  rangoFin?: number;
  nombreEspecifico?: string;
  plantillasEspecificas?: string[];
  outputDir?: string;
}

export interface ResultadoProceso {
  exitos: number;
  errores: number;
  tiempoTotal: number;
  personasProcesadas: number;
  detalles: ResultadoPersona[];
}

export interface ResultadoPersona {
  nombre: string;
  exitos: number;
  errores: number;
  documentosGenerados: string[];
  erroresDetallados: string[];
}

export interface LogLevel {
  INFO: 'info';
  WARN: 'warn';
  ERROR: 'error';
  SUCCESS: 'success';
}