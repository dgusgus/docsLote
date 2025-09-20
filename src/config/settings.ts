import { ConfiguracionApp } from '../types/index.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const projectRoot = join(__dirname, '../..');

// ====================== CONFIGURACIÓN PRINCIPAL ======================
export const CONFIG: ConfiguracionApp = {
  // 👈 GOOGLE SHEETS
  SPREADSHEET_ID: "1sMu2QaY2kAy1h-V0YiKhZ2jD6VnEbYaCRXm72Fj_r58",
  RANGE: "'Hoja1'!C2:Z",
  LIMITE_PERSONAS: 500,

  // 👈 LIBREOFFICE
  SOFFICE_PATH: process.platform === 'win32' 
    ? "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
    : "/usr/bin/libreoffice", // Linux/Mac

  // 👈 PLANTILLAS (Configura aquí tus plantillas)
  PLANTILLAS: [
    {
      nombre: "certificado_trabajo",
      archivo: join(projectRoot, "plantillas/Carta certificado de trabajo TEDO 2025.docx"),
      tipo: "word",
      descripcion: "Certificado de trabajo oficial"
    },
    {
      nombre: "credenciales",
      archivo: join(projectRoot, "plantillas/06 FORM 006.xlsx"),
      tipo: "excel",
      descripcion: "Formulario de credenciales"
    },
    /* {
      nombre: "constancia_estudios",
      archivo: join(projectRoot, "plantillas/constancia_estudios.docx"),
      tipo: "word",
      descripcion: "Constancia de estudios"
    },
    {
      nombre: "diploma",
      archivo: join(projectRoot, "plantillas/diploma_template.docx"),
      tipo: "word",
      descripcion: "Diploma de certificación"
    },
    {
      nombre: "evaluacion",
      archivo: join(projectRoot, "plantillas/evaluacion.xlsx"),
      tipo: "excel",
      descripcion: "Formulario de evaluación"
    },
    {
      nombre: "registro_asistencia",
      archivo: join(projectRoot, "plantillas/asistencia.xlsx"),
      tipo: "excel",
      descripcion: "Registro de asistencia"
    },
    {
      nombre: "carta_recomendacion",
      archivo: join(projectRoot, "plantillas/recomendacion.docx"),
      tipo: "word",
      descripcion: "Carta de recomendación"
    },
    {
      nombre: "perfil_estudiante",
      archivo: join(projectRoot, "plantillas/perfil.xlsx"),
      tipo: "excel",
      descripcion: "Perfil del estudiante"
    },
    {
      nombre: "certificado_curso",
      archivo: join(projectRoot, "plantillas/certificado_curso.docx"),
      tipo: "word",
      descripcion: "Certificado de finalización de curso"
    },
    {
      nombre: "reporte_final",
      archivo: join(projectRoot, "plantillas/reporte.docx"),
      tipo: "word",
      descripcion: "Reporte final del estudiante"
    } */
  ],

  // 👈 SALIDA
  OUTPUT_DIR: join(projectRoot, "docs_generados"),
  TIPO_SALIDA: "ambos",

  // 👈 OPTIMIZACIÓN
  LIMPIAR_TEMPORALES: true,
  PROCESO_SECUENCIAL: true,
};

// ====================== CONFIGURACIONES PREDEFINIDAS ======================
export const CONFIGURACIONES_RAPIDAS = {
  // Solo generar PDFs (más rápido)
  SOLO_PDF: {
    TIPO_SALIDA: "solo-pdf" as const,
    LIMPIAR_TEMPORALES: true,
    OUTPUT_DIR: join(projectRoot, "pdfs_generados")
  },

  // Solo documentos originales (más rápido, sin LibreOffice)
  SOLO_ORIGINALES: {
    TIPO_SALIDA: "solo-originals" as const,
    LIMPIAR_TEMPORALES: false,
    OUTPUT_DIR: join(projectRoot, "documentos_originales")
  },

  // Producción completa
  PRODUCCION: {
    TIPO_SALIDA: "ambos" as const,
    LIMITE_PERSONAS: 500,
    LIMPIAR_TEMPORALES: true,
    OUTPUT_DIR: join(projectRoot, "produccion_completa")
  },

  // Pruebas rápidas
  DESARROLLO: {
    TIPO_SALIDA: "ambos" as const,
    LIMITE_PERSONAS: 3,
    LIMPIAR_TEMPORALES: false,
    OUTPUT_DIR: join(projectRoot, "pruebas_desarrollo")
  }
};

// ====================== MAPEOS EXCEL DINÁMICOS ======================
export const MAPEOS_EXCEL: Record<string, Record<string, string>> = {
  credenciales: {
    nombre: "C8",
    apellido1: "M8",
    apellido2: "U8"
  },
  evaluacion: {
    nombre: "B5",
    apellido1: "D5",
    apellido2: "F5"
  },
  registro_asistencia: {
    nombre: "A1",
    apellido1: "B1",
    apellido2: "C1"
  },
  perfil_estudiante: {
    nombre: "C3",
    apellido1: "E3",
    apellido2: "G3",
    email: "I3",
    telefono: "K3"
  }
};

// ====================== PATHS Y CONSTANTES ======================
export const PATHS = {
  CREDENTIALS: join(projectRoot, "generador-docs-31f4b831a196.json"),
  PLANTILLAS_DIR: join(projectRoot, "plantillas"),
  LOGS_DIR: join(projectRoot, "logs"),
  TEMP_DIR: join(projectRoot, "temp"),
} as const;

export const TIMEOUTS = {
  PDF_CONVERSION: 60000, // 60 segundos
  GOOGLE_SHEETS: 30000,  // 30 segundos
  FILE_OPERATION: 10000, // 10 segundos
} as const;