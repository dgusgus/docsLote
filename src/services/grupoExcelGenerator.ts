import fs from 'fs/promises';
import path from 'path';
import ExcelJS from 'exceljs';
import { Persona } from '../types/index.js';
import { Logger, crearDirectorioSeguro } from '../utils/fileUtils.js';

// ====================== CONFIGURACIÓN DE LA PLANTILLA DE ASISTENCIA ======================
//
// La plantilla tiene DOS tablas en la misma hoja:
//
//   TABLA URBANO  → filas de datos: 11 a 20  (máx. 9 operadores)
//   TABLA MÓVIL   → filas de datos: 32 a 41  (máx. 9 operadores)
//
// Cada fila de dato tiene:
//   Col A  → número de orden (1, 2, 3...)
//   Col B  → NOMBRE COMPLETO (celda combinada B:C)
//   Col D  → UNIDAD/DIRECCIÓN
//
// El GRUPO se escribe en la celda E8 (y E29 referencia a E8 con fórmula =E8)
// en el formato: "FECHA: DD-MM-YYYY                              GRUPO: XX"
//
// TIPO esperado en el campo `tipo` de Persona: "URBANO" o "MOVIL"
// (se normaliza: uppercase, sin tildes, se acepta también "MOVIL", "MÓVIL", "URBANO")

const FILA_INICIO_URBANO = 11;
const FILA_INICIO_MOVIL  = 33;
const MAX_OPERADORES     = 9;

const COL_NUMERO    = 1; // A
const COL_NOMBRE    = 2; // B
const COL_UNIDAD    = 4; // D

// Celda que contiene "FECHA: ... GRUPO: XX"
const CELDA_FECHA_GRUPO = 'E8';

// ====================== TIPOS ======================
export interface ResultadoGrupoExcel {
  grupo: string;
  archivo: string;
  urbanos: number;
  moviles: number;
  errores: string[];
}

// ====================== SERVICIO ======================
export class GrupoExcelGenerator {

  /**
   * Normaliza el tipo de operador a 'URBANO' o 'MOVIL'.
   * Acepta variantes con/sin tilde, mayúsculas/minúsculas.
   */
  private normalizarTipo(tipo: string | undefined): 'URBANO' | 'MOVIL' | null {
    if (!tipo) return null;
    const t = tipo.toUpperCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '') // quitar tildes
      .trim();

    if (t.includes('URBAN')) return 'URBANO';
    if (t.includes('MOVIL') || t.includes('MOBIL')) return 'MOVIL';
    return null;
  }

  /**
   * Agrupa un array de personas por número de grupo.
   */
  agruparPorGrupo(personas: Persona[]): Map<string, Persona[]> {
    const mapa = new Map<string, Persona[]>();
    for (const p of personas) {
      const g = (p.grupo || 'SIN_GRUPO').toString().trim();
      if (!mapa.has(g)) mapa.set(g, []);
      mapa.get(g)!.push(p);
    }
    return mapa;
  }

  /**
   * Genera un solo archivo Excel para un grupo, rellenando ambas tablas.
   *
   * @param plantillaPath  Ruta al archivo .xlsx plantilla
   * @param personas       Todas las personas de ese grupo (mezcla de urbanos y móviles)
   * @param grupo          Número/nombre del grupo
   * @param fecha          Fecha en formato DD-MM-YYYY (se usa en la cabecera)
   * @param outputDir      Carpeta donde se guarda el archivo
   * @returns              Ruta del archivo generado
   */
  async generarExcelGrupo(
    plantillaPath: string,
    personas: Persona[],
    grupo: string,
    fecha: string,
    outputDir: string
  ): Promise<ResultadoGrupoExcel> {
    const errores: string[] = [];

    // Separar por tipo
    const urbanos = personas.filter(p => this.normalizarTipo(p.tipo) === 'URBANO');
    const moviles = personas.filter(p => this.normalizarTipo(p.tipo) === 'MOVIL');

    // Advertir si hay personas sin tipo reconocido
    const sinTipo = personas.filter(p => this.normalizarTipo(p.tipo) === null);
    for (const p of sinTipo) {
      const msg = `${p.nombre} ${p.apellido1} tiene tipo no reconocido: "${p.tipo}"`;
      errores.push(msg);
      Logger.warn(`  ⚠️  ${msg}`);
    }

    if (urbanos.length > MAX_OPERADORES) {
      errores.push(`Grupo ${grupo}: hay ${urbanos.length} urbanos pero la tabla solo tiene ${MAX_OPERADORES} filas`);
      Logger.warn(`  ⚠️  Se truncarán los urbanos a ${MAX_OPERADORES}`);
    }
    if (moviles.length > MAX_OPERADORES) {
      errores.push(`Grupo ${grupo}: hay ${moviles.length} móviles pero la tabla solo tiene ${MAX_OPERADORES} filas`);
      Logger.warn(`  ⚠️  Se truncarán los móviles a ${MAX_OPERADORES}`);
    }

    // Cargar plantilla
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(plantillaPath);
    const ws = wb.worksheets[0];

    if (!ws) throw new Error('La plantilla no tiene hojas de trabajo');

    // --- Actualizar cabecera: FECHA y GRUPO ---
    // Preservar el texto exacto que espera la plantilla
    const celdaFechaGrupo = ws.getCell(CELDA_FECHA_GRUPO);
    // Construir el texto manteniendo el espaciado original de la plantilla
    const espacios = ' '.repeat(50);
    celdaFechaGrupo.value = `FECHA: ${fecha}${espacios}GRUPO: ${grupo}`;

    // --- Rellenar tabla URBANO (filas 11–19) ---
    this.rellenarTabla(ws, urbanos.slice(0, MAX_OPERADORES), FILA_INICIO_URBANO);

    // --- Rellenar tabla MÓVIL (filas 32–40) ---
    this.rellenarTabla(ws, moviles.slice(0, MAX_OPERADORES), FILA_INICIO_MOVIL);

    // --- Guardar ---
    await crearDirectorioSeguro(outputDir);
    const nombreArchivo = `asistencia_grupo_${grupo.replace(/\s+/g, '_')}.xlsx`;
    const rutaSalida = path.join(outputDir, nombreArchivo);
    await wb.xlsx.writeFile(rutaSalida);

    Logger.success(`  📋 Grupo ${grupo}: ${urbanos.length} urbanos, ${moviles.length} móviles → ${nombreArchivo}`);

    return {
      grupo,
      archivo: rutaSalida,
      urbanos: urbanos.length,
      moviles: moviles.length,
      errores,
    };
  }

  /**
   * Rellena las filas de datos de una tabla (urbano o móvil).
   * Limpia primero las filas que tenían datos previos de la plantilla,
   * luego escribe los nuevos datos.
   */
  private rellenarTabla(
    ws: ExcelJS.Worksheet,
    personas: Persona[],
    filaInicio: number
  ): void {
    for (let i = 0; i < MAX_OPERADORES; i++) {
      const fila = filaInicio + i;
      const persona = personas[i]; // undefined si no hay suficientes personas

      // Número de orden — siempre va (ya está en la plantilla pero lo confirmamos)
      ws.getCell(fila, COL_NUMERO).value = i + 1;

      if (persona) {
        // Nombre completo en mayúsculas
        const nombreCompleto = [persona.nombre, persona.apellido1, persona.apellido2]
          .filter(Boolean)
          .join(' ')
          .toUpperCase()
          .trim();

        ws.getCell(fila, COL_NOMBRE).value = nombreCompleto;
        ws.getCell(fila, COL_UNIDAD).value = persona.unidad || "TIC'S";
      } else {
        // Fila vacía — limpiar por si la plantilla tenía datos de ejemplo
        ws.getCell(fila, COL_NOMBRE).value = null;
        ws.getCell(fila, COL_UNIDAD).value = null;
      }
    }
  }

  /**
   * Genera los Excel de asistencia para TODOS los grupos encontrados
   * en la lista de personas.
   *
   * @param plantillaPath  Ruta al .xlsx plantilla
   * @param todasPersonas  Todas las personas (de todos los grupos)
   * @param fecha          Fecha para la cabecera (DD-MM-YYYY)
   * @param outputDir      Carpeta raíz de salida
   */
  async generarTodosLosGrupos(
    plantillaPath: string,
    todasPersonas: Persona[],
    fecha: string,
    outputDir: string
  ): Promise<ResultadoGrupoExcel[]> {
    const grupos = this.agruparPorGrupo(todasPersonas);
    const resultados: ResultadoGrupoExcel[] = [];

    Logger.titulo(`Generando registros de asistencia para ${grupos.size} grupo(s)`);

    for (const [grupo, personas] of grupos) {
      try {
        const carpetaGrupo = path.join(outputDir, `grupo_${grupo.replace(/\s+/g, '_')}`);
        const resultado = await this.generarExcelGrupo(
          plantillaPath,
          personas,
          grupo,
          fecha,
          carpetaGrupo
        );
        resultados.push(resultado);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        Logger.error(`  ❌ Error en grupo ${grupo}: ${msg}`);
        resultados.push({ grupo, archivo: '', urbanos: 0, moviles: 0, errores: [msg] });
      }
    }

    return resultados;
  }
}