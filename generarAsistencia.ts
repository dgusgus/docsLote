#!/usr/bin/env node

/**
 * generarAsistencia.ts — Genera registros de asistencia SIREPRE 2026
 *
 * Fuentes de datos (mismo Google Spreadsheet):
 *   - Hoja "Hoja1"       → personas: grupo, nombre, apellidos, tipo (URBANO/MOVIL)
 *   - Hoja "Actividades" → eventos:  columna A=Fecha, B=Actividad, C=Ubicacion
 *
 * Por cada grupo × por cada actividad → un archivo .xlsx
 *
 * ESTRUCTURA DE SALIDA:
 *   asistencia_generada/
 *     grupo_1/
 *       11-03-2026_SIMULACRO_NACIONAL_SIREPRE.xlsx
 *       12-03-2026_Prueba_local_SIREPRE.xlsx
 *     grupo_26/
 *       11-03-2026_SIMULACRO_NACIONAL_SIREPRE.xlsx
 *
 * CELDAS QUE SE ESCRIBEN:
 *   Bloque MOVIL  → E6 (tipo actividad), E8 (cargo), A9 (ubicación), E9 (fecha), I9 (grupo), filas 13-22
 *   Bloque URBANO → E29 (tipo actividad), E31 (cargo), E32 (fecha), I32 (grupo), filas 35-44
 *   (A29 = =A6  y  A32 = =A9  ya son fórmulas en la plantilla — no se tocan)
 *
 * USO:
 *   tsx generarAsistencia.ts
 *   tsx generarAsistencia.ts --grupo 26
 *   tsx generarAsistencia.ts --output ./mi_salida
 *   tsx generarAsistencia.ts --help
 */

import path from 'path';
import { fileURLToPath } from 'url';
import chalk from 'chalk';
import ExcelJS from 'exceljs';
import { google } from 'googleapis';
import { Logger, crearDirectorioSeguro } from './src/utils/fileUtils.js';
import { CONFIG, PATHS } from './src/config/settings.js';
import { Persona } from './src/types/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

const PLANTILLA_PATH = path.join(__dirname, 'plantillas/LISTAS DE ASISTENCIA SIREPRE 2026.xlsx');
const OUTPUT_DEFAULT = path.join(__dirname, 'asistencia_generada');

// ====================== MAPA DE CELDAS DE LA PLANTILLA ======================
//
//  Bloque MOVIL (parte superior):
//    E6  → "TIPO DE ACTIVIDAD: ..."
//    E8  → "CARGO: Operador de Transmisión Móvil"
//    A9  → "UBICACIÓN: ..."
//    E9  → "FECHA: DD-MM-YYYY"
//    I9  → "GRUPO:  XX"
//    filas 13–22 (col A=nro, B=nombre, D=unidad)
//
//  Bloque URBANO (parte inferior):
//    E29 → "TIPO DE ACTIVIDAD: ..."
//    E31 → "CARGO: Operador de Transmisión Urbano"
//    A29 → =A6  (fórmula en plantilla, NO se toca)
//    A32 → =A9  (fórmula en plantilla, NO se toca)
//    E32 → "FECHA: DD-MM-YYYY"
//    I32 → "GRUPO:  XX"
//    filas 35–44 (col A=nro, B=nombre, D=unidad)

const BLOQUE_MOVIL = {
  celdaTipoActividad: 'E6',
  celdaCargo:         'E8',
  celdaUbicacion:     'A9',
  celdaFecha:         'E9',
  celdaGrupo:         'I9',
  filaInicio:         13,
  maxPersonas:        9,   // filas 13–22
  cargoTexto:         'CARGO: Operador de Transmisión Móvil',
};

const BLOQUE_URBANO = {
  celdaTipoActividad: 'E29',
  celdaCargo:         'E31',
  // A29 y A32 son fórmulas =A6 / =A9 — no se escriben
  celdaFecha:         'E32',
  celdaGrupo:         'I32',
  filaInicio:         35,
  maxPersonas:        9,   // filas 35–44
  cargoTexto:         'CARGO: Operador de Transmisión Urbano',
};

// ====================== TIPOS ======================

interface Actividad {
  fecha:     string;   // DD-MM-YYYY
  actividad: string;
  ubicacion: string;
}

// ====================== GOOGLE SHEETS ======================

async function crearSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: PATHS.CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  return google.sheets({ version: 'v4', auth });
}

async function obtenerPersonas(): Promise<Persona[]> {
  const sheets = await crearSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: "'Hoja1'!B2:Z",
  });
  const filas = res.data.values || [];
  return filas
    .filter((f: any[]) => f[1]?.toString().trim())
    .map((f: any[], i: number): Persona => ({
      indice:       i + 1,
      grupo:        f[0]?.toString().trim()  || '',
      nombre:       f[1]?.toString().trim()  || '',
      apellido1:    f[2]?.toString().trim()  || '',
      apellido2:    f[3]?.toString().trim()  || '',
      documento:    f[4]?.toString().trim()  || '',
      telefono:     f[5]?.toString().trim()  || '',
      email:        f[6]?.toString().trim()  || '',
      cargo:        f[7]?.toString().trim()  || '',
      fecha_inicio: f[8]?.toString().trim()  || '',
      fecha_fin:    f[9]?.toString().trim()  || '',
      tipo:         f[10]?.toString().trim() || '',
    }));
}

async function obtenerActividades(): Promise<Actividad[]> {
  const sheets = await crearSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: "'Actividades'!A2:C",
  });
  const filas = res.data.values || [];
  return filas
    .filter((f: any[]) => f[0]?.toString().trim() && f[1]?.toString().trim())
    .map((f: any[]): Actividad => ({
      fecha:     normalizarFecha(f[0]?.toString().trim() || ''),
      actividad: f[1]?.toString().trim() || '',
      ubicacion: f[2]?.toString().trim() || '',
    }));
}

// ====================== UTILIDADES ======================

/** Acepta DD/MM/YYYY o DD-MM-YYYY → devuelve DD-MM-YYYY */
function normalizarFecha(str: string): string {
  return str.replace(/\//g, '-');
}

/** Slug seguro para nombre de archivo */
function slugActividad(str: string): string {
  return str
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9\s]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .substring(0, 40);
}

function normalizarTipo(tipo: string | undefined): 'URBANO' | 'MOVIL' | null {
  if (!tipo) return null;
  const t = tipo.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  if (t.includes('URBAN')) return 'URBANO';
  if (t.includes('MOVIL') || t.includes('MOBIL')) return 'MOVIL';
  return null;
}

// ====================== LÓGICA EXCEL ======================

function rellenarTabla(
  ws: ExcelJS.Worksheet,
  personas: Persona[],
  filaInicio: number,
  maxPersonas: number
): void {
  for (let i = 0; i < maxPersonas; i++) {
    const fila = filaInicio + i;
    const p    = personas[i];
    ws.getCell(fila, 1).value = i + 1;
    if (p) {
      ws.getCell(fila, 2).value = [p.nombre, p.apellido1, p.apellido2]
        .filter(Boolean).join(' ').toUpperCase().trim();
      ws.getCell(fila, 4).value = (p as any).unidad || "TIC'S";
    } else {
      ws.getCell(fila, 2).value = null;
      ws.getCell(fila, 4).value = null;
    }
  }
}

async function generarArchivo(
  moviles:   Persona[],
  urbanos:   Persona[],
  grupo:     string,
  actividad: Actividad,
  outputDir: string
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(PLANTILLA_PATH);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('La plantilla no tiene hojas de trabajo');

  // ── Bloque MOVIL ──────────────────────────────────────────────────────────
  ws.getCell(BLOQUE_MOVIL.celdaTipoActividad).value = `TIPO DE ACTIVIDAD: ${actividad.actividad}`;
  ws.getCell(BLOQUE_MOVIL.celdaCargo).value         = BLOQUE_MOVIL.cargoTexto;
  ws.getCell(BLOQUE_MOVIL.celdaUbicacion).value     = `UBICACIÓN: ${actividad.ubicacion}`;
  ws.getCell(BLOQUE_MOVIL.celdaFecha).value         = `FECHA: ${actividad.fecha}`;
  ws.getCell(BLOQUE_MOVIL.celdaGrupo).value         = `GRUPO:  ${grupo}`;
  rellenarTabla(ws, moviles.slice(0, BLOQUE_MOVIL.maxPersonas), BLOQUE_MOVIL.filaInicio, BLOQUE_MOVIL.maxPersonas);

  // ── Bloque URBANO ─────────────────────────────────────────────────────────
  // A29 = =A6  y  A32 = =A9  son fórmulas en la plantilla → NO se tocan
  ws.getCell(BLOQUE_URBANO.celdaTipoActividad).value = `TIPO DE ACTIVIDAD: ${actividad.actividad}`;
  ws.getCell(BLOQUE_URBANO.celdaCargo).value         = BLOQUE_URBANO.cargoTexto;
  ws.getCell(BLOQUE_URBANO.celdaFecha).value         = `FECHA: ${actividad.fecha}`;
  ws.getCell(BLOQUE_URBANO.celdaGrupo).value         = `GRUPO:  ${grupo}`;
  rellenarTabla(ws, urbanos.slice(0, BLOQUE_URBANO.maxPersonas), BLOQUE_URBANO.filaInicio, BLOQUE_URBANO.maxPersonas);

  // ── Guardar ───────────────────────────────────────────────────────────────
  const carpetaGrupo = path.join(outputDir, `grupo_${grupo.replace(/\s+/g, '_')}`);
  await crearDirectorioSeguro(carpetaGrupo);

  const nombreArchivo = `${actividad.fecha}_${slugActividad(actividad.actividad)}.xlsx`;
  await wb.xlsx.writeFile(path.join(carpetaGrupo, nombreArchivo));
}

// ====================== ARGUMENTOS ======================

interface Args { grupo: string | null; outputDir: string; }

function parsearArgs(): Args {
  const argv = process.argv.slice(2);
  if (argv.includes('--help') || argv.includes('-h')) { mostrarAyuda(); process.exit(0); }
  const get = (flag: string) => { const i = argv.indexOf(flag); return i !== -1 && argv[i + 1] ? argv[i + 1] : null; };
  return { grupo: get('--grupo'), outputDir: get('--output') || OUTPUT_DEFAULT };
}

function mostrarAyuda(): void {
  console.log(chalk.bold.cyan(`
╔══════════════════════════════════════════════════════════════╗
║  📋 GENERADOR DE ASISTENCIA SIREPRE 2026                     ║
╚══════════════════════════════════════════════════════════════╝
`));
  console.log('Genera un Excel por grupo × actividad.');
  console.log('  - Personas  → hoja "Hoja1"       (grupo, nombre, tipo URBANO/MOVIL)');
  console.log('  - Actividades → hoja "Actividades" (Fecha | Actividad | Ubicacion)');
  console.log('');
  console.log('USO:');
  console.log('  tsx generarAsistencia.ts');
  console.log('  tsx generarAsistencia.ts --grupo 26');
  console.log('  tsx generarAsistencia.ts --output ./mi_salida');
  console.log('');
  console.log('ESTRUCTURA DE SALIDA:');
  console.log('  asistencia_generada/');
  console.log('    grupo_1/');
  console.log('      11-03-2026_SIMULACRO_NACIONAL_SIREPRE.xlsx');
  console.log('      12-03-2026_Prueba_local_SIREPRE.xlsx');
  console.log('    grupo_26/');
  console.log('      11-03-2026_SIMULACRO_NACIONAL_SIREPRE.xlsx');
}

// ====================== MAIN ======================

async function main(): Promise<void> {
  console.log(chalk.bold.cyan('\n📋 GENERADOR DE ASISTENCIA SIREPRE 2026\n'));

  const { grupo, outputDir } = parsearArgs();
  if (grupo) Logger.info(`Grupo:  ${grupo}`);
  Logger.info(`Salida: ${outputDir}\n`);

  // 1. Obtener personas y actividades en paralelo
  Logger.progress('Conectando con Google Sheets...');
  let personas: Persona[];
  let actividades: Actividad[];

  try {
    [personas, actividades] = await Promise.all([obtenerPersonas(), obtenerActividades()]);
  } catch (err) {
    Logger.error(`Error leyendo Google Sheets: ${err}`);
    process.exit(1);
  }

  Logger.success(`Personas: ${personas.length}  |  Actividades: ${actividades.length}`);

  if (actividades.length === 0) {
    Logger.error('No hay actividades en la hoja "Actividades". Verificá columnas A (Fecha), B (Actividad), C (Ubicacion).');
    process.exit(1);
  }

  // 2. Filtrar por grupo si se indicó
  if (grupo) {
    personas = personas.filter(p => p.grupo?.toString().trim() === grupo.trim());
    if (personas.length === 0) { Logger.error(`No hay personas en el grupo "${grupo}"`); process.exit(1); }
  }

  // 3. Agrupar personas por número de grupo
  const porGrupo = new Map<string, Persona[]>();
  for (const p of personas) {
    const g = (p.grupo || 'SIN_GRUPO').toString().trim();
    if (!porGrupo.has(g)) porGrupo.set(g, []);
    porGrupo.get(g)!.push(p);
  }

  const totalArchivos = porGrupo.size * actividades.length;
  Logger.info(`Grupos: ${porGrupo.size}  ×  Actividades: ${actividades.length}  =  ${totalArchivos} archivos\n`);

  // 4. Generar grupo × actividad
  let exitosos = 0;
  let errores  = 0;

  for (const [g, miembros] of porGrupo) {
    const moviles = miembros.filter(p => normalizarTipo(p.tipo) === 'MOVIL');
    const urbanos = miembros.filter(p => normalizarTipo(p.tipo) === 'URBANO');
    const sinTipo = miembros.filter(p => normalizarTipo(p.tipo) === null);

    console.log(chalk.bold(`\n📁 grupo_${g}/  (${moviles.length}M + ${urbanos.length}U${sinTipo.length ? ` + ${sinTipo.length} sin tipo` : ''})`));
    if (sinTipo.length) {
      sinTipo.forEach(p => Logger.warn(`   ⚠️  Sin tipo reconocido: ${p.nombre} ${p.apellido1} — "${p.tipo}"`));
    }

    for (const act of actividades) {
      const nombreArchivo = `${act.fecha}_${slugActividad(act.actividad)}.xlsx`;
      try {
        await generarArchivo(moviles, urbanos, g, act, outputDir);
        console.log(`   ✅ ${nombreArchivo}`);
        exitosos++;
      } catch (err) {
        console.log(`   ❌ ${nombreArchivo} — ${err}`);
        errores++;
      }
    }
  }

  // 5. Resumen
  Logger.separador();
  console.log(chalk.bold('\n📊 RESUMEN'));
  console.log(`✅ Generados: ${exitosos} / ${totalArchivos}`);
  if (errores) console.log(`❌ Errores:   ${errores}`);
  console.log(`📂 En:        ${outputDir}`);

  process.exit(errores > 0 ? 1 : 0);
}

process.on('unhandledRejection', (r) => { Logger.error(`${r}`); process.exit(1); });
process.on('SIGINT', () => { console.log(chalk.yellow('\n⚠️  Interrumpido')); process.exit(0); });

main();