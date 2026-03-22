#!/usr/bin/env node

/**
 * generarAsistencia.ts — Genera registros de asistencia SIREPRE 2026
 *
 * Fuentes de datos (mismo Google Spreadsheet):
 *   - Hoja "Hoja1"       → personas: grupo, nombre, apellidos, tipo (URBANO/MOVIL)
 *   - Hoja "Actividades" → eventos:  columna A=Fecha, B=Actividad, C=Ubicacion
 *
 * Flujo por cada grupo:
 *   1. Genera un .xlsx por actividad  →  grupo_1/11-03-2026_SIMULACRO.xlsx
 *   2. Convierte cada .xlsx a .pdf    →  grupo_1/11-03-2026_SIMULACRO.pdf  (temporal)
 *   3. Mergea todos los PDFs ordenados por fecha → grupo_1/grupo_1.pdf
 *   4. Elimina los PDFs intermedios y los xlsx
 *
 * ESTRUCTURA DE SALIDA:
 *   asistencia_generada/
 *     grupo_1/
 *       grupo_1.pdf        ← todas las fechas en orden
 *     grupo_26/
 *       grupo_26.pdf
 *
 * USO:
 *   tsx generarAsistencia.ts
 *   tsx generarAsistencia.ts --grupo 26
 *   tsx generarAsistencia.ts --keep-xlsx      → conserva los .xlsx generados
 *   tsx generarAsistencia.ts --output ./salida
 *   tsx generarAsistencia.ts --help
 */

import path from 'path';
import fs from 'fs/promises';
import { fileURLToPath } from 'url';
import { spawn } from 'child_process';
import chalk from 'chalk';
import ExcelJS from 'exceljs';
import { google } from 'googleapis';
import { PDFDocument } from 'pdf-lib';
import { Logger, crearDirectorioSeguro } from './src/utils/fileUtils.js';
import { CONFIG, PATHS } from './src/config/settings.js';
import { Persona } from './src/types/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

const PLANTILLA_PATH = path.join(__dirname, 'plantillas/LISTAS DE ASISTENCIA SIREPRE 2026.xlsx');
const OUTPUT_DEFAULT = path.join(__dirname, 'asistencia_generada');

// ====================== MAPA DE CELDAS ======================

const BLOQUE_MOVIL = {
  celdaTipoActividad: 'E6',
  celdaCargo:         'E8',
  celdaUbicacion:     'A9',
  celdaFecha:         'E9',
  celdaGrupo:         'I9',
  filaInicio:         13,
  maxPersonas:        9,
  cargoTexto:         'CARGO: Operador de Transmisión Móvil',
};

const BLOQUE_URBANO = {
  celdaTipoActividad: 'E29',
  celdaCargo:         'E31',
  celdaFecha:         'E32',
  celdaGrupo:         'I32',
  filaInicio:         35,
  maxPersonas:        9,
  cargoTexto:         'CARGO: Operador de Transmisión Urbano',
};

// ====================== TIPOS ======================

interface Actividad {
  fecha:     string;
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

function normalizarFecha(str: string): string {
  return str.replace(/\//g, '-');
}

/** DD-MM-YYYY → YYYY-MM-DD para ordenar cronológicamente */
function fechaParaOrdenar(fechaDDMMYYYY: string): string {
  const [dd, mm, yyyy] = fechaDDMMYYYY.split('-');
  return `${yyyy}-${mm}-${dd}`;
}

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

// ====================== EXCEL ======================

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

async function generarXlsx(
  moviles:     Persona[],
  urbanos:     Persona[],
  grupo:       string,
  actividad:   Actividad,
  carpetaGrupo: string
): Promise<string> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(PLANTILLA_PATH);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('La plantilla no tiene hojas de trabajo');

  ws.getCell(BLOQUE_MOVIL.celdaTipoActividad).value = `TIPO DE ACTIVIDAD: ${actividad.actividad}`;
  ws.getCell(BLOQUE_MOVIL.celdaCargo).value         = BLOQUE_MOVIL.cargoTexto;
  ws.getCell(BLOQUE_MOVIL.celdaUbicacion).value     = `UBICACIÓN: ${actividad.ubicacion}`;
  ws.getCell(BLOQUE_MOVIL.celdaFecha).value         = `FECHA: ${actividad.fecha}`;
  ws.getCell(BLOQUE_MOVIL.celdaGrupo).value         = `GRUPO:  ${grupo}`;
  rellenarTabla(ws, moviles.slice(0, BLOQUE_MOVIL.maxPersonas), BLOQUE_MOVIL.filaInicio, BLOQUE_MOVIL.maxPersonas);

  ws.getCell(BLOQUE_URBANO.celdaTipoActividad).value = `TIPO DE ACTIVIDAD: ${actividad.actividad}`;
  ws.getCell(BLOQUE_URBANO.celdaCargo).value         = BLOQUE_URBANO.cargoTexto;
  ws.getCell(BLOQUE_URBANO.celdaFecha).value         = `FECHA: ${actividad.fecha}`;
  ws.getCell(BLOQUE_URBANO.celdaGrupo).value         = `GRUPO:  ${grupo}`;
  rellenarTabla(ws, urbanos.slice(0, BLOQUE_URBANO.maxPersonas), BLOQUE_URBANO.filaInicio, BLOQUE_URBANO.maxPersonas);

  const nombreArchivo = `${actividad.fecha}_${slugActividad(actividad.actividad)}.xlsx`;
  const rutaXlsx = path.join(carpetaGrupo, nombreArchivo);
  await wb.xlsx.writeFile(rutaXlsx);
  return rutaXlsx;
}

// ====================== PDF — conversión y merge ======================

/** Convierte un .xlsx a .pdf con LibreOffice. Devuelve la ruta del PDF generado. */
function xlsxToPdf(xlsxPath: string, outputDir: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const pdfPath = path.join(outputDir, path.basename(xlsxPath, '.xlsx') + '.pdf');

    const proc = spawn(CONFIG.SOFFICE_PATH, [
      '--headless',
      '--convert-to', 'pdf',
      '--outdir', outputDir,
      xlsxPath,
    ], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });

    let stderr = '';
    proc.stderr?.on('data', (d) => { stderr += d.toString(); });

    proc.on('exit', async (code) => {
      if (code !== 0) return reject(new Error(`LibreOffice falló (código ${code}): ${stderr}`));
      try {
        await fs.access(pdfPath);
        resolve(pdfPath);
      } catch {
        reject(new Error(`PDF no generado en: ${pdfPath}`));
      }
    });

    proc.on('error', (err) => reject(new Error(`Error ejecutando LibreOffice: ${err.message}`)));

    // Timeout 60s
    setTimeout(() => { proc.kill(); reject(new Error('Timeout convirtiendo a PDF')); }, 60_000);
  });
}

/**
 * Mergea una lista de PDFs (en el orden dado) en un solo PDF de salida.
 * Usa pdf-lib que ya está en el proyecto.
 */
async function mergePdfs(rutasPdf: string[], rutaSalida: string): Promise<void> {
  const mergeDoc = await PDFDocument.create();

  for (const ruta of rutasPdf) {
    const bytes  = await fs.readFile(ruta);
    const srcDoc = await PDFDocument.load(bytes);
    const pages  = await mergeDoc.copyPages(srcDoc, srcDoc.getPageIndices());
    pages.forEach(p => mergeDoc.addPage(p));
  }

  const pdfBytes = await mergeDoc.save();
  await fs.writeFile(rutaSalida, pdfBytes);
}

/** Elimina una lista de archivos ignorando errores individuales. */
async function eliminarArchivos(rutas: string[]): Promise<void> {
  await Promise.allSettled(rutas.map(r => fs.unlink(r)));
}

// ====================== ARGUMENTOS ======================

interface Args { grupo: string | null; outputDir: string; keepXlsx: boolean; }

function parsearArgs(): Args {
  const argv = process.argv.slice(2);
  if (argv.includes('--help') || argv.includes('-h')) { mostrarAyuda(); process.exit(0); }
  const get = (flag: string) => { const i = argv.indexOf(flag); return i !== -1 && argv[i + 1] ? argv[i + 1] : null; };
  return {
    grupo:    get('--grupo'),
    outputDir: get('--output') || OUTPUT_DEFAULT,
    keepXlsx: argv.includes('--keep-xlsx'),
  };
}

function mostrarAyuda(): void {
  console.log(chalk.bold.cyan(`
╔══════════════════════════════════════════════════════════════╗
║  📋 GENERADOR DE ASISTENCIA SIREPRE 2026                     ║
╚══════════════════════════════════════════════════════════════╝
`));
  console.log('Genera un PDF unificado por grupo con todas las fechas en orden.');
  console.log('');
  console.log('USO:');
  console.log('  tsx generarAsistencia.ts');
  console.log('  tsx generarAsistencia.ts --grupo 26');
  console.log('  tsx generarAsistencia.ts --keep-xlsx   → conserva los .xlsx intermedios');
  console.log('  tsx generarAsistencia.ts --output ./mi_salida');
  console.log('');
  console.log('ESTRUCTURA DE SALIDA:');
  console.log('  asistencia_generada/');
  console.log('    grupo_1/');
  console.log('      grupo_1.pdf    ← todas las fechas en orden');
  console.log('    grupo_26/');
  console.log('      grupo_26.pdf');
}

// ====================== MAIN ======================

async function main(): Promise<void> {
  console.log(chalk.bold.cyan('\n📋 GENERADOR DE ASISTENCIA SIREPRE 2026\n'));

  const { grupo, outputDir, keepXlsx } = parsearArgs();
  if (grupo)    Logger.info(`Grupo:      ${grupo}`);
  if (keepXlsx) Logger.info('Modo:       conservando .xlsx intermedios');
  Logger.info(`Salida:     ${outputDir}\n`);

  // 1. Obtener datos
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

  // Ordenar actividades cronológicamente por fecha
  actividades.sort((a, b) => fechaParaOrdenar(a.fecha).localeCompare(fechaParaOrdenar(b.fecha)));

  // 2. Filtrar por grupo
  if (grupo) {
    personas = personas.filter(p => p.grupo?.toString().trim() === grupo.trim());
    if (personas.length === 0) { Logger.error(`No hay personas en el grupo "${grupo}"`); process.exit(1); }
  }

  // 3. Agrupar personas
  const porGrupo = new Map<string, Persona[]>();
  for (const p of personas) {
    const g = (p.grupo || 'SIN_GRUPO').toString().trim();
    if (!porGrupo.has(g)) porGrupo.set(g, []);
    porGrupo.get(g)!.push(p);
  }

  Logger.info(`Grupos: ${porGrupo.size}  ×  Actividades: ${actividades.length}  =  ${porGrupo.size * actividades.length} hojas\n`);

  // 4. Procesar grupo por grupo
  let exitosos = 0;
  let errores  = 0;

  for (const [g, miembros] of porGrupo) {
    const moviles = miembros.filter(p => normalizarTipo(p.tipo) === 'MOVIL');
    const urbanos = miembros.filter(p => normalizarTipo(p.tipo) === 'URBANO');
    const sinTipo = miembros.filter(p => normalizarTipo(p.tipo) === null);

    console.log(chalk.bold(`\n📁 grupo_${g}/  (${moviles.length}M + ${urbanos.length}U${sinTipo.length ? ` + ${sinTipo.length} sin tipo` : ''})`));
    if (sinTipo.length) sinTipo.forEach(p => Logger.warn(`   ⚠️  Sin tipo: ${p.nombre} ${p.apellido1} — "${p.tipo}"`));

    const carpetaGrupo = path.join(outputDir, `grupo_${g.replace(/\s+/g, '_')}`);
    await crearDirectorioSeguro(carpetaGrupo);

    const rutasPdfIntermedios: string[] = [];
    const rutasXlsxIntermedios: string[] = [];
    let errorEnGrupo = false;

    // 4a. Generar xlsx + convertir a PDF por cada actividad
    for (const act of actividades) {
      const etiqueta = `${act.fecha} — ${act.actividad.substring(0, 35)}`;
      try {
        // xlsx
        const rutaXlsx = await generarXlsx(moviles, urbanos, g, act, carpetaGrupo);
        rutasXlsxIntermedios.push(rutaXlsx);

        // pdf
        process.stdout.write(`   🔄 ${etiqueta} → PDF...`);
        const rutaPdf = await xlsxToPdf(rutaXlsx, carpetaGrupo);
        rutasPdfIntermedios.push(rutaPdf);
        process.stdout.write(' ✅\n');

      } catch (err) {
        process.stdout.write('\n');
        Logger.error(`   ❌ ${etiqueta}: ${err}`);
        errorEnGrupo = true;
      }
    }

    if (rutasPdfIntermedios.length === 0) {
      Logger.error(`   Sin PDFs generados para grupo ${g}`);
      errores++;
      continue;
    }

    // 4b. Mergear todos los PDFs del grupo en uno solo (ya están ordenados por fecha)
    const rutaPdfFinal = path.join(carpetaGrupo, `grupo_${g.replace(/\s+/g, '_')}.pdf`);
    try {
      process.stdout.write(`   📎 Mergeando ${rutasPdfIntermedios.length} PDFs → grupo_${g}.pdf...`);
      await mergePdfs(rutasPdfIntermedios, rutaPdfFinal);
      process.stdout.write(' ✅\n');
      Logger.success(`   📄 grupo_${g}.pdf  (${rutasPdfIntermedios.length} fechas)`);
      exitosos++;
    } catch (err) {
      process.stdout.write('\n');
      Logger.error(`   ❌ Error mergeando PDFs del grupo ${g}: ${err}`);
      errores++;
      errorEnGrupo = true;
    }

    // 4c. Limpiar archivos intermedios
    await eliminarArchivos(rutasPdfIntermedios);
    if (!keepXlsx) await eliminarArchivos(rutasXlsxIntermedios);

    if (errorEnGrupo) errores++;
  }

  // 5. Resumen
  Logger.separador();
  console.log(chalk.bold('\n📊 RESUMEN'));
  console.log(`✅ PDFs unificados: ${exitosos} / ${porGrupo.size}`);
  if (errores) console.log(`❌ Grupos con error: ${errores}`);
  console.log(`📂 En: ${outputDir}`);

  process.exit(errores > 0 ? 1 : 0);
}

process.on('unhandledRejection', (r) => { Logger.error(`${r}`); process.exit(1); });
process.on('SIGINT', () => { console.log(chalk.yellow('\n⚠️  Interrumpido')); process.exit(0); });

main();