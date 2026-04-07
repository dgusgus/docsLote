#!/usr/bin/env node

/**
 * generarAsistencia.ts — Genera registros de asistencia SIREPRE 2026
 *
 * Fuentes de datos (mismo Google Spreadsheet):
 *   - Hoja "Hoja1"       → personas: grupo, nombre, apellidos, tipo (URBANO/MOVIL)
 *   - Hoja "Actividades" → eventos:  columna A=Fecha, B=Actividad, C=Ubicacion
 *
 * Flujo por cada grupo (según --tipo):
 *   solo-excel   → genera .xlsx por actividad, termina (sin PDF)
 *   solo-pdf     → genera .xlsx → convierte cada uno a .pdf → borra .xlsx
 *   solo-merged  → genera .xlsx → convierte a .pdf → mergea en uno → borra intermedios  [DEFAULT]
 *   todos        → genera .xlsx + .pdf individuales + .pdf mergeado (conserva todo)
 *
 * ESTRUCTURA DE SALIDA:
 *   asistencia_generada/
 *     grupo_1/
 *       11-03-2026_SIMULACRO.xlsx          ← si tipo incluye excel
 *       11-03-2026_SIMULACRO.pdf           ← si tipo incluye pdfs individuales
 *       grupo_1.pdf                        ← si tipo incluye merged
 *     grupo_26/
 *       ...
 *
 * USO:
 *   tsx generarAsistencia.ts
 *   tsx generarAsistencia.ts --grupo 26
 *   tsx generarAsistencia.ts --tipo solo-excel
 *   tsx generarAsistencia.ts --tipo solo-pdf
 *   tsx generarAsistencia.ts --tipo solo-merged      ← comportamiento por defecto
 *   tsx generarAsistencia.ts --tipo todos
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

// ====================== TIPOS DE SALIDA ======================

type TipoAsistencia = 'solo-excel' | 'solo-pdf' | 'solo-merged' | 'todos';

/**
 * Qué pasos se ejecutan según el tipo elegido:
 *
 *              xlsx   pdf_ind   merged
 * solo-excel    ✓       ✗         ✗
 * solo-pdf      ✓→del   ✓         ✗
 * solo-merged   ✓→del   ✓→del     ✓   (comportamiento original)
 * todos         ✓       ✓         ✓
 */
function debeGenerarExcel(tipo: TipoAsistencia)  { return true; } // siempre se necesita como paso 1
function debeBorrarExcel(tipo: TipoAsistencia)   { return tipo === 'solo-pdf' || tipo === 'solo-merged'; }
function debeGenerarPDF(tipo: TipoAsistencia)    { return tipo !== 'solo-excel'; }
function debeMergear(tipo: TipoAsistencia)        { return tipo === 'solo-merged' || tipo === 'todos'; }
function debeBorrarPDFInd(tipo: TipoAsistencia)  { return tipo === 'solo-merged'; }

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
  moviles:      Persona[],
  urbanos:      Persona[],
  grupo:        string,
  actividad:    Actividad,
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
    setTimeout(() => { proc.kill(); reject(new Error('Timeout convirtiendo a PDF')); }, 60_000);
  });
}

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

async function eliminarArchivos(rutas: string[]): Promise<void> {
  await Promise.allSettled(rutas.map(r => fs.unlink(r)));
}

// ====================== ARGUMENTOS ======================

interface Args {
  grupo:     string | null;
  outputDir: string;
  tipo:      TipoAsistencia;
}

function parsearArgs(): Args {
  const argv = process.argv.slice(2);
  if (argv.includes('--help') || argv.includes('-h')) { mostrarAyuda(); process.exit(0); }

  const get = (flag: string) => {
    const i = argv.indexOf(flag);
    return i !== -1 && argv[i + 1] ? argv[i + 1] : null;
  };

  // Validar --tipo
  const tiposValidos: TipoAsistencia[] = ['solo-excel', 'solo-pdf', 'solo-merged', 'todos'];
  const tipoRaw = get('--tipo');
  let tipo: TipoAsistencia = 'solo-merged'; // default igual al comportamiento original

  if (tipoRaw) {
    if (!tiposValidos.includes(tipoRaw as TipoAsistencia)) {
      console.error(chalk.red(`❌ Tipo inválido: "${tipoRaw}". Opciones: ${tiposValidos.join(', ')}`));
      process.exit(1);
    }
    tipo = tipoRaw as TipoAsistencia;
  }

  return {
    grupo:     get('--grupo'),
    outputDir: get('--output') || OUTPUT_DEFAULT,
    tipo,
  };
}

function mostrarAyuda(): void {
  console.log(chalk.bold.cyan(`
╔══════════════════════════════════════════════════════════════╗
║  📋 GENERADOR DE ASISTENCIA SIREPRE 2026                     ║
╚══════════════════════════════════════════════════════════════╝
`));
  console.log('Genera registros de asistencia por grupo y actividad.');
  console.log('');
  console.log(chalk.bold('USO:'));
  console.log('  tsx generarAsistencia.ts [opciones]');
  console.log('');
  console.log(chalk.bold('OPCIONES:'));
  console.log('  --grupo  <n>           Procesar solo ese grupo (ej: --grupo 26)');
  console.log('  --output <carpeta>     Carpeta de salida (default: asistencia_generada/)');
  console.log('  --tipo   <modo>        Qué archivos generar (ver abajo)');
  console.log('  --help                 Muestra esta ayuda');
  console.log('');
  console.log(chalk.bold('MODOS DE --tipo:'));
  console.log(`  ${chalk.cyan('solo-excel')}    → Solo los .xlsx por fecha (sin PDF)`);
  console.log(`  ${chalk.cyan('solo-pdf')}      → Un .pdf por fecha (borra los .xlsx)`);
  console.log(`  ${chalk.cyan('solo-merged')}   → Un único .pdf por grupo con todas las fechas  ${chalk.gray('[DEFAULT]')}`);
  console.log(`  ${chalk.cyan('todos')}         → .xlsx + .pdf por fecha + .pdf unificado`);
  console.log('');
  console.log(chalk.bold('EJEMPLOS:'));
  console.log('  tsx generarAsistencia.ts');
  console.log('  tsx generarAsistencia.ts --grupo 26 --tipo solo-excel');
  console.log('  tsx generarAsistencia.ts --tipo todos --output ./mi_salida');
  console.log('');
  console.log(chalk.bold('ESTRUCTURA DE SALIDA (--tipo todos):'));
  console.log('  asistencia_generada/');
  console.log('    grupo_26/');
  console.log('      11-03-2026_SIMULACRO.xlsx');
  console.log('      11-03-2026_SIMULACRO.pdf');
  console.log('      25-03-2026_CAPACITACION.xlsx');
  console.log('      25-03-2026_CAPACITACION.pdf');
  console.log('      grupo_26.pdf   ← todas las fechas unidas');
}

// ====================== MAIN ======================

async function main(): Promise<void> {
  console.log(chalk.bold.cyan('\n📋 GENERADOR DE ASISTENCIA SIREPRE 2026\n'));

  const { grupo, outputDir, tipo } = parsearArgs();

  // Mostrar configuración activa
  Logger.info(`Modo:       ${chalk.cyan(tipo)}`);
  if (grupo) Logger.info(`Grupo:      ${grupo}`);
  Logger.info(`Salida:     ${outputDir}`);

  // Explicar qué va a generar
  const explicacion: Record<TipoAsistencia, string> = {
    'solo-excel':  'Solo archivos Excel (.xlsx) por fecha',
    'solo-pdf':    'Un PDF por fecha (sin conservar Excel)',
    'solo-merged': 'Un PDF unificado por grupo con todas las fechas',
    'todos':       'Excel + PDF por fecha + PDF unificado por grupo',
  };
  Logger.info(`Genera:     ${explicacion[tipo]}\n`);

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

  // Ordenar actividades cronológicamente
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

  // 4. Verificar LibreOffice si se necesita PDF
  if (debeGenerarPDF(tipo)) {
    try {
      await fs.access(CONFIG.SOFFICE_PATH);
    } catch {
      Logger.error(`LibreOffice no encontrado en: ${CONFIG.SOFFICE_PATH}`);
      Logger.info('   Instalá LibreOffice o usá --tipo solo-excel para evitar la conversión PDF');
      process.exit(1);
    }
  }

  // 5. Procesar grupo por grupo
  let exitosos = 0;
  let errores  = 0;

  for (const [g, miembros] of porGrupo) {
    const moviles = miembros.filter(p => normalizarTipo(p.tipo) === 'MOVIL');
    const urbanos = miembros.filter(p => normalizarTipo(p.tipo) === 'URBANO');
    const sinTipo = miembros.filter(p => normalizarTipo(p.tipo) === null);

    console.log(chalk.bold(`\n📁 grupo_${g}/  (${moviles.length} móvil + ${urbanos.length} urbano${sinTipo.length ? ` + ${sinTipo.length} sin tipo` : ''})`));
    if (sinTipo.length) sinTipo.forEach(p => Logger.warn(`   ⚠️  Sin tipo: ${p.nombre} ${p.apellido1} — "${p.tipo}"`));

    const carpetaGrupo = path.join(outputDir, `grupo_${g.replace(/\s+/g, '_')}`);
    await crearDirectorioSeguro(carpetaGrupo);

    const rutasXlsx:   string[] = [];
    const rutasPdfInd: string[] = [];
    let errorEnGrupo = false;

    // 5a. Generar xlsx por actividad
    for (const act of actividades) {
      const etiqueta = `${act.fecha} — ${act.actividad.substring(0, 35)}`;
      try {
        const rutaXlsx = await generarXlsx(moviles, urbanos, g, act, carpetaGrupo);
        rutasXlsx.push(rutaXlsx);
        Logger.success(`   📊 ${etiqueta} → xlsx`);
      } catch (err) {
        Logger.error(`   ❌ ${etiqueta} (excel): ${err}`);
        errorEnGrupo = true;
      }
    }

    if (rutasXlsx.length === 0) {
      Logger.error(`   Sin Excel generados para grupo ${g}`);
      errores++;
      continue;
    }

    // Si solo queremos Excel, terminamos aquí para este grupo
    if (!debeGenerarPDF(tipo)) {
      Logger.success(`   ✅ ${rutasXlsx.length} archivo(s) Excel generado(s)`);
      exitosos++;
      continue;
    }

    // 5b. Convertir xlsx → pdf
    for (const rutaXlsx of rutasXlsx) {
      const etiqueta = path.basename(rutaXlsx, '.xlsx');
      try {
        process.stdout.write(`   🔄 ${etiqueta} → PDF...`);
        const rutaPdf = await xlsxToPdf(rutaXlsx, carpetaGrupo);
        rutasPdfInd.push(rutaPdf);
        process.stdout.write(' ✅\n');
      } catch (err) {
        process.stdout.write('\n');
        Logger.error(`   ❌ ${etiqueta} (pdf): ${err}`);
        errorEnGrupo = true;
      }
    }

    // Borrar xlsx si el tipo lo pide
    if (debeBorrarExcel(tipo)) {
      await eliminarArchivos(rutasXlsx);
    }

    // Si solo queremos pdfs individuales, terminamos aquí
    if (!debeMergear(tipo)) {
      Logger.success(`   ✅ ${rutasPdfInd.length} PDF(s) individual(es) generado(s)`);
      if (!errorEnGrupo) exitosos++;
      continue;
    }

    // 5c. Mergear todos los PDFs en uno
    if (rutasPdfInd.length === 0) {
      Logger.error(`   Sin PDFs individuales para mergear en grupo ${g}`);
      errores++;
      continue;
    }

    const rutaPdfFinal = path.join(carpetaGrupo, `grupo_${g.replace(/\s+/g, '_')}.pdf`);
    try {
      process.stdout.write(`   📎 Uniendo ${rutasPdfInd.length} PDF(s) → grupo_${g}.pdf...`);
      await mergePdfs(rutasPdfInd, rutaPdfFinal);
      process.stdout.write(' ✅\n');
      Logger.success(`   📄 grupo_${g}.pdf  (${rutasPdfInd.length} fechas)`);
      if (!errorEnGrupo) exitosos++;
    } catch (err) {
      process.stdout.write('\n');
      Logger.error(`   ❌ Error mergeando grupo ${g}: ${err}`);
      errores++;
      errorEnGrupo = true;
    }

    // Borrar PDFs individuales si el tipo lo pide (solo-merged)
    if (debeBorrarPDFInd(tipo)) {
      await eliminarArchivos(rutasPdfInd);
    }
  }

  // 6. Resumen
  Logger.separador();
  console.log(chalk.bold('\n📊 RESUMEN'));
  console.log(`✅ Grupos completados: ${exitosos} / ${porGrupo.size}`);
  if (errores) console.log(`❌ Grupos con error: ${errores}`);
  console.log(`📂 Archivos en: ${outputDir}`);
  console.log(`📋 Tipo generado: ${tipo}`);

  process.exit(errores > 0 ? 1 : 0);
}

process.on('unhandledRejection', (r) => { Logger.error(`${r}`); process.exit(1); });
process.on('SIGINT', () => { console.log(chalk.yellow('\n⚠️  Interrumpido')); process.exit(0); });

main();