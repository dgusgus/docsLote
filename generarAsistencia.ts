#!/usr/bin/env node

/**
 * generarAsistencia.ts — Genera registros de asistencia por grupo
 *
 * Lee fecha_inicio y fecha_fin de cada persona desde Google Sheets.
 * Por cada grupo, genera un archivo .xlsx por cada día en el que
 * al menos una persona de ese grupo esté activa (dentro de su rango).
 *
 * ESTRUCTURA DE SALIDA:
 *   asistencia_generada/
 *     grupo_1/
 *       01-03-2026.xlsx   ← solo las personas activas ese día
 *       02-03-2026.xlsx
 *       ...
 *     grupo_26/
 *       01-03-2026.xlsx
 *
 * USO:
 *   tsx generarAsistencia.ts                  → todos los grupos, fechas desde el Sheet
 *   tsx generarAsistencia.ts --grupo 26       → solo grupo 26
 *   tsx generarAsistencia.ts --output ./salida
 *   tsx generarAsistencia.ts --help
 *
 * FORMATO DE FECHAS EN EL SHEET: DD/MM/YYYY o DD-MM-YYYY
 */

import path from 'path';
import { fileURLToPath } from 'url';
import chalk from 'chalk';
import ExcelJS from 'exceljs';
import { GoogleSheetsService } from './src/services/googleSheets.js';
import { Logger, crearDirectorioSeguro } from './src/utils/fileUtils.js';
import { CONFIG } from './src/config/settings.js';
import { Persona } from './src/types/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

const PLANTILLA_ASISTENCIA = path.join(__dirname, 'plantillas/Registro de asistencia OPERADORES.xlsx');
const OUTPUT_DEFAULT        = path.join(__dirname, 'asistencia_generada');

// ====================== UTILIDADES DE FECHA ======================

/**
 * Parsea DD/MM/YYYY o DD-MM-YYYY → Date.
 * Devuelve null si el formato es inválido o está vacío.
 */
function parsearFechaSheet(str: string | undefined): Date | null {
  if (!str || !str.trim()) return null;
  const normalizado = str.trim().replace(/\//g, '-');
  const partes = normalizado.split('-');
  if (partes.length !== 3) return null;
  const [dia, mes, anio] = partes.map(Number);
  if (isNaN(dia) || isNaN(mes) || isNaN(anio)) return null;
  const d = new Date(anio, mes - 1, dia);
  return isNaN(d.getTime()) ? null : d;
}

function formatearFecha(d: Date): string {
  return [
    String(d.getDate()).padStart(2, '0'),
    String(d.getMonth() + 1).padStart(2, '0'),
    d.getFullYear(),
  ].join('-');
}

/** Devuelve todas las fechas DD-MM-YYYY entre inicio y fin, inclusive. */
function rangoFechas(inicio: Date, fin: Date): string[] {
  const fechas: string[] = [];
  const cur = new Date(inicio);
  while (cur <= fin) {
    fechas.push(formatearFecha(cur));
    cur.setDate(cur.getDate() + 1);
  }
  return fechas;
}

/** True si la persona está activa en esa fecha. */
function estaActiva(persona: Persona, fecha: string): boolean {
  const inicio = parsearFechaSheet(persona.fecha_inicio);
  const fin    = parsearFechaSheet(persona.fecha_fin);
  if (!inicio || !fin) return false;
  const d = parsearFechaSheet(fecha)!;
  return d >= inicio && d <= fin;
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
║  📋 GENERADOR DE REGISTROS DE ASISTENCIA POR GRUPO           ║
╚══════════════════════════════════════════════════════════════╝
`));
  console.log('Las fechas se leen directamente desde el Google Sheets (columnas J y K).');
  console.log('Se genera un archivo por cada día en que al menos una persona esté activa.');
  console.log('');
  console.log('USO:');
  console.log('  tsx generarAsistencia.ts');
  console.log('  tsx generarAsistencia.ts --grupo 26');
  console.log('  tsx generarAsistencia.ts --output ./mi_salida');
  console.log('');
  console.log('ESTRUCTURA DE SALIDA:');
  console.log('  asistencia_generada/');
  console.log('    grupo_1/');
  console.log('      01-03-2026.xlsx');
  console.log('      02-03-2026.xlsx');
  console.log('    grupo_26/');
  console.log('      01-03-2026.xlsx');
  console.log('');
  console.log('FORMATO DE FECHAS EN EL SHEET: DD/MM/YYYY o DD-MM-YYYY');
}

// ====================== LÓGICA EXCEL ======================

const FILA_INICIO_URBANO = 11;
const FILA_INICIO_MOVIL  = 33;
const MAX_OPERADORES     = 10;

function normalizarTipo(tipo: string | undefined): 'URBANO' | 'MOVIL' | null {
  if (!tipo) return null;
  const t = tipo.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  if (t.includes('URBAN')) return 'URBANO';
  if (t.includes('MOVIL') || t.includes('MOBIL')) return 'MOVIL';
  return null;
}

function rellenarTabla(ws: ExcelJS.Worksheet, personas: Persona[], filaInicio: number): void {
  for (let i = 0; i < MAX_OPERADORES; i++) {
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
  plantillaPath: string,
  personasDelDia: Persona[],
  grupo: string,
  fecha: string,
  outputDir: string
): Promise<void> {
  const urbanos = personasDelDia.filter(p => normalizarTipo(p.tipo) === 'URBANO');
  const moviles = personasDelDia.filter(p => normalizarTipo(p.tipo) === 'MOVIL');

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(plantillaPath);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('La plantilla no tiene hojas de trabajo');

  // E8 tiene fecha y grupo — E30 ya tiene =E8 en la plantilla
  ws.getCell('E8').value = `FECHA: ${fecha}${' '.repeat(50)}GRUPO: ${grupo}`;

  rellenarTabla(ws, urbanos.slice(0, MAX_OPERADORES), FILA_INICIO_URBANO);
  rellenarTabla(ws, moviles.slice(0, MAX_OPERADORES), FILA_INICIO_MOVIL);

  const carpetaGrupo = path.join(outputDir, `grupo_${grupo.replace(/\s+/g, '_')}`);
  await crearDirectorioSeguro(carpetaGrupo);
  await wb.xlsx.writeFile(path.join(carpetaGrupo, `${fecha}.xlsx`));
}

// ====================== MAIN ======================

async function main(): Promise<void> {
  console.log(chalk.bold.cyan('\n📋 GENERADOR DE REGISTROS DE ASISTENCIA\n'));

  const { grupo, outputDir } = parsearArgs();
  Logger.info('Fechas:  leídas desde Google Sheets (fecha_inicio / fecha_fin por persona)');
  if (grupo) Logger.info(`Grupo:   ${grupo}`);
  Logger.info(`Salida:  ${outputDir}\n`);

  // 1. Obtener personas
  const sheets = new GoogleSheetsService();
  let personas: Persona[];
  try {
    await sheets.inicializar();
    personas = await sheets.obtenerPersonas(CONFIG.LIMITE_PERSONAS);
  } catch (err) {
    Logger.error(`Error conectando con Google Sheets: ${err}`);
    process.exit(1);
  }

  if (personas.length === 0) { Logger.warn('No se encontraron personas.'); process.exit(0); }

  // 2. Filtrar por grupo si se indicó
  if (grupo) {
    personas = personas.filter(p => p.grupo?.toString().trim() === grupo.trim());
    if (personas.length === 0) { Logger.error(`No hay personas en el grupo "${grupo}"`); process.exit(1); }
  }

  // 3. Advertir y excluir personas sin fechas válidas
  const sinFechas = personas.filter(p => !parsearFechaSheet(p.fecha_inicio) || !parsearFechaSheet(p.fecha_fin));
  if (sinFechas.length > 0) {
    Logger.warn(`${sinFechas.length} persona(s) sin fechas válidas en el Sheet (serán ignoradas):`);
    sinFechas.forEach(p =>
      Logger.warn(`   • ${p.nombre} ${p.apellido1} (grupo ${p.grupo}) — inicio: "${p.fecha_inicio}" fin: "${p.fecha_fin}"`)
    );
    console.log('');
  }
  personas = personas.filter(p => parsearFechaSheet(p.fecha_inicio) && parsearFechaSheet(p.fecha_fin));

  if (personas.length === 0) { Logger.error('Ninguna persona tiene fechas válidas.'); process.exit(1); }

  // 4. Agrupar por número de grupo
  const porGrupo = new Map<string, Persona[]>();
  for (const p of personas) {
    const g = (p.grupo || 'SIN_GRUPO').toString().trim();
    if (!porGrupo.has(g)) porGrupo.set(g, []);
    porGrupo.get(g)!.push(p);
  }

  // 5. Por cada grupo: calcular días únicos → generar un archivo por día
  let totalArchivos = 0;
  let exitosos = 0;
  let errores  = 0;

  for (const [g, miembros] of porGrupo) {

    // Unión de todos los días en los que al menos alguien está activo
    const fechasSet = new Set<string>();
    for (const p of miembros) {
      const inicio = parsearFechaSheet(p.fecha_inicio)!;
      const fin    = parsearFechaSheet(p.fecha_fin)!;
      rangoFechas(inicio, fin).forEach(f => fechasSet.add(f));
    }
    const fechasOrdenadas = [...fechasSet].sort(); // orden cronológico (YYYY es el último segmento, sort alfabético no funciona)
    // Reordenar correctamente por fecha real
    fechasOrdenadas.sort((a, b) => {
      const toMs = (s: string) => parsearFechaSheet(s)!.getTime();
      return toMs(a) - toMs(b);
    });

    totalArchivos += fechasOrdenadas.length;

    const nU = miembros.filter(p => normalizarTipo(p.tipo) === 'URBANO').length;
    const nM = miembros.filter(p => normalizarTipo(p.tipo) === 'MOVIL').length;
    console.log(chalk.bold(`\n📁 grupo_${g}/  (${miembros.length} personas: ${nU}U + ${nM}M | ${fechasOrdenadas.length} días)`));

    for (const fecha of fechasOrdenadas) {
      // Solo las personas activas ese día específico
      const activos = miembros.filter(p => estaActiva(p, fecha));
      const u = activos.filter(p => normalizarTipo(p.tipo) === 'URBANO').length;
      const m = activos.filter(p => normalizarTipo(p.tipo) === 'MOVIL').length;
      try {
        await generarArchivo(PLANTILLA_ASISTENCIA, activos, g, fecha, outputDir);
        console.log(`   ✅ ${fecha}.xlsx  (${activos.length} activos: ${u}U + ${m}M)`);
        exitosos++;
      } catch (err) {
        console.log(`   ❌ ${fecha}.xlsx  — ${err}`);
        errores++;
      }
    }
  }

  // 6. Resumen
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