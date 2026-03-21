#!/usr/bin/env node

/**
 * generarAsistencia.ts — Genera registros de asistencia por grupo
 *
 * Lee fecha_inicio y fecha_fin de cada persona desde Google Sheets.
 * Por cada grupo genera un archivo por día y opcionalmente un PDF unificado.
 *
 * ESTRUCTURA DE SALIDA:
 *   asistencia_generada/
 *     grupo_1/
 *       01-03-2026.xlsx
 *       01-03-2026.pdf
 *       02-03-2026.pdf
 *       grupo_1_COMPLETO.pdf   ← todos los días unidos en orden
 *     grupo_26/
 *       ...
 *       grupo_26_COMPLETO.pdf
 *
 * USO:
 *   tsx generarAsistencia.ts                        → Excel + PDF por día + PDF unificado
 *   tsx generarAsistencia.ts --tipo solo-excel      → solo .xlsx por día
 *   tsx generarAsistencia.ts --tipo solo-pdf        → solo .pdf por día + PDF unificado
 *   tsx generarAsistencia.ts --tipo ambos           → .xlsx y .pdf por día + PDF unificado
 *   tsx generarAsistencia.ts --sin-unificar         → no genera el PDF unificado
 *   tsx generarAsistencia.ts --grupo 26
 *   tsx generarAsistencia.ts --output ./mi_salida
 *   tsx generarAsistencia.ts --help
 *
 * FORMATO DE FECHAS EN EL SHEET: DD/MM/YYYY o DD-MM-YYYY
 */

import path from 'path';
import fs from 'fs/promises';
import { fileURLToPath } from 'url';
import chalk from 'chalk';
import ExcelJS from 'exceljs';
import { PDFDocument } from 'pdf-lib';
import { GoogleSheetsService } from './src/services/googleSheets.js';
import { PDFConverter } from './src/services/pdfConverter.js';
import { Logger, crearDirectorioSeguro } from './src/utils/fileUtils.js';
import { CONFIG } from './src/config/settings.js';
import { Persona } from './src/types/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

const PLANTILLA_ASISTENCIA = path.join(__dirname, 'plantillas/Registro de asistencia OPERADORES.xlsx');
const OUTPUT_DEFAULT        = path.join(__dirname, 'asistencia_generada');

type TipoSalida = 'solo-excel' | 'solo-pdf' | 'ambos';

// ====================== UTILIDADES DE FECHA ======================

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

function rangoFechas(inicio: Date, fin: Date): string[] {
  const fechas: string[] = [];
  const cur = new Date(inicio);
  while (cur <= fin) {
    fechas.push(formatearFecha(cur));
    cur.setDate(cur.getDate() + 1);
  }
  return fechas;
}

function estaActiva(persona: Persona, fecha: string): boolean {
  const inicio = parsearFechaSheet(persona.fecha_inicio);
  const fin    = parsearFechaSheet(persona.fecha_fin);
  if (!inicio || !fin) return false;
  const d = parsearFechaSheet(fecha)!;
  return d >= inicio && d <= fin;
}

// ====================== UNIFICADOR DE PDFs ======================

/**
 * Une una lista de archivos PDF en un solo PDF, en el orden dado.
 * Usa pdf-lib que ya está instalado en el proyecto.
 */
async function unificarPDFs(rutasPDF: string[], rutaSalida: string): Promise<void> {
  const pdfUnificado = await PDFDocument.create();

  for (const rutaPDF of rutasPDF) {
    const bytes    = await fs.readFile(rutaPDF);
    const pdfOrigen = await PDFDocument.load(bytes);
    const paginas  = await pdfUnificado.copyPages(pdfOrigen, pdfOrigen.getPageIndices());
    paginas.forEach(pagina => pdfUnificado.addPage(pagina));
  }

  const bytesFinales = await pdfUnificado.save();
  await fs.writeFile(rutaSalida, bytesFinales);
}

// ====================== ARGUMENTOS ======================

interface Args {
  grupo:      string | null;
  outputDir:  string;
  tipo:       TipoSalida;
  unificar:   boolean;
}

function parsearArgs(): Args {
  const argv = process.argv.slice(2);
  if (argv.includes('--help') || argv.includes('-h')) { mostrarAyuda(); process.exit(0); }
  const get = (flag: string) => { const i = argv.indexOf(flag); return i !== -1 && argv[i + 1] ? argv[i + 1] : null; };

  const tipoRaw = get('--tipo') || 'ambos';
  const tiposValidos: TipoSalida[] = ['solo-excel', 'solo-pdf', 'ambos'];
  if (!tiposValidos.includes(tipoRaw as TipoSalida)) {
    Logger.error(`--tipo inválido: "${tipoRaw}". Opciones: solo-excel, solo-pdf, ambos`);
    process.exit(1);
  }

  return {
    grupo:     get('--grupo'),
    outputDir: get('--output') || OUTPUT_DEFAULT,
    tipo:      tipoRaw as TipoSalida,
    unificar:  !argv.includes('--sin-unificar'),
  };
}

function mostrarAyuda(): void {
  console.log(chalk.bold.cyan(`
╔══════════════════════════════════════════════════════════════╗
║  📋 GENERADOR DE REGISTROS DE ASISTENCIA POR GRUPO           ║
╚══════════════════════════════════════════════════════════════╝
`));
  console.log('Las fechas se leen desde Google Sheets (columnas J=fecha_inicio, K=fecha_fin).');
  console.log('');
  console.log('USO:');
  console.log('  tsx generarAsistencia.ts                        Excel + PDF + unificado');
  console.log('  tsx generarAsistencia.ts --tipo solo-excel      Solo .xlsx por día');
  console.log('  tsx generarAsistencia.ts --tipo solo-pdf        Solo .pdf por día + unificado');
  console.log('  tsx generarAsistencia.ts --tipo ambos           .xlsx + .pdf + unificado');
  console.log('  tsx generarAsistencia.ts --sin-unificar         Sin generar el PDF unificado');
  console.log('  tsx generarAsistencia.ts --grupo 26             Filtrar por grupo');
  console.log('  tsx generarAsistencia.ts --output ./mi_salida   Carpeta de salida');
  console.log('');
  console.log('ESTRUCTURA DE SALIDA:');
  console.log('  asistencia_generada/');
  console.log('    grupo_1/');
  console.log('      01-03-2026.xlsx');
  console.log('      01-03-2026.pdf');
  console.log('      02-03-2026.pdf');
  console.log('      grupo_1_COMPLETO.pdf   ← todos los días en un solo PDF');
  console.log('    grupo_26/');
  console.log('      grupo_26_COMPLETO.pdf');
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

/**
 * Genera el Excel y lo convierte a PDF si corresponde.
 * Devuelve la ruta del PDF generado (o null si no se generó).
 */
async function generarArchivo(
  plantillaPath: string,
  personasDelDia: Persona[],
  grupo: string,
  fecha: string,
  carpetaGrupo: string,
  tipo: TipoSalida,
  pdfConverter: PDFConverter
): Promise<string | null> {
  const urbanos = personasDelDia.filter(p => normalizarTipo(p.tipo) === 'URBANO');
  const moviles = personasDelDia.filter(p => normalizarTipo(p.tipo) === 'MOVIL');

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(plantillaPath);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('La plantilla no tiene hojas de trabajo');

  ws.getCell('E8').value = `FECHA: ${fecha}${' '.repeat(50)}GRUPO: ${grupo}`;
  rellenarTabla(ws, urbanos.slice(0, MAX_OPERADORES), FILA_INICIO_URBANO);
  rellenarTabla(ws, moviles.slice(0, MAX_OPERADORES), FILA_INICIO_MOVIL);

  const rutaXlsx = path.join(carpetaGrupo, `${fecha}.xlsx`);
  await wb.xlsx.writeFile(rutaXlsx);

  // Convertir a PDF
  let rutaPDF: string | null = null;
  if (tipo === 'solo-pdf' || tipo === 'ambos') {
    await pdfConverter.convertirAPdf(rutaXlsx, carpetaGrupo);
    rutaPDF = path.join(carpetaGrupo, `${fecha}.pdf`);

    // Borrar xlsx si solo se quiere PDF
    if (tipo === 'solo-pdf') {
      await fs.unlink(rutaXlsx);
    }
  }

  return rutaPDF;
}

// ====================== MAIN ======================

async function main(): Promise<void> {
  console.log(chalk.bold.cyan('\n📋 GENERADOR DE REGISTROS DE ASISTENCIA\n'));

  const { grupo, outputDir, tipo, unificar } = parsearArgs();

  Logger.info('Fechas:    leídas desde Google Sheets (fecha_inicio / fecha_fin por persona)');
  Logger.info(`Formato:   ${tipo}`);
  Logger.info(`Unificar:  ${unificar && tipo !== 'solo-excel' ? 'sí → grupo_XX_COMPLETO.pdf' : 'no'}`);
  if (grupo) Logger.info(`Grupo:     ${grupo}`);
  Logger.info(`Salida:    ${outputDir}\n`);

  // Verificar LibreOffice si se necesita PDF
  const pdfConverter = PDFConverter.obtenerInstancia();
  if (tipo !== 'solo-excel') {
    const libreOfficeOk = await pdfConverter.verificarLibreOffice();
    if (!libreOfficeOk) {
      Logger.error('LibreOffice no encontrado. Necesario para generar PDFs.');
      Logger.info('Usá --tipo solo-excel para omitir la conversión, o instalá LibreOffice.');
      process.exit(1);
    }
    Logger.info('LibreOffice: ✅\n');
  }

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

  // 2. Filtrar por grupo
  if (grupo) {
    personas = personas.filter(p => p.grupo?.toString().trim() === grupo.trim());
    if (personas.length === 0) { Logger.error(`No hay personas en el grupo "${grupo}"`); process.exit(1); }
  }

  // 3. Excluir personas sin fechas válidas
  const sinFechas = personas.filter(p => !parsearFechaSheet(p.fecha_inicio) || !parsearFechaSheet(p.fecha_fin));
  if (sinFechas.length > 0) {
    Logger.warn(`${sinFechas.length} persona(s) sin fechas válidas (serán ignoradas):`);
    sinFechas.forEach(p =>
      Logger.warn(`   • ${p.nombre} ${p.apellido1} (grupo ${p.grupo}) — inicio: "${p.fecha_inicio}" fin: "${p.fecha_fin}"`)
    );
    console.log('');
  }
  personas = personas.filter(p => parsearFechaSheet(p.fecha_inicio) && parsearFechaSheet(p.fecha_fin));
  if (personas.length === 0) { Logger.error('Ninguna persona tiene fechas válidas.'); process.exit(1); }

  // 4. Agrupar por grupo
  const porGrupo = new Map<string, Persona[]>();
  for (const p of personas) {
    const g = (p.grupo || 'SIN_GRUPO').toString().trim();
    if (!porGrupo.has(g)) porGrupo.set(g, []);
    porGrupo.get(g)!.push(p);
  }

  // 5. Generar archivos
  let totalArchivos = 0;
  let exitosos = 0;
  let errores  = 0;

  for (const [g, miembros] of porGrupo) {
    // Calcular días únicos del grupo
    const fechasSet = new Set<string>();
    for (const p of miembros) {
      rangoFechas(parsearFechaSheet(p.fecha_inicio)!, parsearFechaSheet(p.fecha_fin)!)
        .forEach(f => fechasSet.add(f));
    }
    const fechasOrdenadas = [...fechasSet].sort((a, b) =>
      parsearFechaSheet(a)!.getTime() - parsearFechaSheet(b)!.getTime()
    );

    totalArchivos += fechasOrdenadas.length;
    const nU = miembros.filter(p => normalizarTipo(p.tipo) === 'URBANO').length;
    const nM = miembros.filter(p => normalizarTipo(p.tipo) === 'MOVIL').length;
    console.log(chalk.bold(`\n📁 grupo_${g}/  (${miembros.length} personas: ${nU}U + ${nM}M | ${fechasOrdenadas.length} días)`));

    const carpetaGrupo = path.join(outputDir, `grupo_${g.replace(/\s+/g, '_')}`);
    await crearDirectorioSeguro(carpetaGrupo);

    // PDFs generados en orden cronológico (para unificar después)
    const pdfsDelGrupo: string[] = [];

    for (const fecha of fechasOrdenadas) {
      const activos = miembros.filter(p => estaActiva(p, fecha));
      const u = activos.filter(p => normalizarTipo(p.tipo) === 'URBANO').length;
      const m = activos.filter(p => normalizarTipo(p.tipo) === 'MOVIL').length;
      try {
        const rutaPDF = await generarArchivo(
          PLANTILLA_ASISTENCIA, activos, g, fecha, carpetaGrupo, tipo, pdfConverter
        );
        if (rutaPDF) pdfsDelGrupo.push(rutaPDF);
        const iconos = [tipo !== 'solo-pdf' ? '📊' : '', rutaPDF ? '📄' : ''].filter(Boolean).join(' ');
        console.log(`   ✅ ${fecha}  ${iconos}  (${activos.length} activos: ${u}U + ${m}M)`);
        exitosos++;
      } catch (err) {
        console.log(`   ❌ ${fecha}  — ${err}`);
        errores++;
      }
    }

    // Unificar PDFs del grupo en un solo archivo
    if (unificar && pdfsDelGrupo.length > 1) {
      const rutaCompleto = path.join(carpetaGrupo, `grupo_${g}_COMPLETO.pdf`);
      try {
        await unificarPDFs(pdfsDelGrupo, rutaCompleto);
        console.log(chalk.green(`   📎 grupo_${g}_COMPLETO.pdf  (${pdfsDelGrupo.length} días unidos)`));
      } catch (err) {
        Logger.warn(`   No se pudo unificar el PDF del grupo ${g}: ${err}`);
      }
    } else if (unificar && pdfsDelGrupo.length === 1) {
      // Solo un día — no tiene sentido unificar, ya existe el PDF individual
      console.log(chalk.gray(`   ℹ️  Un solo día, no se genera PDF unificado`));
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