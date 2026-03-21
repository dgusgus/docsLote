#!/usr/bin/env node

/**
 * generarAsistencia.ts — Genera registros de asistencia por grupo
 *
 * Lee personas desde Google Sheets, las agrupa por número de grupo,
 * y genera un Excel por grupo con las tablas URBANO y MÓVIL rellenas.
 *
 * El campo TIPO en el Google Sheets debe decir:
 *   - "URBANO" / "Urbano" / "OPERADOR DE TRANSMISION URBANO"   → tabla superior
 *   - "MOVIL"  / "Movil"  / "OPERADOR DE TRANSMISION MOVIL"    → tabla inferior
 *
 * USO:
 *   tsx generarAsistencia.ts                        → todos los grupos
 *   tsx generarAsistencia.ts --grupo 26             → solo el grupo 26
 *   tsx generarAsistencia.ts --fecha 19-10-2025     → fecha personalizada
 *   tsx generarAsistencia.ts --output ./salida      → carpeta de salida
 *   tsx generarAsistencia.ts --help                 → esta ayuda
 *
 * COLUMNAS ESPERADAS EN GOOGLE SHEETS (igual que el sistema principal):
 *   B=grupo, C=nombre, D=apellido1, E=apellido2, F=documento,
 *   G=telefono, H=email, I=cargo, J=fecha_inicio, K=fecha_fin, L=tipo
 *
 *   Si el campo TIPO está en otra columna, ajustá COLUMNA_TIPO más abajo.
 */

import path from 'path';
import { fileURLToPath } from 'url';
import chalk from 'chalk';
import { GoogleSheetsService } from './src/services/googleSheets.js';
import { GrupoExcelGenerator } from './src/services/grupoExcelGenerator.js';
import { Logger } from './src/utils/fileUtils.js';
import { CONFIG } from './src/config/settings.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

// ====================== CONFIGURACIÓN ======================
// Ruta a la plantilla de asistencia
const PLANTILLA_ASISTENCIA = path.join(__dirname, 'plantillas/Registro de asistencia OPERADORES.xlsx');

// Carpeta de salida por defecto
const OUTPUT_DEFAULT = path.join(__dirname, 'asistencia_generada');

// ====================== PARSEO DE ARGUMENTOS ======================
function parsearArgs() {
  const args = process.argv.slice(2);

  if (args.includes('--help') || args.includes('-h')) {
    mostrarAyuda();
    process.exit(0);
  }

  const get = (flag: string) => {
    const i = args.indexOf(flag);
    return i !== -1 && args[i + 1] ? args[i + 1] : null;
  };

  // Fecha por defecto: hoy en formato DD-MM-YYYY
  const hoy = new Date();
  const fechaDefault = `${String(hoy.getDate()).padStart(2,'0')}-${String(hoy.getMonth()+1).padStart(2,'0')}-${hoy.getFullYear()}`;

  return {
    grupo:     get('--grupo'),
    fecha:     get('--fecha') || fechaDefault,
    outputDir: get('--output') || OUTPUT_DEFAULT,
  };
}

// ====================== AYUDA ======================
function mostrarAyuda(): void {
  console.log(chalk.bold.cyan(`
╔══════════════════════════════════════════════════════════════╗
║  📋 GENERADOR DE REGISTROS DE ASISTENCIA POR GRUPO           ║
╚══════════════════════════════════════════════════════════════╝
`));
  console.log('USO:');
  console.log('  tsx generarAsistencia.ts                     Todos los grupos');
  console.log('  tsx generarAsistencia.ts --grupo 26          Solo grupo 26');
  console.log('  tsx generarAsistencia.ts --fecha 19-10-2025  Fecha personalizada');
  console.log('  tsx generarAsistencia.ts --output ./salida   Carpeta de salida');
  console.log('');
  console.log('COLUMNAS REQUERIDAS EN GOOGLE SHEETS:');
  console.log('  B=grupo, C=nombre, D=apellido1, E=apellido2');
  console.log('  F=documento, G=telefono, H=email, I=cargo');
  console.log('  J=fecha_inicio, K=fecha_fin, L=tipo (URBANO o MOVIL)');
  console.log('');
  console.log('VALORES ACEPTADOS PARA TIPO:');
  console.log('  URBANO, Urbano, OPERADOR DE TRANSMISION URBANO, urbano...');
  console.log('  MOVIL,  Movil,  OPERADOR DE TRANSMISION MOVIL,  movil...');
}

// ====================== MAIN ======================
async function main(): Promise<void> {
  console.log(chalk.bold.cyan('\n📋 GENERADOR DE REGISTROS DE ASISTENCIA\n'));

  const { grupo, fecha, outputDir } = parsearArgs();

  // 1. Conectar a Google Sheets y obtener personas
  const sheets = new GoogleSheetsService();
  let personas;

  try {
    await sheets.inicializar();
    personas = await sheets.obtenerPersonas(CONFIG.LIMITE_PERSONAS);
  } catch (err) {
    Logger.error(`Error conectando con Google Sheets: ${err}`);
    process.exit(1);
  }

  if (personas.length === 0) {
    Logger.warn('No se encontraron personas en el Google Sheets.');
    process.exit(0);
  }

  // 2. Filtrar por grupo si se especificó
  if (grupo) {
    personas = personas.filter(p => p.grupo?.toString().trim() === grupo.trim());
    if (personas.length === 0) {
      Logger.error(`No se encontraron personas en el grupo "${grupo}"`);
      process.exit(1);
    }
    Logger.info(`Filtrado a grupo ${grupo}: ${personas.length} persona(s)`);
  }

  // 3. Generar Excel por grupo
  const generador = new GrupoExcelGenerator();

  try {
    const resultados = await generador.generarTodosLosGrupos(
      PLANTILLA_ASISTENCIA,
      personas,
      fecha,
      outputDir
    );

    // 4. Resumen final
    Logger.separador();
    console.log(chalk.bold('\n📊 RESUMEN FINAL'));

    let totalExcel = 0;
    let totalErrores = 0;

    for (const r of resultados) {
      if (r.archivo) {
        console.log(`✅ Grupo ${r.grupo}: ${r.urbanos} urbanos + ${r.moviles} móviles → ${path.basename(r.archivo)}`);
        totalExcel++;
      } else {
        console.log(`❌ Grupo ${r.grupo}: ERROR`);
        totalErrores++;
      }
      if (r.errores.length > 0) {
        r.errores.forEach(e => console.log(`   ⚠️  ${e}`));
      }
    }

    console.log(`\n📁 Archivos generados: ${totalExcel}`);
    if (totalErrores > 0) console.log(`❌ Grupos con error:  ${totalErrores}`);
    console.log(`📂 Guardados en: ${outputDir}`);

    process.exit(totalErrores > 0 ? 1 : 0);

  } catch (err) {
    Logger.error(`Error generando asistencias: ${err}`);
    process.exit(1);
  }
}

process.on('unhandledRejection', (reason) => {
  Logger.error(`Error no manejado: ${reason}`);
  process.exit(1);
});

process.on('SIGINT', () => {
  console.log(chalk.yellow('\n⚠️  Proceso interrumpido'));
  process.exit(0);
});

main();