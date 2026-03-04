#!/usr/bin/env node

/**
 * generarQR.ts — Script independiente para generar QR codes desde Google Sheets
 *
 * USO:
 *   tsx generarQR.ts                          → todos (hasta límite configurado)
 *   tsx generarQR.ts --rango 1 50             → solo personas del índice 1 al 50
 *   tsx generarQR.ts --nombre "Juan Perez"    → solo esa persona
 *   tsx generarQR.ts --output ./mis_qrs       → directorio de salida personalizado
 *   tsx generarQR.ts --help                   → muestra esta ayuda
 *
 * SALIDA:
 *   qrs_generados/
 *     1_grupo_apellido1_apellido2_nombre/
 *       qr.png          ← imagen PNG 400×400 px
 *     2_grupo_...
 *       qr.png
 *     ...
 *
 * CONTENIDO del QR (texto plano):
 *   GRUPO: ...
 *   NOMBRE: ...
 *   APELLIDO 1: ...
 *   APELLIDO 2: ...
 *   DOCUMENTO: ...
 *   CARGO: ...
 *   TELEFONO: ...
 *   EMAIL: ...
 *   FECHA INICIO: ...
 *   FECHA FIN: ...
 */

import path from 'path';
import { fileURLToPath } from 'url';
import chalk from 'chalk';
import { GoogleSheetsService } from './src/services/googleSheets.js';
import { QRGenerator } from './src/services/qrGenerator.js';
import { CONFIG } from './src/config/settings.js';
import { Logger } from './src/utils/fileUtils.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

// ====================== PARSEO DE ARGUMENTOS ======================
function parsearArgs(): {
  modo: 'todos' | 'rango' | 'nombre';
  inicio?: number;
  fin?: number;
  nombre?: string;
  outputDir: string;
} {
  const args = process.argv.slice(2);

  if (args.includes('--help') || args.includes('-h')) {
    mostrarAyuda();
    process.exit(0);
  }

  const outputIdx = args.indexOf('--output');
  const outputDir = outputIdx !== -1 && args[outputIdx + 1]
    ? args[outputIdx + 1]
    : path.join(__dirname, 'qrs_generados');

  // Modo: persona por nombre
  const nombreIdx = args.indexOf('--nombre');
  if (nombreIdx !== -1 && args[nombreIdx + 1]) {
    return { modo: 'nombre', nombre: args[nombreIdx + 1], outputDir };
  }

  // Modo: rango
  const rangoIdx = args.indexOf('--rango');
  if (rangoIdx !== -1 && args[rangoIdx + 1] && args[rangoIdx + 2]) {
    const inicio = parseInt(args[rangoIdx + 1]);
    const fin    = parseInt(args[rangoIdx + 2]);

    if (isNaN(inicio) || isNaN(fin) || inicio < 1 || fin < inicio) {
      Logger.error('Rango inválido. Ejemplo: --rango 1 50');
      process.exit(1);
    }

    return { modo: 'rango', inicio, fin, outputDir };
  }

  // Modo por defecto: todos
  return { modo: 'todos', outputDir };
}

// ====================== AYUDA ======================
function mostrarAyuda(): void {
  console.log(chalk.bold.cyan(`
╔══════════════════════════════════════════════════════════╗
║  📱 GENERADOR DE QR CODES — desde Google Sheets          ║
╚══════════════════════════════════════════════════════════╝
`));
  console.log(chalk.white('USO:'));
  console.log('  tsx generarQR.ts                         Todos (límite config)');
  console.log('  tsx generarQR.ts --rango 1 50            Personas del 1 al 50');
  console.log('  tsx generarQR.ts --nombre "Juan Perez"   Una persona específica');
  console.log('  tsx generarQR.ts --output ./carpeta      Directorio personalizado');
  console.log('');
  console.log(chalk.white('SALIDA:'));
  console.log('  qrs_generados/<nombre_persona>/qr.png');
  console.log('');
  console.log(chalk.white('CONTENIDO DEL QR:'));
  console.log('  Texto plano con todos los campos del Google Sheets');
  console.log('  (grupo, nombre, apellidos, documento, cargo, teléfono, email, fechas)');
}

// ====================== FUNCIÓN PRINCIPAL ======================
async function main(): Promise<void> {
  console.log(chalk.bold.cyan('\n📱 GENERADOR DE QR CODES\n'));

  const opciones = parsearArgs();
  const sheets   = new GoogleSheetsService();
  const qrGen    = new QRGenerator();

  // ── Obtener personas ──────────────────────────────────────────
  let personas;

  try {
    await sheets.inicializar();

    switch (opciones.modo) {
      case 'rango':
        personas = await sheets.obtenerPersonas(undefined, opciones.inicio, opciones.fin);
        Logger.info(`Rango seleccionado: ${opciones.inicio} → ${opciones.fin}`);
        break;

      case 'nombre': {
        const p = await sheets.obtenerPersonaPorNombre(opciones.nombre!);
        if (!p) {
          Logger.error(`No se encontró persona con nombre: "${opciones.nombre}"`);
          process.exit(1);
        }
        personas = [p];
        break;
      }

      default: // 'todos'
        personas = await sheets.obtenerPersonas(CONFIG.LIMITE_PERSONAS);
        Logger.info(`Modo: todas las personas (máx. ${CONFIG.LIMITE_PERSONAS})`);
        break;
    }
  } catch (error) {
    Logger.error(`Error conectando con Google Sheets: ${error}`);
    process.exit(1);
  }

  if (personas.length === 0) {
    Logger.warn('No se encontraron personas para procesar.');
    process.exit(0);
  }

  // ── Generar QRs ───────────────────────────────────────────────
  const resultado = await qrGen.generarQRMasivo(personas, opciones.outputDir);

  // ── Resumen final ─────────────────────────────────────────────
  Logger.separador();
  console.log(chalk.bold('\n📊 RESUMEN FINAL'));
  console.log(`✅ QRs generados:  ${resultado.exitosos}`);
  console.log(`❌ Errores:        ${resultado.errores}`);
  console.log(`📁 Guardados en:   ${opciones.outputDir}`);

  process.exit(resultado.errores > 0 ? 1 : 0);
}

// ── Manejo de errores no capturados ───────────────────────────────
process.on('unhandledRejection', (reason) => {
  Logger.error(`Error no manejado: ${reason}`);
  process.exit(1);
});

process.on('SIGINT', () => {
  console.log(chalk.yellow('\n⚠️  Proceso interrumpido'));
  process.exit(0);
});

main();