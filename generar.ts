#!/usr/bin/env node

import { crearComandos } from './src/cli/commands.js';
import { Logger } from './src/utils/fileUtils.js';
import chalk from 'chalk';

// ====================== BANNER DE BIENVENIDA ======================
function mostrarBanner(): void {
  console.log(chalk.bold.cyan(`
╔══════════════════════════════════════════════════════════════════╗
║                                                                  ║
║    🚀 GENERADOR MASIVO DE DOCUMENTOS v2.0                       ║
║    📄 Desde Google Sheets → Word/Excel → PDF                    ║
║    ⚡ Optimizado para procesamiento en lotes                    ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝
  `));
}

// ====================== FUNCIÓN PRINCIPAL ======================
async function main(): Promise<void> {
  try {
    // Mostrar banner solo si no hay argumentos (modo interactivo por defecto)
    if (process.argv.length <= 2) {
      mostrarBanner();
      console.log(chalk.yellow('💡 Usa --help para ver todas las opciones disponibles\n'));
      
      // Ejecutar modo interactivo por defecto
      const programa = crearComandos();
      await programa.parseAsync(['node', 'generar.ts', 'interactivo']);
      return;
    }

    // Procesar comandos CLI
    const programa = crearComandos();
    await programa.parseAsync(process.argv);

  } catch (error) {
    Logger.error(`Error fatal: ${error}`);
    process.exit(1);
  }
}

// ====================== MANEJO DE ERRORES NO CAPTURADOS ======================
process.on('unhandledRejection', (reason, promise) => {
  Logger.error(`Error no manejado en Promise: ${reason}`);
  process.exit(1);
});

process.on('uncaughtException', (error) => {
  Logger.error(`Excepción no capturada: ${error.message}`);
  process.exit(1);
});

// ====================== MANEJO DE CTRL+C ======================
process.on('SIGINT', () => {
  console.log(chalk.yellow('\n⚠️ Proceso interrumpido por el usuario'));
  process.exit(0);
});

// ====================== EJECUTAR ======================
main();