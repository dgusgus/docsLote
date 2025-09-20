import fs from 'fs/promises';
import { existsSync } from 'fs';
import path from 'path';
import chalk from 'chalk';

// ====================== UTILIDADES DE ARCHIVOS ======================
export function limpiarNombre(nombre: string): string {
  return nombre
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Quita acentos
    .replace(/[^\w\s-]/g, "") // Solo letras, números, espacios y guiones
    .replace(/\s+/g, "_") // Espacios → guiones bajos
    .toLowerCase()
    .substring(0, 50);
}

export async function crearDirectorioSeguro(rutaDir: string): Promise<void> {
  try {
    await fs.mkdir(rutaDir, { recursive: true });
  } catch (error) {
    throw new Error(`No se pudo crear directorio ${rutaDir}: ${error}`);
  }
}

export function verificarArchivoExiste(rutaArchivo: string): boolean {
  return existsSync(rutaArchivo);
}

export async function verificarPermisos(rutaArchivo: string): Promise<boolean> {
  try {
    await fs.access(rutaArchivo, fs.constants.R_OK | fs.constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

export async function limpiarArchivosTemporales(directorio: string, extensiones: string[] = ['.tmp', '.temp']): Promise<void> {
  try {
    const archivos = await fs.readdir(directorio);
    const archivosTemporales = archivos.filter(archivo => 
      extensiones.some(ext => archivo.endsWith(ext))
    );

    await Promise.all(
      archivosTemporales.map(archivo => 
        fs.unlink(path.join(directorio, archivo)).catch(() => {})
      )
    );
  } catch {
    // Ignorar errores de limpieza
  }
}

export function obtenerTamañoArchivo(rutaArchivo: string): Promise<number> {
  return fs.stat(rutaArchivo).then(stats => stats.size);
}

export function formatearTamaño(bytes: number): string {
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  if (bytes === 0) return '0 Bytes';
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
}

// ====================== LOGGER SIMPLE ======================
export class Logger {
  private static escribirLog = true;

  static info(mensaje: string): void {
    console.log(chalk.blue('ℹ️ '), mensaje);
  }

  static success(mensaje: string): void {
    console.log(chalk.green('✅'), mensaje);
  }

  static warn(mensaje: string): void {
    console.log(chalk.yellow('⚠️ '), mensaje);
  }

  static error(mensaje: string): void {
    console.log(chalk.red('❌'), mensaje);
  }

  static progress(mensaje: string): void {
    console.log(chalk.cyan('🔄'), mensaje);
  }

  static titulo(mensaje: string): void {
    console.log(chalk.bold.magenta(`\n🚀 ${mensaje}`));
  }

  static separador(): void {
    console.log(chalk.gray('─'.repeat(50)));
  }
}