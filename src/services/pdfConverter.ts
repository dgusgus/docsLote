import { spawn } from 'child_process';
import path from 'path';
import { CONFIG, TIMEOUTS } from '../config/settings.js';
import { Logger, verificarArchivoExiste } from '../utils/fileUtils.js';

// ====================== SERVICIO CONVERSOR PDF ======================
export class PDFConverter {
  private static instancia: PDFConverter;
  private procesoActivo = false;

  static obtenerInstancia(): PDFConverter {
    if (!PDFConverter.instancia) {
      PDFConverter.instancia = new PDFConverter();
    }
    return PDFConverter.instancia;
  }

  async verificarLibreOffice(): Promise<boolean> {
    if (!verificarArchivoExiste(CONFIG.SOFFICE_PATH)) {
      Logger.error(`LibreOffice no encontrado en: ${CONFIG.SOFFICE_PATH}`);
      Logger.info("💡 Soluciones:");
      Logger.info("   1. Instala LibreOffice desde: https://www.libreoffice.org/");
      Logger.info("   2. Ajusta SOFFICE_PATH en src/config/settings.ts");
      Logger.info("   3. Usa modo 'solo-originals' para evitar conversión PDF");
      return false;
    }
    return true;
  }

  async convertirAPdf(
    inputPath: string, 
    outputDir: string, 
    intentos: number = 3
  ): Promise<string> {
    // Verificar que el archivo de entrada existe
    if (!verificarArchivoExiste(inputPath)) {
      throw new Error(`Archivo no encontrado: ${inputPath}`);
    }

    // Esperar si hay un proceso activo (LibreOffice no maneja bien concurrencia)
    while (this.procesoActivo) {
      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    this.procesoActivo = true;

    try {
      for (let intento = 1; intento <= intentos; intento++) {
        try {
          Logger.progress(`Convirtiendo a PDF (intento ${intento}/${intentos}): ${path.basename(inputPath)}`);
          
          const pdfGenerado = await this.ejecutarConversion(inputPath, outputDir);
          Logger.success(`PDF generado: ${path.basename(pdfGenerado)}`);
          
          return pdfGenerado;
        } catch (error) {
          if (intento === intentos) {
            throw error;
          }
          Logger.warn(`Intento ${intento} falló, reintentando...`);
          await new Promise(resolve => setTimeout(resolve, 2000)); // Esperar 2s entre intentos
        }
      }
      
      throw new Error("Todos los intentos de conversión fallaron");
    } finally {
      this.procesoActivo = false;
    }
  }

  private async ejecutarConversion(inputPath: string, outputDir: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error(`⏰ Timeout convirtiendo: ${path.basename(inputPath)}`));
      }, TIMEOUTS.PDF_CONVERSION);

      // Generar nombre del PDF resultante
      const inputName = path.basename(inputPath, path.extname(inputPath));
      const pdfPath = path.join(outputDir, `${inputName}.pdf`);

      const args = [
        "--headless",
        "--convert-to", "pdf",
        "--outdir", outputDir,
        inputPath
      ];

      Logger.info(`Ejecutando: ${CONFIG.SOFFICE_PATH} ${args.join(' ')}`);

      const proceso = spawn(CONFIG.SOFFICE_PATH, args, {
        stdio: ["ignore", "pipe", "pipe"], // Capturar stdout y stderr
        windowsHide: true // Ocultar ventana en Windows
      });

      let stdout = '';
      let stderr = '';

      if (proceso.stdout) {
        proceso.stdout.on('data', (data) => {
          stdout += data.toString();
        });
      }

      if (proceso.stderr) {
        proceso.stderr.on('data', (data) => {
          stderr += data.toString();
        });
      }

      proceso.on('exit', (code) => {
        clearTimeout(timeout);
        
        if (code === 0) {
          // Verificar que el PDF se creó realmente
          if (verificarArchivoExiste(pdfPath)) {
            resolve(pdfPath);
          } else {
            reject(new Error(`PDF no se generó en: ${pdfPath}`));
          }
        } else {
          const mensaje = stderr || stdout || `Código de salida: ${code}`;
          reject(new Error(`❌ LibreOffice falló: ${mensaje}`));
        }
      });

      proceso.on('error', (err) => {
        clearTimeout(timeout);
        reject(new Error(`❌ Error ejecutando LibreOffice: ${err.message}`));
      });
    });
  }

  async convertirMultiplesArchivos(
    archivos: string[], 
    outputDir: string,
    onProgress?: (completados: number, total: number) => void
  ): Promise<{ exitosos: string[], fallidos: { archivo: string, error: string }[] }> {
    const exitosos: string[] = [];
    const fallidos: { archivo: string, error: string }[] = [];

    Logger.info(`Iniciando conversión de ${archivos.length} archivos a PDF`);

    for (let i = 0; i < archivos.length; i++) {
      const archivo = archivos[i];
      try {
        const pdfGenerado = await this.convertirAPdf(archivo, outputDir);
        exitosos.push(pdfGenerado);
        
        if (onProgress) {
          onProgress(i + 1, archivos.length);
        }
      } catch (error) {
        const mensajeError = error instanceof Error ? error.message : String(error);
        fallidos.push({ archivo, error: mensajeError });
        Logger.error(`Error convirtiendo ${path.basename(archivo)}: ${mensajeError}`);
      }

      // Pausa pequeña entre conversiones para estabilidad
      if (i < archivos.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }

    Logger.success(`Conversión completada: ${exitosos.length} éxitos, ${fallidos.length} errores`);
    
    return { exitosos, fallidos };
  }

  // ====================== MÉTODOS DE UTILIDAD ======================
  obtenerExtensionesValidas(): string[] {
    return ['.docx', '.xlsx', '.doc', '.xls', '.odt', '.ods', '.rtf', '.txt'];
  }

  esArchivoConvertible(archivo: string): boolean {
    const extension = path.extname(archivo).toLowerCase();
    return this.obtenerExtensionesValidas().includes(extension);
  }

  async obtenerVersionLibreOffice(): Promise<string | null> {
    try {
      return new Promise((resolve) => {
        const proceso = spawn(CONFIG.SOFFICE_PATH, ['--version'], {
          stdio: ['ignore', 'pipe', 'ignore']
        });

        let version = '';
        if (proceso.stdout) {
          proceso.stdout.on('data', (data) => {
            version += data.toString();
          });
        }

        proceso.on('exit', () => {
          resolve(version.trim() || null);
        });

        proceso.on('error', () => {
          resolve(null);
        });

        // Timeout de 5 segundos
        setTimeout(() => {
          proceso.kill();
          resolve(null);
        }, 5000);
      });
    } catch {
      return null;
    }
  }
}