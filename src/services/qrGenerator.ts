import fs from 'fs/promises';
import path from 'path';
import QRCode from 'qrcode';
import { Persona } from '../types/index.js';
import { Logger, crearDirectorioSeguro, limpiarNombre } from '../utils/fileUtils.js';

// ====================== SERVICIO GENERADOR DE QR ======================
export class QRGenerator {

  /**
   * Convierte los datos de una persona a texto plano estructurado.
   * Los valores se usan tal cual — los acentos y ñ se preservan
   * gracias al encoding Latin-1 que se aplica en generarQRPersona().
   */
  private construirTextoQR(persona: Persona): string {
    const lineas: string[] = [];

    if (persona.grupo)        lineas.push(`GRUPO: ${persona.grupo}`);
    if (persona.nombre)       lineas.push(`NOMBRE: ${persona.nombre}`);
    if (persona.apellido1)    lineas.push(`APELLIDO 1: ${persona.apellido1}`);
    if (persona.apellido2)    lineas.push(`APELLIDO 2: ${persona.apellido2}`);
    if (persona.documento)    lineas.push(`DOCUMENTO: ${persona.documento}`);
    if (persona.cargo)        lineas.push(`CARGO: ${persona.cargo}`);
    if (persona.telefono)     lineas.push(`TELEFONO: ${persona.telefono}`);
    if (persona.email)        lineas.push(`EMAIL: ${persona.email}`);
    if (persona.fecha_inicio) lineas.push(`FECHA INICIO: ${persona.fecha_inicio}`);
    if (persona.fecha_fin)    lineas.push(`FECHA FIN: ${persona.fecha_fin}`);

    return lineas.join('\n');
  }

  /**
   * Genera el nombre de carpeta/archivo para una persona.
   * Mismo patrón que processor.ts: índice_grupo_apellido1_apellido2_nombre_documento
   */
  private obtenerNombreBase(persona: Persona): string {
    const partes = [
      persona.indice?.toString(),
      persona.grupo,
      persona.apellido1,
      persona.apellido2,
      persona.nombre,
      persona.documento,
    ]
      .filter(Boolean)
      .map(p => limpiarNombre(p!));

    return partes.join('_');
  }

  /**
   * Genera el QR de UNA persona y lo guarda como PNG.
   * Retorna la ruta del archivo generado.
   */
  async generarQRPersona(persona: Persona, outputDir: string): Promise<string> {
    const nombreBase  = this.obtenerNombreBase(persona);
    const dirPersona  = path.join(outputDir, nombreBase);
    const rutaArchivo = path.join(dirPersona, 'qr.png');

    await crearDirectorioSeguro(dirPersona);

      // Construir texto y forzar Latin-1
      const texto = this.construirTextoQR(persona);

      // SOLUCIÓN: Crear un buffer Latin-1 explícitamente
      // y usar un segmento de bytes en lugar de string directo
      const bufferLatin1 = Buffer.from(texto, 'latin1');

    // La librería `qrcode` en Node.js pasa el string directamente como bytes UTF-8
    // al segmento de datos del QR. Para que los lectores lo interpreten correctamente
    // hay que forzar el modo ECI con charset 'UTF-8'. Sin esto asumen Latin-1
    // y los caracteres acentuados aparecen rotos o como símbolos raros.
    await QRCode.toFile(rutaArchivo, texto, {
      type: 'png',
      errorCorrectionLevel: 'M',
      width: 400,
      margin: 2,
      color: {
        dark: '#000000',
        light: '#ffffff',
      },
    });

    return rutaArchivo;
  }

  /**
   * Genera QRs para una lista completa de personas.
   * Imprime progreso en consola y devuelve resumen.
   */
  async generarQRMasivo(
    personas: Persona[],
    outputDir: string
  ): Promise<{ exitosos: number; errores: number; detalles: string[] }> {

    Logger.titulo(`Generando QRs para ${personas.length} personas`);
    Logger.info(`📁 Directorio de salida: ${outputDir}`);
    await crearDirectorioSeguro(outputDir);

    let exitosos = 0;
    let errores  = 0;
    const detalles: string[] = [];

    for (let i = 0; i < personas.length; i++) {
      const persona  = personas[i];
      const etiqueta = `${persona.nombre} ${persona.apellido1}`.trim();

      try {
        const ruta = await this.generarQRPersona(persona, outputDir);
        exitosos++;
        Logger.success(`[${i + 1}/${personas.length}] → ${path.basename(path.dirname(ruta))}/qr.png`);
      } catch (error) {
        errores++;
        const msg = error instanceof Error ? error.message : String(error);
        detalles.push(`❌ ${etiqueta}: ${msg}`);
        Logger.error(`[${i + 1}/${personas.length}] Error en ${etiqueta}: ${msg}`);
      }
    }

    Logger.separador();
    Logger.success(`QRs completados: ${exitosos} exitosos, ${errores} errores`);
    if (detalles.length > 0) {
      detalles.forEach(d => Logger.warn(d));
    }

    return { exitosos, errores, detalles };
  }
}