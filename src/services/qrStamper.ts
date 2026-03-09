import fs from 'fs/promises';
import path from 'path';
import { PDFDocument } from 'pdf-lib';
import QRCode from 'qrcode';
import { Persona } from '../types/index.js';
import { Logger } from '../utils/fileUtils.js';
import { CONFIG } from '../config/settings.js';

// ====================== TIPOS ======================

export type PosicionQR =
  | 'inferior-derecha'
  | 'inferior-izquierda'
  | 'superior-derecha'
  | 'superior-izquierda'
  | { x: number; y: number }; // coordenadas manuales en puntos PDF (1 pt = 1/72 pulgada)

export interface ConfigQREnPDF {
  /**
   * Nombre de la plantilla (debe coincidir con PlantillaConfig.nombre en settings.ts).
   * Si se pone "*" aplica a todas las plantillas.
   */
  plantilla: string;

  /** Posición del QR en la página */
  posicion: PosicionQR;

  /** Tamaño del QR en puntos PDF. 1 cm ≈ 28.35 pt */
  tamanoPt: number;

  /**
   * Margen respecto al borde cuando se usa posición predefinida (en puntos).
   * Por defecto: 20 pt (~7 mm)
   */
  margenPt?: number;

  /**
   * Número de página donde se estampa (1-based).
   * Por defecto: última página.
   */
  pagina?: number | 'primera' | 'ultima';

  /** Opacidad del QR: 0 = transparente, 1 = sólido. Por defecto: 1 */
  opacidad?: number;
}

// ====================== CONFIGURACIÓN POR PLANTILLA ======================
// 👈 EDITÁ AQUÍ para cambiar posición/tamaño por documento
//
// Conversión rápida cm → pt:  1 cm ≈ 28.35 pt
//   2 cm =  57 pt  |  3 cm =  85 pt  |  4 cm = 113 pt
//   5 cm = 142 pt  |  6 cm = 170 pt  |  7 cm = 198 pt
//
// Posiciones disponibles:
//   'superior-derecha' | 'superior-izquierda'
//   'inferior-derecha' | 'inferior-izquierda'
//   { x: 450, y: 600 }  ← coordenadas manuales en pt (origen = esquina inferior-izquierda)
//
// Páginas: 'primera' | 'ultima' | 1 | 2 | 3 ...

export const CONFIG_QR_EN_PDF: ConfigQREnPDF[] = [
  {
    // Comodín: aplica a TODOS los PDFs generados
    plantilla: '*',
    posicion: 'superior-derecha',
    tamanoPt: 170,   // 6×6 cm
    margenPt: 20,    // ~7 mm desde el borde
    pagina: 'primera',
  },
];

// ====================== SERVICIO ESTAMPADOR ======================

export class QRStamper {

  /**
   * Construye el texto plano que irá codificado en el QR.
   * Mismo formato que qrGenerator.ts para consistencia.
   */
  // Elimina tildes y caracteres especiales: a -> a, e -> e, n~/ -> n, etc.
  // Soluciona el problema de letras chinas en el QR cuando hay acentos.
  private normalizarTexto(texto: string): string {
    return texto
      .normalize('NFD')                // descompone: á -> a + tilde
      .replace(/[\u0300-\u036f]/g, '') // elimina diacriticos (tildes, dieresis)
      .replace(/[\u00d1]/g, 'N')       // N con tilde
      .replace(/[\u00f1]/g, 'n')       // n con tilde
      .replace(/[^\x00-\x7F]/g, '');  // elimina cualquier otro no-ASCII
  }

  private construirTextoQR(persona: Persona): string {
    const campos: [string, string | undefined][] = [
      ['GRUPO',        persona.grupo],
      ['NOMBRE',       persona.nombre],
      ['APELLIDO 1',   persona.apellido1],
      ['APELLIDO 2',   persona.apellido2],
      ['DOCUMENTO',    persona.documento],
      ['CARGO',        persona.cargo],
      ['TELEFONO',     persona.telefono],
      ['EMAIL',        persona.email],
      ['FECHA INICIO', persona.fecha_inicio],
      ['FECHA FIN',    persona.fecha_fin],
    ];

    return campos
      .filter(([, v]) => v && v.trim())
      .map(([k, v]) => `${k}: ${this.normalizarTexto(v!)}`)
      .join('\n');
  }

  /**
   * Genera el QR como buffer PNG en memoria (sin tocar el disco).
   *
   * ¿Por qué Buffer.from(texto, 'latin1')?
   * La librería `qrcode` codifica en modo Byte sin declarar la cabecera ECI.
   * Sin esa cabecera, los lectores QR asumen que los bytes son Latin-1
   * (ISO-8859-1). Cuando el texto viene como string UTF-8 de Node.js, los
   * caracteres acentuados (á, é, ñ, etc.) se representan con 2 bytes UTF-8,
   * que Latin-1 interpreta como símbolos raros o "letras chinas".
   *
   * La solución: convertir el string a Buffer con encoding 'latin1'.
   * Así los bytes del buffer ya son Latin-1 real, y al escanearlo cualquier
   * cámara o app lo decodifica correctamente mostrando á, é, í, ó, ú, ñ, ü.
   * Latin-1 cubre todos los caracteres del español.
   */
  private async generarQRBuffer(persona: Persona): Promise<Buffer> {
    const texto = this.construirTextoQR(persona);

    // Forzar encoding Latin-1 usando QRCodeSegment con modo 'byte'.
    // QRCode.toBuffer() solo acepta string o QRCodeSegment[], no Buffer directamente.
    // Al pasar los bytes como Uint8Array en un segmento modo 'byte', la librería
    // los escribe tal cual en el QR — sin reinterpretar como UTF-8.
    // Los lectores QR asumen Latin-1 en modo byte sin cabecera ECI,
    // por lo que á=0xE1, é=0xE9, ñ=0xF1 se muestran correctamente.
    // NFC + pasar string directo activa el modo ECI UTF-8 en la librería qrcode,
    // lo que garantiza que á/é/ñ/ü se lean correctamente en cualquier lector.
    return QRCode.toBuffer(texto.normalize('NFC'), {
      type: 'png',
      errorCorrectionLevel: 'M',
      width: 600,
      margin: 1,
      color: { dark: '#000000', light: '#ffffff' },
    });
  }

  /**
   * Calcula las coordenadas X, Y (esquina inferior-izquierda del QR)
   * en el sistema de coordenadas de pdf-lib (origen = esquina inferior-izquierda de la página).
   */
  private calcularCoordenadas(
    posicion: PosicionQR,
    tamanoPt: number,
    margenPt: number,
    pageWidth: number,
    pageHeight: number
  ): { x: number; y: number } {
    if (typeof posicion === 'object') {
      return posicion; // coordenadas manuales directas
    }

    switch (posicion) {
      case 'inferior-derecha':
        return { x: pageWidth - tamanoPt - margenPt, y: margenPt };
      case 'inferior-izquierda':
        return { x: margenPt, y: margenPt };
      case 'superior-derecha':
        return { x: pageWidth - tamanoPt - margenPt, y: pageHeight - tamanoPt - margenPt };
      case 'superior-izquierda':
        return { x: margenPt, y: pageHeight - tamanoPt - margenPt };
    }
  }

  /**
   * Obtiene la configuración QR para una plantilla dada.
   * Primero busca coincidencia exacta, luego usa el comodín "*".
   */
  private obtenerConfig(nombrePlantilla: string): ConfigQREnPDF {
    const exacta = CONFIG_QR_EN_PDF.find(c => c.plantilla === nombrePlantilla);
    if (exacta) return exacta;

    const comodin = CONFIG_QR_EN_PDF.find(c => c.plantilla === '*');
    if (comodin) return comodin;

    // Fallback si no hay nada configurado
    return {
      plantilla: '*',
      posicion: 'inferior-derecha',
      tamanoPt: 85,
      margenPt: 20,
      pagina: 'ultima',
    };
  }

  /**
   * Estampa el QR de una persona sobre un PDF existente.
   *
   * @param rutaPDF      Ruta al PDF generado por LibreOffice
   * @param persona      Datos de la persona (para generar el QR)
   * @param plantillaNombre  Nombre de la plantilla (para buscar config)
   * @param sobreescribir    Si true, reemplaza el PDF original. Default: true.
   * @returns            Ruta del PDF final con el QR incrustado
   */
  async estamparQREnPDF(
    rutaPDF: string,
    persona: Persona,
    plantillaNombre: string,
    sobreescribir = true
  ): Promise<string> {
    const config   = this.obtenerConfig(plantillaNombre);
    const margen   = config.margenPt ?? 20;
    const opacidad = config.opacidad ?? 1;

    // 1. Leer el PDF existente
    const pdfBytes     = await fs.readFile(rutaPDF);
    const pdfDoc       = await PDFDocument.load(pdfBytes);
    const paginas      = pdfDoc.getPages();

    if (paginas.length === 0) {
      throw new Error(`El PDF no tiene páginas: ${rutaPDF}`);
    }

    // 2. Seleccionar página destino
    let indicePagina: number;
    if (config.pagina === 'primera' || config.pagina === 1) {
      indicePagina = 0;
    } else if (!config.pagina || config.pagina === 'ultima') {
      indicePagina = paginas.length - 1;
    } else {
      indicePagina = Math.min((config.pagina as number) - 1, paginas.length - 1);
    }

    const pagina = paginas[indicePagina];
    const { width: pageWidth, height: pageHeight } = pagina.getSize();

    // 3. Generar QR como PNG en memoria
    const qrBuffer = await this.generarQRBuffer(persona);

    // 4. Embeber imagen PNG en el PDF
    const qrImage = await pdfDoc.embedPng(qrBuffer);

    // 5. Calcular posición
    const { x, y } = this.calcularCoordenadas(
      config.posicion,
      config.tamanoPt,
      margen,
      pageWidth,
      pageHeight
    );

    // 6. Dibujar el QR sobre la página
    pagina.drawImage(qrImage, {
      x,
      y,
      width: config.tamanoPt,
      height: config.tamanoPt,
      opacity: opacidad,
    });

    // 7. Serializar el PDF modificado
    const pdfFinal = await pdfDoc.save();

    // 8. Guardar (sobreescribe o nuevo archivo)
    const rutaSalida = sobreescribir
      ? rutaPDF
      : rutaPDF.replace('.pdf', '_qr.pdf');

    await fs.writeFile(rutaSalida, pdfFinal);
    return rutaSalida;
  }

  /**
   * Estampa QR en todos los PDFs de una carpeta de persona.
   * Busca archivos .pdf y los procesa según la configuración de cada plantilla.
   *
   * @param dirPersona      Carpeta de salida de una persona (ej: docs_generados/1_perez_juan/)
   * @param persona         Datos de la persona
   * @param plantillasActivas  Nombres de plantillas que se procesaron
   */
  async estamparQREnCarpeta(
    dirPersona: string,
    persona: Persona,
    plantillasActivas: string[]
  ): Promise<{ exitosos: number; errores: number }> {
    let exitosos = 0;
    let errores  = 0;

    // Filtrar solo las plantillas que tienen qr: true en settings.ts
    const plantillasConQR = CONFIG.PLANTILLAS
      .filter(p => p.qr === true && plantillasActivas.includes(p.nombre))
      .map(p => p.nombre);

    if (plantillasConQR.length === 0) {
      Logger.info('  Ninguna plantilla activa tiene qr: true, se omite el estampado');
      return { exitosos: 0, errores: 0 };
    }

    for (const nombrePlantilla of plantillasConQR) {
      // El PDF tiene el mismo nombre que la plantilla
      const rutaPDF = path.join(dirPersona, `${nombrePlantilla}.pdf`);

      try {
        await fs.access(rutaPDF); // verificar que existe
      } catch {
        // No hay PDF para esta plantilla (puede ser solo-originals), saltar
        continue;
      }

      try {
        await this.estamparQREnPDF(rutaPDF, persona, nombrePlantilla);
        exitosos++;
      } catch (err) {
        errores++;
        const msg = err instanceof Error ? err.message : String(err);
        Logger.error(`  QR stamp falló en ${nombrePlantilla}.pdf: ${msg}`);
      }
    }

    return { exitosos, errores };
  }
}