import fs from 'fs/promises';
import path from 'path';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import ExcelJS from 'exceljs';
import { Persona, PlantillaConfig } from '../types/index.js';
import { CONFIG, MAPEOS_EXCEL } from '../config/settings.js';
import { Logger, verificarArchivoExiste } from '../utils/fileUtils.js';

// ====================== SERVICIO GENERADOR DE DOCUMENTOS ======================
export class DocumentGenerator {
  private plantillaCache = new Map<string, Buffer>();

  async precargarPlantillas(plantillas?: string[]): Promise<void> {
    Logger.progress("Precargando plantillas en memoria...");
    
    const plantillasACargar = plantillas 
      ? CONFIG.PLANTILLAS.filter(p => plantillas.includes(p.nombre))
      : CONFIG.PLANTILLAS;

    let cargadas = 0;
    let errores = 0;

    for (const plantilla of plantillasACargar) {
      try {
        if (!verificarArchivoExiste(plantilla.archivo)) {
          Logger.warn(`Plantilla no encontrada: ${plantilla.archivo}`);
          errores++;
          continue;
        }

        const content = await fs.readFile(plantilla.archivo);
        this.plantillaCache.set(plantilla.nombre, content);
        Logger.info(`  ✅ ${plantilla.nombre} (${plantilla.descripcion || 'Sin descripción'})`);
        cargadas++;
      } catch (err) {
        Logger.error(`  ❌ Error cargando ${plantilla.nombre}: ${err}`);
        errores++;
      }
    }

    if (cargadas === 0) {
      throw new Error("No se pudo cargar ninguna plantilla");
    }

    Logger.success(`${cargadas} plantillas cargadas en memoria${errores > 0 ? ` (${errores} errores)` : ''}`);
  }

  async generarWord(
    datos: Persona,
    plantillaNombre: string,
    salidaPath: string
  ): Promise<void> {
    const plantillaContent = this.plantillaCache.get(plantillaNombre);
    if (!plantillaContent) {
      throw new Error(`❌ Plantilla ${plantillaNombre} no encontrada en cache`);
    }

    try {
      const zip = new PizZip(plantillaContent);
      const doc = new Docxtemplater(zip, {
        delimiters: { start: "[[", end: "]]" },
        errorLogging: false,
      });

      // 👈 DATOS ENRIQUECIDOS AUTOMÁTICOS
      const datosCompletos = this.enriquecerDatos(datos);

      doc.render(datosCompletos);
      const buffer = doc.getZip().generate({ type: "nodebuffer" });
      await fs.writeFile(salidaPath, buffer);

    } catch (err: any) {
      if (err.name === 'TemplateError') {
        throw new Error(`❌ Error en plantilla ${plantillaNombre}: ${err.message}`);
      }
      throw new Error(`❌ Error generando Word ${plantillaNombre}: ${err}`);
    }
  }

  async generarExcel(
    datos: Persona,
    plantillaPath: string,
    salidaPath: string,
    plantillaNombre: string
  ): Promise<void> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(plantillaPath);
      const worksheet = workbook.worksheets[0];

      if (!worksheet) {
        throw new Error("La plantilla Excel no tiene hojas de trabajo");
      }

      // 👈 MAPEO DINÁMICO BASADO EN CONFIGURACIÓN
      const mapeo = MAPEOS_EXCEL[plantillaNombre];
      if (mapeo) {
        Object.entries(mapeo).forEach(([campo, celda]) => {
          const valor = datos[campo as keyof Persona];
          if (valor && worksheet.getCell(celda)) {
            worksheet.getCell(celda).value = valor;
          }
        });
      } else {
        // 👈 MAPEO POR DEFECTO PARA PLANTILLAS SIN CONFIGURACIÓN
        this.aplicarMapeoDefaultExcel(worksheet, datos);
      }

      await workbook.xlsx.writeFile(salidaPath);

    } catch (err) {
      throw new Error(`❌ Error generando Excel ${plantillaNombre}: ${err}`);
    }
  }

  private aplicarMapeoDefaultExcel(worksheet: ExcelJS.Worksheet, datos: Persona): void {
    // Mapeo genérico para plantillas sin configuración específica
    const celdas = ['C8', 'M8', 'U8', 'H10', 'H16', 'U16'];
    const valores = [
      datos.nombre,
      datos.apellido1,
      datos.apellido2,
      datos.documento,
      datos.telefono,
      datos.email
    ];

    celdas.forEach((celda, index) => {
      if (valores[index] && worksheet.getCell(celda)) {
        try {
          worksheet.getCell(celda).value = valores[index];
        } catch {
          // Ignorar si la celda no es válida
        }
      }
    });
  }

  private enriquecerDatos(datos: Persona): Persona & Record<string, any> {
    const ahora = new Date();
    
    return {
      ...datos,
      // 👈 DATOS AUTOMÁTICOS DE FECHA
      fecha_actual: ahora.toLocaleDateString('es-ES'),
      fecha_actual_larga: ahora.toLocaleDateString('es-ES', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      }),
      año_actual: ahora.getFullYear(),
      mes_actual: ahora.toLocaleDateString('es-ES', { month: 'long' }),
      
      // 👈 COMBINACIONES DE NOMBRES
      nombre_completo: `${datos.nombre} ${datos.apellido1} ${datos.apellido2}`.trim(),
      apellidos_completos: `${datos.apellido1} ${datos.apellido2}`.trim(),
      iniciales: this.obtenerIniciales(datos.nombre, datos.apellido1, datos.apellido2),
      
      // 👈 DATOS PROCESADOS
      email_dominio: datos.email ? datos.email.split('@')[1] : '',
      telefono_formateado: this.formatearTelefono(datos.telefono || ''),
      
      // 👈 DATOS DE CURSO
      curso_mayuscula: datos.curso?.toUpperCase() || '',
      curso_minuscula: datos.curso?.toLowerCase() || '',
      
      // 👈 UTILIDADES
      numero_aleatorio: Math.floor(Math.random() * 10000),
      codigo_unico: this.generarCodigoUnico(datos.nombre, datos.apellido1)
    };
  }

  private obtenerIniciales(nombre: string, apellido1: string, apellido2: string): string {
    const iniciales = [nombre, apellido1, apellido2]
      .filter(Boolean)
      .map(str => str.charAt(0).toUpperCase())
      .join('.');
    
    return iniciales;
  }

  private formatearTelefono(telefono: string): string {
    const numeros = telefono.replace(/\D/g, '');
    if (numeros.length >= 8) {
      return `${numeros.slice(0, 3)}-${numeros.slice(3, 6)}-${numeros.slice(6)}`;
    }
    return telefono;
  }

  private generarCodigoUnico(nombre: string, apellido: string): string {
    const base = `${nombre.slice(0, 2)}${apellido.slice(0, 2)}`.toUpperCase();
    const numero = Date.now().toString().slice(-4);
    return `${base}${numero}`;
  }

  // ====================== MÉTODOS DE UTILIDAD ======================
  obtenerPlantillasDisponibles(): PlantillaConfig[] {
    return CONFIG.PLANTILLAS;
  }

  obtenerPlantillasPorTipo(tipo: 'word' | 'excel'): PlantillaConfig[] {
    return CONFIG.PLANTILLAS.filter(p => p.tipo === tipo);
  }

  validarPlantillasExisten(): { existentes: string[], faltantes: string[] } {
    const existentes: string[] = [];
    const faltantes: string[] = [];

    CONFIG.PLANTILLAS.forEach(plantilla => {
      if (verificarArchivoExiste(plantilla.archivo)) {
        existentes.push(plantilla.nombre);
      } else {
        faltantes.push(plantilla.nombre);
      }
    });

    return { existentes, faltantes };
  }
}