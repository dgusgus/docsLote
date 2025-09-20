import path from 'path';
import cliProgress from 'cli-progress';
import ora from 'ora';
import { GoogleSheetsService } from '../services/googleSheets.js';
import { DocumentGenerator } from '../services/documentGenerator.js';
import { PDFConverter } from '../services/pdfConverter.js';
import { CONFIG, PATHS } from '../config/settings.js';
import { Logger, crearDirectorioSeguro, limpiarNombre, limpiarArchivosTemporales } from '../utils/fileUtils.js';
import { OpcionesEjecucion, Persona, ResultadoProceso, ResultadoPersona } from '../types/index.js';

// ====================== PROCESADOR PRINCIPAL ======================
export class ProcesadorPrincipal {
  private sheetsService = new GoogleSheetsService();
  private docGenerator = new DocumentGenerator();
  private pdfConverter = PDFConverter.obtenerInstancia();

  async ejecutar(opciones: OpcionesEjecucion): Promise<ResultadoProceso> {
    const inicioTiempo = Date.now();
    
    Logger.titulo(`🚀 INICIANDO PROCESAMIENTO - MODO: ${opciones.modo.toUpperCase()}`);
    
    // Paso 1: Inicialización
    await this.inicializar(opciones);

    // Paso 2: Obtener personas
    const personas = await this.obtenerPersonas(opciones);
    
    if (personas.length === 0) {
      throw new Error('No se encontraron personas para procesar');
    }

    // Paso 3: Configurar directorio de salida
    const outputDir = opciones.outputDir || CONFIG.OUTPUT_DIR;
    await crearDirectorioSeguro(outputDir);

    // Paso 4: Procesar documentos
    const resultados = await this.procesarPersonas(personas, opciones, outputDir);

    // Paso 5: Generar reporte final
    const tiempoTotal = (Date.now() - inicioTiempo) / 1000;
    const resultado: ResultadoProceso = {
      exitos: resultados.reduce((sum, r) => sum + r.exitos, 0),
      errores: resultados.reduce((sum, r) => sum + r.errores, 0),
      tiempoTotal,
      personasProcesadas: personas.length,
      detalles: resultados
    };

    this.mostrarReporteFinal(resultado, outputDir);
    return resultado;
  }

  private async inicializar(opciones: OpcionesEjecucion): Promise<void> {
    const spinner = ora('Inicializando servicios...').start();

    try {
      // Inicializar Google Sheets
      await this.sheetsService.inicializar();
      spinner.text = 'Conectado a Google Sheets ✓';

      // Verificar LibreOffice si es necesario
      if (opciones.tipoSalida !== 'solo-originals') {
        const libreOfficeOk = await this.pdfConverter.verificarLibreOffice();
        if (!libreOfficeOk) {
          spinner.fail('LibreOffice no encontrado');
          throw new Error('LibreOffice es requerido para generar PDFs');
        }
        spinner.text = 'LibreOffice verificado ✓';
      }

      // Precargar plantillas
      await this.docGenerator.precargarPlantillas(opciones.plantillasEspecificas);
      spinner.succeed('Servicios inicializados correctamente');

    } catch (error) {
      spinner.fail('Error en inicialización');
      throw error;
    }
  }

  private async obtenerPersonas(opciones: OpcionesEjecucion): Promise<Persona[]> {
    Logger.progress('Obteniendo datos de Google Sheets...');

    switch (opciones.modo) {
      case 'todos':
        return await this.sheetsService.obtenerPersonas(CONFIG.LIMITE_PERSONAS);

      case 'rango':
        if (!opciones.rangoInicio || !opciones.rangoFin) {
          throw new Error('Rango no especificado correctamente');
        }
        return await this.sheetsService.obtenerPersonas(
          undefined, 
          opciones.rangoInicio, 
          opciones.rangoFin
        );

      case 'especifico':
        if (!opciones.nombreEspecifico) {
          throw new Error('Nombre no especificado');
        }
        const persona = await this.sheetsService.obtenerPersonaPorNombre(opciones.nombreEspecifico);
        return persona ? [persona] : [];

      default:
        throw new Error(`Modo no válido: ${opciones.modo}`);
    }
  }

  private async procesarPersonas(
    personas: Persona[], 
    opciones: OpcionesEjecucion,
    outputDir: string
  ): Promise<ResultadoPersona[]> {
    
    Logger.info(`📄 Procesando ${personas.length} personas con ${CONFIG.PLANTILLAS.length} plantillas cada una`);
    Logger.info(`📁 Directorio de salida: ${outputDir}`);
    Logger.separador();

    // Configurar barra de progreso
    const progressBar = new cliProgress.SingleBar({
      format: 'Progreso |{bar}| {percentage}% | {value}/{total} | {duration_formatted} | ETA: {eta_formatted}',
      barCompleteChar: '\u2588',
      barIncompleteChar: '\u2591',
      hideCursor: true
    });

    progressBar.start(personas.length, 0);

    const resultados: ResultadoPersona[] = [];

    // Procesar cada persona secuencialmente
    for (let i = 0; i < personas.length; i++) {
      const persona = personas[i];
      const resultado = await this.procesarPersonaIndividual(persona, i + 1, personas.length, opciones, outputDir);
      resultados.push(resultado);
      progressBar.update(i + 1);
      
      // Pequeña pausa entre personas para estabilidad
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    progressBar.stop();
    return resultados;
  }

  private async procesarPersonaIndividual(
    persona: Persona,
    indice: number,
    total: number,
    opciones: OpcionesEjecucion,
    outputDir: string
  ): Promise<ResultadoPersona> {
    
    const nombreBase = limpiarNombre(persona.nombre);
    const dirUsuario = path.join(outputDir, nombreBase);
    
    const resultado: ResultadoPersona = {
      nombre: `${persona.nombre} ${persona.apellido1}`.trim(),
      exitos: 0,
      errores: 0,
      documentosGenerados: [],
      erroresDetallados: []
    };

    try {
      await crearDirectorioSeguro(dirUsuario);

      // Obtener plantillas a procesar
      const plantillasAProcesar = opciones.plantillasEspecificas
        ? CONFIG.PLANTILLAS.filter(p => opciones.plantillasEspecificas!.includes(p.nombre))
        : CONFIG.PLANTILLAS;

      // Procesar cada plantilla
      for (const plantilla of plantillasAProcesar) {
        try {
          await this.procesarPlantillaParaPersona(persona, plantilla, dirUsuario, opciones.tipoSalida);
          resultado.exitos++;
          resultado.documentosGenerados.push(plantilla.nombre);
        } catch (error) {
          resultado.errores++;
          const mensajeError = error instanceof Error ? error.message : String(error);
          resultado.erroresDetallados.push(`${plantilla.nombre}: ${mensajeError}`);
          Logger.error(`  ❌ ${plantilla.nombre}: ${mensajeError}`);
        }
      }

      // Limpiar archivos temporales si es necesario
      if (opciones.tipoSalida === 'solo-pdf') {
        await limpiarArchivosTemporales(dirUsuario, ['.docx', '.xlsx']);
      }

    } catch (error) {
      const mensajeError = error instanceof Error ? error.message : String(error);
      resultado.errores += CONFIG.PLANTILLAS.length;
      resultado.erroresDetallados.push(`Error general: ${mensajeError}`);
    }

    return resultado;
  }

  private async procesarPlantillaParaPersona(
    persona: Persona,
    plantilla: any,
    dirUsuario: string,
    tipoSalida: string
  ): Promise<void> {
    
    const extension = plantilla.tipo === "word" ? "docx" : "xlsx";
    const archivoOriginal = path.join(dirUsuario, `${plantilla.nombre}.${extension}`);

    // Generar documento original
    if (["solo-originals", "ambos"].includes(tipoSalida)) {
      if (plantilla.tipo === "word") {
        await this.docGenerator.generarWord(persona, plantilla.nombre, archivoOriginal);
      } else {
        await this.docGenerator.generarExcel(persona, plantilla.archivo, archivoOriginal, plantilla.nombre);
      }
    }

    // Convertir a PDF
    if (["solo-pdf", "ambos"].includes(tipoSalida)) {
      // Si solo queremos PDF, generar temporal
      if (tipoSalida === "solo-pdf") {
        if (plantilla.tipo === "word") {
          await this.docGenerator.generarWord(persona, plantilla.nombre, archivoOriginal);
        } else {
          await this.docGenerator.generarExcel(persona, plantilla.archivo, archivoOriginal, plantilla.nombre);
        }
      }

      await this.pdfConverter.convertirAPdf(archivoOriginal, dirUsuario);
    }
  }

  private mostrarReporteFinal(resultado: ResultadoProceso, outputDir: string): void {
    Logger.separador();
    Logger.titulo('🎉 PROCESAMIENTO COMPLETADO');
    
    console.log(`✅ Documentos generados exitosamente: ${resultado.exitos}`);
    console.log(`❌ Errores: ${resultado.errores}`);
    console.log(`👥 Personas procesadas: ${resultado.personasProcesadas}`);
    console.log(`⏱️  Tiempo total: ${Math.round(resultado.tiempoTotal)}s`);
    console.log(`📁 Resultados guardados en: ${outputDir}`);
    
    if (resultado.exitos > 0) {
      const velocidad = (resultado.exitos / resultado.tiempoTotal).toFixed(1);
      console.log(`📊 Velocidad promedio: ${velocidad} documentos/segundo`);
    }

    // Mostrar errores detallados si los hay
    if (resultado.errores > 0) {
      Logger.separador();
      Logger.warn('⚠️ RESUMEN DE ERRORES:');
      resultado.detalles.forEach(detalle => {
        if (detalle.errores > 0) {
          console.log(`❌ ${detalle.nombre}: ${detalle.errores} errores`);
          detalle.erroresDetallados.forEach(error => {
            console.log(`   • ${error}`);
          });
        }
      });
    }

    // Mostrar éxitos por persona
    if (resultado.exitos > 0) {
      Logger.separador();
      Logger.success('✅ DOCUMENTOS GENERADOS POR PERSONA:');
      resultado.detalles.forEach(detalle => {
        if (detalle.exitos > 0) {
          console.log(`✅ ${detalle.nombre}: ${detalle.exitos} documentos`);
          console.log(`   📄 ${detalle.documentosGenerados.join(', ')}`);
        }
      });
    }
  }
}