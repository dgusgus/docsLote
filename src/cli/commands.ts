import { Command } from 'commander';
import inquirer from 'inquirer';
import chalk from 'chalk';
import { GoogleSheetsService } from '../services/googleSheets.js';
import { DocumentGenerator } from '../services/documentGenerator.js';
import { PDFConverter } from '../services/pdfConverter.js';
import { CONFIG, CONFIGURACIONES_RAPIDAS } from '../config/settings.js';
import { Logger } from '../utils/fileUtils.js';
import { OpcionesEjecucion, TipoSalida, ModoEjecucion } from '../types/index.js';
import { ProcesadorPrincipal } from './processor.js';

// ====================== COMANDOS CLI ======================
export function crearComandos(): Command {
  const programa = new Command();

  programa
    .name('generador-docs')
    .description('Generador masivo de documentos desde Google Sheets')
    .version('2.0.0');

  // ====================== COMANDO INTERACTIVO ======================
  programa
    .command('interactivo')
    .alias('i')
    .description('Modo interactivo - configuración paso a paso')
    .action(async () => {
      try {
        await modoInteractivo();
      } catch (error) {
        Logger.error(`Error en modo interactivo: ${error}`);
        process.exit(1);
      }
    });

  // ====================== COMANDO RÁPIDO - TODOS ======================
  programa
    .command('todos')
    .description('Procesar todas las personas con todas las plantillas')
    .option('-t, --tipo <tipo>', 'Tipo de salida (solo-pdf, solo-originals, ambos)', 'ambos')
    .option('-l, --limite <numero>', 'Límite de personas a procesar', '500')
    .option('-o, --output <directorio>', 'Directorio de salida personalizado')
    .action(async (opciones) => {
      const config: OpcionesEjecucion = {
        modo: 'todos',
        tipoSalida: opciones.tipo as TipoSalida,
        outputDir: opciones.output
      };

      if (opciones.limite) {
        CONFIG.LIMITE_PERSONAS = parseInt(opciones.limite);
      }

      await ejecutarProcesamiento(config);
    });

  // ====================== COMANDO RANGO ======================
  programa
    .command('rango')
    .description('Procesar un rango específico de personas')
    .requiredOption('-i, --inicio <numero>', 'Número de persona inicial (1-based)')
    .requiredOption('-f, --fin <numero>', 'Número de persona final (1-based)')
    .option('-t, --tipo <tipo>', 'Tipo de salida (solo-pdf, solo-originals, ambos)', 'ambos')
    .option('-p, --plantillas <nombres>', 'Plantillas específicas (separadas por coma)')
    .option('-o, --output <directorio>', 'Directorio de salida personalizado')
    .action(async (opciones) => {
      const inicio = parseInt(opciones.inicio);
      const fin = parseInt(opciones.fin);

      if (inicio < 1 || fin < inicio) {
        Logger.error('El rango debe ser válido (inicio >= 1, fin >= inicio)');
        process.exit(1);
      }

      const config: OpcionesEjecucion = {
        modo: 'rango',
        tipoSalida: opciones.tipo as TipoSalida,
        rangoInicio: inicio,
        rangoFin: fin,
        plantillasEspecificas: opciones.plantillas?.split(','),
        outputDir: opciones.output
      };

      await ejecutarProcesamiento(config);
    });

  // ====================== COMANDO ESPECÍFICO ======================
  programa
    .command('persona')
    .description('Procesar una persona específica')
    .requiredOption('-n, --nombre <nombre>', 'Nombre o parte del nombre de la persona')
    .option('-t, --tipo <tipo>', 'Tipo de salida (solo-pdf, solo-originals, ambos)', 'ambos')
    .option('-p, --plantillas <nombres>', 'Plantillas específicas (separadas por coma)')
    .option('-o, --output <directorio>', 'Directorio de salida personalizado')
    .action(async (opciones) => {
      const config: OpcionesEjecucion = {
        modo: 'especifico',
        tipoSalida: opciones.tipo as TipoSalida,
        nombreEspecifico: opciones.nombre,
        plantillasEspecificas: opciones.plantillas?.split(','),
        outputDir: opciones.output
      };

      await ejecutarProcesamiento(config);
    });

  // ====================== COMANDOS DE INFORMACIÓN ======================
  programa
    .command('listar')
    .description('Listar personas disponibles en Google Sheets')
    .option('-l, --limite <numero>', 'Límite de personas a mostrar', '20')
    .action(async (opciones) => {
      try {
        const sheetsService = new GoogleSheetsService();
        await sheetsService.listarPersonas(parseInt(opciones.limite));
      } catch (error) {
        Logger.error(`Error listando personas: ${error}`);
      }
    });

  programa
    .command('plantillas')
    .description('Listar plantillas disponibles')
    .action(() => {
      Logger.titulo('Plantillas configuradas:');
      CONFIG.PLANTILLAS.forEach((plantilla, index) => {
        console.log(`${index + 1}. ${chalk.cyan(plantilla.nombre)} (${plantilla.tipo.toUpperCase()})`);
        console.log(`   📁 ${plantilla.archivo}`);
        if (plantilla.descripcion) {
          console.log(`   📝 ${plantilla.descripcion}`);
        }
        console.log();
      });
    });

  programa
    .command('verificar')
    .description('Verificar configuración y dependencias')
    .action(async () => {
      await verificarSistema();
    });

  programa
    .command('estadisticas')
    .description('Mostrar estadísticas de Google Sheets')
    .action(async () => {
      try {
        const sheetsService = new GoogleSheetsService();
        const stats = await sheetsService.obtenerEstadisticas();
        
        Logger.titulo('Estadísticas de Google Sheets:');
        console.log(`📊 Total de personas: ${stats.totalPersonas}`);
        console.log(`📧 Con email: ${stats.personasConEmail}`);
        console.log(`📱 Con teléfono: ${stats.personasConTelefono}`);
        console.log(`📚 Cursos únicos: ${stats.cursosUnicos.length}`);
        
        if (stats.cursosUnicos.length > 0) {
          console.log(`   ${stats.cursosUnicos.join(', ')}`);
        }
      } catch (error) {
        Logger.error(`Error obteniendo estadísticas: ${error}`);
      }
    });

  return programa;
}

// ====================== MODO INTERACTIVO ======================
async function modoInteractivo(): Promise<void> {
  console.clear();
  Logger.titulo('🚀 GENERADOR MASIVO DE DOCUMENTOS - MODO INTERACTIVO');
  
  // Paso 1: Seleccionar modo de ejecución
  const { modo } = await inquirer.prompt([
    {
      type: 'list',
      name: 'modo',
      message: '¿Qué quieres hacer?',
      choices: [
        { name: '📄 Procesar TODAS las personas', value: 'todos' },
        { name: '📊 Procesar un RANGO específico', value: 'rango' },
        { name: '👤 Procesar UNA PERSONA específica', value: 'especifico' },
        { name: '🔍 Ver información del sistema', value: 'info' }
      ]
    }
  ]);

  if (modo === 'info') {
    await verificarSistema();
    return;
  }

  // Paso 2: Configuración específica según el modo
  let opciones: OpcionesEjecucion = { modo, tipoSalida: 'ambos' };

  if (modo === 'rango') {
    const respuestas = await inquirer.prompt([
      {
        type: 'number',
        name: 'inicio',
        message: 'Número de persona inicial (desde):',
        default: 1,
        validate: (input) => input >= 1 || 'Debe ser mayor a 0'
      },
      {
        type: 'number',
        name: 'fin',
        message: 'Número de persona final (hasta):',
        default: 10,
        validate: (input) => input >= 1 || 'Debe ser mayor a 0'
      }
    ]);

    opciones.rangoInicio = respuestas.inicio;
    opciones.rangoFin = respuestas.fin;
  }

  if (modo === 'especifico') {
    const { nombre } = await inquirer.prompt([
      {
        type: 'input',
        name: 'nombre',
        message: 'Nombre de la persona (o parte del nombre):',
        validate: (input) => input.trim().length > 0 || 'Debes ingresar un nombre'
      }
    ]);

    opciones.nombreEspecifico = nombre;
  }

  // Paso 3: Tipo de salida
  const { tipoSalida } = await inquirer.prompt([
    {
      type: 'list',
      name: 'tipoSalida',
      message: '¿Qué formato de documentos quieres generar?',
      choices: [
        { name: '🔄 Ambos (Originales + PDF)', value: 'ambos' },
        { name: '📄 Solo PDF (más rápido)', value: 'solo-pdf' },
        { name: '📝 Solo documentos originales (.docx/.xlsx)', value: 'solo-originals' }
      ]
    }
  ]);

  opciones.tipoSalida = tipoSalida;

  // Paso 4: Selección de plantillas (opcional)
  const { usarTodasLasPlantillas } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'usarTodasLasPlantillas',
      message: '¿Usar todas las plantillas disponibles?',
      default: true
    }
  ]);

  if (!usarTodasLasPlantillas) {
    const { plantillasSeleccionadas } = await inquirer.prompt([
      {
        type: 'checkbox',
        name: 'plantillasSeleccionadas',
        message: 'Selecciona las plantillas a usar:',
        choices: CONFIG.PLANTILLAS.map(p => ({
          name: `${p.nombre} (${p.tipo}) - ${p.descripcion || 'Sin descripción'}`,
          value: p.nombre
        })),
        validate: (input) => input.length > 0 || 'Debes seleccionar al menos una plantilla'
      }
    ]);

    opciones.plantillasEspecificas = plantillasSeleccionadas;
  }

  // Paso 5: Confirmación
  Logger.separador();
  Logger.info('Configuración seleccionada:');
  console.log(`📋 Modo: ${modo}`);
  console.log(`📄 Tipo salida: ${tipoSalida}`);
  if (opciones.rangoInicio && opciones.rangoFin) {
    console.log(`📊 Rango: ${opciones.rangoInicio} a ${opciones.rangoFin}`);
  }
  if (opciones.nombreEspecifico) {
    console.log(`👤 Persona: ${opciones.nombreEspecifico}`);
  }
  if (opciones.plantillasEspecificas) {
    console.log(`📝 Plantillas: ${opciones.plantillasEspecificas.join(', ')}`);
  }

  const { confirmar } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirmar',
      message: '¿Proceder con esta configuración?',
      default: true
    }
  ]);

  if (!confirmar) {
    Logger.warn('Operación cancelada');
    return;
  }

  // Ejecutar procesamiento
  await ejecutarProcesamiento(opciones);
}

// ====================== EJECUTOR PRINCIPAL ======================
async function ejecutarProcesamiento(opciones: OpcionesEjecucion): Promise<void> {
  try {
    const procesador = new ProcesadorPrincipal();
    await procesador.ejecutar(opciones);
  } catch (error) {
    Logger.error(`Error en procesamiento: ${error}`);
    process.exit(1);
  }
}

// ====================== VERIFICACIÓN DEL SISTEMA ======================
async function verificarSistema(): Promise<void> {
  Logger.titulo('🔍 VERIFICACIÓN DEL SISTEMA');

  // 1. Verificar Google Sheets
  try {
    const sheetsService = new GoogleSheetsService();
    await sheetsService.inicializar();
    Logger.success('✅ Google Sheets: Conectado correctamente');
  } catch (error) {
    Logger.error('❌ Google Sheets: Error de conexión');
    Logger.info('   Verifica el archivo de credenciales JSON');
  }

  // 2. Verificar plantillas
  const docGen = new DocumentGenerator();
  const validacion = docGen.validarPlantillasExisten();
  
  Logger.info(`📁 Plantillas encontradas: ${validacion.existentes.length}`);
  validacion.existentes.forEach(p => Logger.success(`   ✅ ${p}`));
  
  if (validacion.faltantes.length > 0) {
    Logger.warn(`📁 Plantillas faltantes: ${validacion.faltantes.length}`);
    validacion.faltantes.forEach(p => Logger.error(`   ❌ ${p}`));
  }

  // 3. Verificar LibreOffice
  const pdfConverter = PDFConverter.obtenerInstancia();
  const libreOfficeOk = await pdfConverter.verificarLibreOffice();
  
  if (libreOfficeOk) {
    const version = await pdfConverter.obtenerVersionLibreOffice();
    Logger.success(`✅ LibreOffice: ${version || 'Instalado'}`);
  } else {
    Logger.error('❌ LibreOffice: No encontrado');
  }

  // 4. Estadísticas de datos
  try {
    const sheetsService = new GoogleSheetsService();
    const stats = await sheetsService.obtenerEstadisticas();
    Logger.info(`📊 Personas disponibles: ${stats.totalPersonas}`);
  } catch {
    Logger.warn('⚠️ No se pudieron obtener estadísticas');
  }

  Logger.separador();
}