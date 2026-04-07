import { Command } from 'commander';
import inquirer from 'inquirer';
import chalk from 'chalk';
import { spawn } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleSheetsService } from '../services/googleSheets.js';
import { DocumentGenerator } from '../services/documentGenerator.js';
import { PDFConverter } from '../services/pdfConverter.js';
import { CONFIG, CONFIGURACIONES_RAPIDAS } from '../config/settings.js';
import { Logger } from '../utils/fileUtils.js';
import { OpcionesEjecucion, TipoSalida, ModoEjecucion } from '../types/index.js';
import { ProcesadorPrincipal } from './processor.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);
const projectRoot = path.join(__dirname, '../..');

// ====================== TIPOS ======================

type TipoAsistencia = 'solo-excel' | 'solo-pdf' | 'solo-merged' | 'todos';

// ====================== RUNNER AUXILIAR ======================
// Lanza generarQR.ts y generarAsistencia.ts como subprocesos.
// Usa process.execPath (node.exe) + tsx como modulo ESM para evitar
// problemas con rutas que tienen espacios en Windows.

function correrScript(scriptName: string, args: string[] = []): Promise<void> {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(projectRoot, scriptName);

    // Usamos el node.exe que ya esta corriendo (sin espacios en su ruta)
    // y cargamos tsx como modulo ESM pasando su entry point directamente.
    // Esto evita el EINVAL de Windows cuando la carpeta del proyecto tiene espacios.
    const tsxEntry = path.join(projectRoot, 'node_modules', 'tsx', 'dist', 'cli.mjs');
    const nodeExe  = process.execPath;   // ruta absoluta al node actual, sin espacios

    const proc = spawn(nodeExe, [tsxEntry, scriptPath, ...args], {
      stdio: 'inherit',
      cwd: projectRoot,
      shell: false,
    });

    proc.on('exit', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${scriptName} termino con codigo ${code}`));
    });
    proc.on('error', (err) => reject(new Error(`Error ejecutando ${scriptName}: ${err.message}`)));
  });
}

// ====================== COMANDOS CLI ======================
export function crearComandos(): Command {
  const programa = new Command();

  programa
    .name('generador-docs')
    .description('Generador masivo de documentos desde Google Sheets')
    .version('2.0.0');

  // ── INTERACTIVO ────────────────────────────────────────────────
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

  // ── DOCUMENTOS: TODOS ──────────────────────────────────────────
  programa
    .command('todos')
    .description('Procesar todas las personas con todas las plantillas')
    .option('-t, --tipo <tipo>', 'Tipo de salida (solo-pdf, solo-originals, ambos)', 'ambos')
    .option('-l, --limite <numero>', 'Límite de personas a procesar', '500')
    .option('-o, --output <directorio>', 'Directorio de salida personalizado')
    .action(async (opciones) => {
      if (opciones.limite) CONFIG.LIMITE_PERSONAS = parseInt(opciones.limite);
      await ejecutarProcesamiento({
        modo: 'todos',
        tipoSalida: opciones.tipo as TipoSalida,
        outputDir: opciones.output,
      });
    });

  // ── DOCUMENTOS: RANGO ─────────────────────────────────────────
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
      const fin    = parseInt(opciones.fin);
      if (inicio < 1 || fin < inicio) {
        Logger.error('El rango debe ser válido (inicio >= 1, fin >= inicio)');
        process.exit(1);
      }
      await ejecutarProcesamiento({
        modo: 'rango',
        tipoSalida: opciones.tipo as TipoSalida,
        rangoInicio: inicio,
        rangoFin: fin,
        plantillasEspecificas: opciones.plantillas?.split(','),
        outputDir: opciones.output,
      });
    });

  // ── DOCUMENTOS: PERSONA ────────────────────────────────────────
  programa
    .command('persona')
    .description('Procesar una persona específica')
    .requiredOption('-n, --nombre <nombre>', 'Nombre o parte del nombre de la persona')
    .option('-t, --tipo <tipo>', 'Tipo de salida (solo-pdf, solo-originals, ambos)', 'ambos')
    .option('-p, --plantillas <nombres>', 'Plantillas específicas (separadas por coma)')
    .option('-o, --output <directorio>', 'Directorio de salida personalizado')
    .action(async (opciones) => {
      await ejecutarProcesamiento({
        modo: 'especifico',
        tipoSalida: opciones.tipo as TipoSalida,
        nombreEspecifico: opciones.nombre,
        plantillasEspecificas: opciones.plantillas?.split(','),
        outputDir: opciones.output,
      });
    });

  // ── QR ────────────────────────────────────────────────────────
  programa
    .command('qr')
    .description('Generar códigos QR desde Google Sheets')
    .option('-m, --modo <modo>', 'todos | rango | nombre', 'todos')
    .option('-i, --inicio <n>', 'Inicio del rango')
    .option('-f, --fin <n>', 'Fin del rango')
    .option('-n, --nombre <nombre>', 'Nombre de la persona')
    .option('-o, --output <directorio>', 'Directorio de salida')
    .action(async (opciones) => {
      const args: string[] = [];
      if (opciones.modo === 'rango' && opciones.inicio && opciones.fin) {
        args.push('--rango', opciones.inicio, opciones.fin);
      } else if (opciones.modo === 'nombre' && opciones.nombre) {
        args.push('--nombre', opciones.nombre);
      }
      if (opciones.output) args.push('--output', opciones.output);
      try {
        await correrScript('generarQR.ts', args);
      } catch (error) {
        Logger.error(`${error}`);
        process.exit(1);
      }
    });

  // ── ASISTENCIA ────────────────────────────────────────────────
  programa
    .command('asistencia')
    .description('Generar registros de asistencia SIREPRE 2026')
    .option('-g, --grupo <n>', 'Procesar solo ese grupo')
    .option('-t, --tipo <tipo>', 'solo-excel | solo-pdf | solo-merged | todos', 'solo-merged')
    .option('-o, --output <directorio>', 'Directorio de salida')
    .action(async (opciones) => {
      const args: string[] = [];
      if (opciones.grupo)  args.push('--grupo',  opciones.grupo);
      if (opciones.tipo)   args.push('--tipo',   opciones.tipo);
      if (opciones.output) args.push('--output', opciones.output);
      try {
        await correrScript('generarAsistencia.ts', args);
      } catch (error) {
        Logger.error(`${error}`);
        process.exit(1);
      }
    });

  // ── INFORMACIÓN ────────────────────────────────────────────────
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
        const qrTag = plantilla.qr ? chalk.green(' [QR]') : '';
        console.log(`${index + 1}. ${chalk.cyan(plantilla.nombre)}${qrTag} (${plantilla.tipo.toUpperCase()})`);
        console.log(`   📁 ${plantilla.archivo}`);
        if (plantilla.descripcion) console.log(`   📝 ${plantilla.descripcion}`);
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
        console.log(`📧 Con email:         ${stats.personasConEmail}`);
        console.log(`📱 Con teléfono:      ${stats.personasConTelefono}`);
        console.log(`📚 Cursos únicos:     ${stats.cursosUnicos.length}`);
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
  Logger.titulo('🚀 GENERADOR MASIVO DE DOCUMENTOS — MENÚ PRINCIPAL');

  const { seccion } = await inquirer.prompt([
    {
      type: 'list',
      name: 'seccion',
      message: '¿Qué querés hacer?',
      choices: [
        { name: '📄  Generar documentos (Word/Excel/PDF por persona)', value: 'documentos' },
        { name: '📋  Generar registros de asistencia',                 value: 'asistencia' },
        { name: '📱  Generar códigos QR',                              value: 'qr'         },
        new inquirer.Separator(),
        { name: '🔍  Información y verificación del sistema',          value: 'info'       },
      ],
    },
  ]);

  if (seccion === 'info')       { await menuInfo();        return; }
  if (seccion === 'qr')         { await menuQR();          return; }
  if (seccion === 'asistencia') { await menuAsistencia();  return; }
  if (seccion === 'documentos') { await menuDocumentos();  return; }
}

// ─────────────────────────────────────────────────────────────
// SUB-MENÚ: DOCUMENTOS
// ─────────────────────────────────────────────────────────────
async function menuDocumentos(): Promise<void> {
  const { modo } = await inquirer.prompt([
    {
      type: 'list',
      name: 'modo',
      message: '¿A quién querés procesar?',
      choices: [
        { name: '👥  Todas las personas',          value: 'todos'     },
        { name: '📊  Un rango de personas',        value: 'rango'     },
        { name: '👤  Una persona específica',      value: 'especifico' },
      ],
    },
  ]);

  let opciones: OpcionesEjecucion = { modo, tipoSalida: 'ambos' };

  if (modo === 'rango') {
    const r = await inquirer.prompt([
      { type: 'number', name: 'inicio', message: 'Desde (número de persona):', default: 1,
        validate: (v) => v >= 1 || 'Debe ser ≥ 1' },
      { type: 'number', name: 'fin',    message: 'Hasta (número de persona):', default: 10,
        validate: (v) => v >= 1 || 'Debe ser ≥ 1' },
    ]);
    opciones.rangoInicio = r.inicio;
    opciones.rangoFin    = r.fin;
  }

  if (modo === 'especifico') {
    const { nombre } = await inquirer.prompt([
      { type: 'input', name: 'nombre', message: 'Nombre (o parte):',
        validate: (v) => v.trim().length > 0 || 'Ingresá un nombre' },
    ]);
    opciones.nombreEspecifico = nombre;
  }

  const { tipoSalida } = await inquirer.prompt([
    {
      type: 'list',
      name: 'tipoSalida',
      message: '¿Qué formato querés generar?',
      choices: [
        { name: '🔄  Ambos (.docx/.xlsx  +  PDF)',      value: 'ambos'          },
        { name: '📄  Solo PDF',                          value: 'solo-pdf'       },
        { name: '📝  Solo originales (.docx / .xlsx)',   value: 'solo-originals' },
      ],
    },
  ]);
  opciones.tipoSalida = tipoSalida;

  const { usarTodas } = await inquirer.prompt([
    { type: 'confirm', name: 'usarTodas', message: '¿Usar todas las plantillas disponibles?', default: true },
  ]);

  if (!usarTodas) {
    const { seleccion } = await inquirer.prompt([
      {
        type: 'checkbox',
        name: 'seleccion',
        message: 'Seleccioná las plantillas:',
        choices: CONFIG.PLANTILLAS.map(p => ({
          name: `${p.nombre} (${p.tipo})${p.descripcion ? ' — ' + p.descripcion : ''}`,
          value: p.nombre,
        })),
        validate: (v) => v.length > 0 || 'Seleccioná al menos una',
      },
    ]);
    opciones.plantillasEspecificas = seleccion;
  }

  await confirmarYEjecutar('documentos', opciones);
}

// ─────────────────────────────────────────────────────────────
// SUB-MENÚ: ASISTENCIA
// ─────────────────────────────────────────────────────────────
async function menuAsistencia(): Promise<void> {
  const { alcance } = await inquirer.prompt([
    {
      type: 'list',
      name: 'alcance',
      message: '¿Qué grupos querés procesar?',
      choices: [
        { name: '👥  Todos los grupos', value: 'todos' },
        { name: '📌  Un grupo específico', value: 'uno' },
      ],
    },
  ]);

  let grupo: string | undefined;
  if (alcance === 'uno') {
    const r = await inquirer.prompt([
      { type: 'input', name: 'grupo', message: 'Número de grupo:',
        validate: (v) => v.trim().length > 0 || 'Ingresá un número de grupo' },
    ]);
    grupo = r.grupo.trim();
  }

  const { tipo } = await inquirer.prompt([
    {
      type: 'list',
      name: 'tipo',
      message: '¿Qué archivos querés generar?',
      choices: [
        { name: '📊  Solo Excel (.xlsx por fecha)',                              value: 'solo-excel'  },
        { name: '📄  Solo PDF por fecha (un PDF por fecha)',                     value: 'solo-pdf'    },
        { name: '📎  Solo PDF unificado por grupo  (todas las fechas juntas)',   value: 'solo-merged' },
        { name: '📦  Todo  (.xlsx  +  PDF por fecha  +  PDF unificado)',         value: 'todos'       },
      ],
    },
  ]);

  // Resumen antes de ejecutar
  Logger.separador();
  Logger.info('Configuración:');
  console.log(`  Grupos:  ${grupo ? `grupo ${grupo}` : 'todos'}`);
  console.log(`  Tipo:    ${tipo}`);

  const { confirmar } = await inquirer.prompt([
    { type: 'confirm', name: 'confirmar', message: '¿Proceder?', default: true },
  ]);
  if (!confirmar) { Logger.warn('Operación cancelada'); return; }

  const args: string[] = ['--tipo', tipo];
  if (grupo) args.push('--grupo', grupo);

  try {
    await correrScript('generarAsistencia.ts', args);
  } catch (error) {
    Logger.error(`${error}`);
    process.exit(1);
  }
}

// ─────────────────────────────────────────────────────────────
// SUB-MENÚ: QR
// ─────────────────────────────────────────────────────────────
async function menuQR(): Promise<void> {
  const { modo } = await inquirer.prompt([
    {
      type: 'list',
      name: 'modo',
      message: '¿Para quién querés generar QRs?',
      choices: [
        { name: '👥  Todas las personas', value: 'todos'  },
        { name: '📊  Un rango de personas', value: 'rango'  },
        { name: '👤  Una persona específica', value: 'nombre' },
      ],
    },
  ]);

  const args: string[] = [];

  if (modo === 'rango') {
    const r = await inquirer.prompt([
      { type: 'number', name: 'inicio', message: 'Desde:', default: 1,
        validate: (v) => v >= 1 || 'Debe ser ≥ 1' },
      { type: 'number', name: 'fin',    message: 'Hasta:', default: 10,
        validate: (v) => v >= 1 || 'Debe ser ≥ 1' },
    ]);
    args.push('--rango', String(r.inicio), String(r.fin));
  }

  if (modo === 'nombre') {
    const { nombre } = await inquirer.prompt([
      { type: 'input', name: 'nombre', message: 'Nombre (o parte):',
        validate: (v) => v.trim().length > 0 || 'Ingresá un nombre' },
    ]);
    args.push('--nombre', nombre);
  }

  Logger.separador();
  Logger.info(`Generando QRs — modo: ${modo}`);

  const { confirmar } = await inquirer.prompt([
    { type: 'confirm', name: 'confirmar', message: '¿Proceder?', default: true },
  ]);
  if (!confirmar) { Logger.warn('Operación cancelada'); return; }

  try {
    await correrScript('generarQR.ts', args);
  } catch (error) {
    Logger.error(`${error}`);
    process.exit(1);
  }
}

// ─────────────────────────────────────────────────────────────
// SUB-MENÚ: INFORMACIÓN
// ─────────────────────────────────────────────────────────────
async function menuInfo(): Promise<void> {
  const { accion } = await inquirer.prompt([
    {
      type: 'list',
      name: 'accion',
      message: '¿Qué información querés ver?',
      choices: [
        { name: '🔍  Verificar configuración y dependencias', value: 'verificar'    },
        { name: '📊  Estadísticas del Google Sheet',          value: 'estadisticas' },
        { name: '👥  Listar personas',                        value: 'listar'       },
        { name: '📝  Ver plantillas configuradas',            value: 'plantillas'   },
      ],
    },
  ]);

  if (accion === 'verificar') {
    await verificarSistema();
    return;
  }

  if (accion === 'estadisticas') {
    try {
      const s = new GoogleSheetsService();
      const stats = await s.obtenerEstadisticas();
      Logger.titulo('Estadísticas de Google Sheets:');
      console.log(`📊 Total de personas: ${stats.totalPersonas}`);
      console.log(`📧 Con email:         ${stats.personasConEmail}`);
      console.log(`📱 Con teléfono:      ${stats.personasConTelefono}`);
      console.log(`📚 Cursos únicos:     ${stats.cursosUnicos.length}`);
    } catch (error) {
      Logger.error(`${error}`);
    }
    return;
  }

  if (accion === 'listar') {
    const { limite } = await inquirer.prompt([
      { type: 'number', name: 'limite', message: '¿Cuántas personas listar?', default: 20 },
    ]);
    try {
      const s = new GoogleSheetsService();
      await s.listarPersonas(limite);
    } catch (error) {
      Logger.error(`${error}`);
    }
    return;
  }

  if (accion === 'plantillas') {
    Logger.titulo('Plantillas configuradas:');
    CONFIG.PLANTILLAS.forEach((p, i) => {
      const qrTag = p.qr ? chalk.green(' [QR]') : '';
      console.log(`${i + 1}. ${chalk.cyan(p.nombre)}${qrTag} (${p.tipo.toUpperCase()})`);
      console.log(`   📁 ${p.archivo}`);
      if (p.descripcion) console.log(`   📝 ${p.descripcion}`);
      console.log();
    });
  }
}

// ====================== HELPERS ======================

async function confirmarYEjecutar(
  tipo: 'documentos',
  opciones: OpcionesEjecucion
): Promise<void> {
  Logger.separador();
  Logger.info('Configuración seleccionada:');
  console.log(`  Modo:      ${opciones.modo}`);
  console.log(`  Formato:   ${opciones.tipoSalida}`);
  if (opciones.rangoInicio) console.log(`  Rango:     ${opciones.rangoInicio} → ${opciones.rangoFin}`);
  if (opciones.nombreEspecifico) console.log(`  Persona:   ${opciones.nombreEspecifico}`);
  if (opciones.plantillasEspecificas) console.log(`  Plantillas: ${opciones.plantillasEspecificas.join(', ')}`);

  const { confirmar } = await inquirer.prompt([
    { type: 'confirm', name: 'confirmar', message: '¿Proceder?', default: true },
  ]);
  if (!confirmar) { Logger.warn('Operación cancelada'); return; }

  await ejecutarProcesamiento(opciones);
}

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

  try {
    const s = new GoogleSheetsService();
    await s.inicializar();
    Logger.success('Google Sheets: conectado');
  } catch {
    Logger.error('Google Sheets: error de conexión — verificá el archivo de credenciales JSON');
  }

  const docGen    = new DocumentGenerator();
  const validacion = docGen.validarPlantillasExisten();
  Logger.info(`Plantillas encontradas: ${validacion.existentes.length}`);
  validacion.existentes.forEach(p => Logger.success(`   ✅ ${p}`));
  if (validacion.faltantes.length > 0) {
    validacion.faltantes.forEach(p => Logger.error(`   ❌ ${p}`));
  }

  const pdfConverter   = PDFConverter.obtenerInstancia();
  const libreOfficeOk  = await pdfConverter.verificarLibreOffice();
  if (libreOfficeOk) {
    const version = await pdfConverter.obtenerVersionLibreOffice();
    Logger.success(`LibreOffice: ${version || 'instalado'}`);
  } else {
    Logger.error('LibreOffice: no encontrado');
  }

  try {
    const s     = new GoogleSheetsService();
    const stats = await s.obtenerEstadisticas();
    Logger.info(`Personas disponibles: ${stats.totalPersonas}`);
  } catch {
    Logger.warn('No se pudieron obtener estadísticas');
  }

  Logger.separador();
}