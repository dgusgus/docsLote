import { google } from "googleapis";
import fs from "fs/promises";
import path from "path";
import { spawn } from "child_process";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import ExcelJS from "exceljs";

// ====================== TIPOS ======================
type TipoSalida = "solo-pdf" | "solo-originals" | "ambos";
type TipoPlantilla = "word" | "excel";

interface Persona {
  apellido2: string;
  nombre: string;
  apellido1: string;
  [key: string]: any; // Para campos adicionales dinámicos
}

interface PlantillaConfig {
  nombre: string;
  archivo: string;
  tipo: TipoPlantilla;
  campos?: string[]; // Campos específicos que usa esta plantilla
}

// ====================== CONFIG MEJORADA ======================
const CONFIG = {
  SPREADSHEET_ID: "1sMu2QaY2kAy1h-V0YiKhZ2jD6VnEbYaCRXm72Fj_r58",
  RANGE: "'Hoja1'!C2:Z", // Ampliado para más campos
  LIMITE_PERSONAS: 5, // Para pruebas, después cambiar a 500
  SOFFICE_PATH: "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
  
  // 👈 PLANTILLAS ORGANIZADAS POR CATEGORÍAS
  PLANTILLAS: [
    { nombre: "certificado_trabajo", archivo: "plantillas/Carta certificado de trabajo TEDO 2025.docx", tipo: "word" as TipoPlantilla },
    { nombre: "credenciales", archivo: "plantillas/06 FORM 006.xlsx", tipo: "excel" as TipoPlantilla },
    // Agrega más plantillas aquí según necesites
  ] as PlantillaConfig[],
  
  OUTPUT_DIR: "docs_masivos",
  
  // 👈 CONFIGURACIÓN DE SALIDA
  TIPO_SALIDA: "solo-originals" as TipoSalida, // "solo-pdf" | "solo-originals" | "ambos"
  
  // 👈 CONFIGURACIÓN DE LIMPIEZA
  LIMPIAR_TEMPORALES: true,
  
  // 👈 CONFIGURACIÓN SIMPLE SIN CONCURRENCIA COMPLEJA
  PROCESO_SECUENCIAL: true, // LibreOffice funciona mejor así
};

// ====================== CACHE DE PLANTILLAS ======================
const plantillaCache = new Map<string, Buffer>();

async function precargarPlantillas(): Promise<void> {
  console.log("📂 Precargando plantillas en memoria...");
  
  for (const plantilla of CONFIG.PLANTILLAS) {
    try {
      const content = await fs.readFile(plantilla.archivo);
      plantillaCache.set(plantilla.nombre, content);
      console.log(`  ✅ ${plantilla.nombre}`);
    } catch (err) {
      console.warn(`  ⚠️ No se pudo cargar ${plantilla.nombre}: ${err}`);
    }
  }
  console.log(`📂 ${plantillaCache.size} plantillas cargadas\n`);
}

// ====================== HELPERS MEJORADOS ======================
function limpiarNombre(nombre: string): string {
  return nombre
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Quita acentos
    .replace(/[^\w\s-]/g, "") // Solo letras, números, espacios y guiones
    .replace(/\s+/g, "_") // Espacios → guiones bajos
    .substring(0, 50);
}

async function verificarLibreOffice(): Promise<boolean> {
  try {
    await fs.access(CONFIG.SOFFICE_PATH);
    return true;
  } catch {
    console.error(`❌ LibreOffice no encontrado en: ${CONFIG.SOFFICE_PATH}`);
    console.log("💡 Instala LibreOffice o ajusta SOFFICE_PATH en CONFIG");
    return false;
  }
}

async function convertirAPdf(inputPath: string, outputDir: string): Promise<void> {
  return new Promise((resolve, reject) => {
    // 👈 TIMEOUT MÁS LARGO PARA ARCHIVOS GRANDES
    const timeout = setTimeout(() => {
      reject(new Error(`⏰ Timeout convirtiendo: ${path.basename(inputPath)}`));
    }, 45000); // 45 segundos

    const libre = spawn(CONFIG.SOFFICE_PATH, [
      "--headless",
      "--convert-to", "pdf",
      "--outdir", outputDir,
      inputPath,
    ], {
      stdio: "ignore" // 👈 Evita logs innecesarios de LibreOffice
    });

    libre.on("exit", (code) => {
      clearTimeout(timeout);
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`❌ LibreOffice falló con código ${code} para ${path.basename(inputPath)}`));
      }
    });

    libre.on("error", (err) => {
      clearTimeout(timeout);
      reject(new Error(`❌ Error ejecutando LibreOffice: ${err.message}`));
    });
  });
}

// ====================== GENERADORES ESPECIALIZADOS ======================
async function generarWord(
  datos: Persona,
  plantillaNombre: string,
  salidaPath: string
): Promise<void> {
  const plantillaContent = plantillaCache.get(plantillaNombre);
  if (!plantillaContent) {
    throw new Error(`❌ Plantilla ${plantillaNombre} no encontrada en cache`);
  }

  try {
    const zip = new PizZip(plantillaContent);
    const doc = new Docxtemplater(zip, {
      delimiters: { start: "[[", end: "]]" },
      errorLogging: false,
    });

    // 👈 DATOS ENRIQUECIDOS
    const datosCompletos = {
      ...datos,
      fecha_actual: new Date().toLocaleDateString('es-ES'),
      año_actual: new Date().getFullYear(),
      nombre_completo: `${datos.nombre} ${datos.apellido1} ${datos.apellido2}`.trim(),
    };

    doc.render(datosCompletos);
    const buffer = doc.getZip().generate({ type: "nodebuffer" });
    await fs.writeFile(salidaPath, buffer);
  } catch (err) {
    throw new Error(`❌ Error generando Word ${plantillaNombre}: ${err}`);
  }
}

async function generarExcel(
  datos: Persona,
  plantillaPath: string,
  salidaPath: string,
  plantillaNombre: string
): Promise<void> {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(plantillaPath);
    const worksheet = workbook.worksheets[0];

    // 👈 MAPEO ESPECÍFICO POR PLANTILLA
    if (plantillaNombre === "credenciales") {
      worksheet.getCell("C8").value = datos.nombre;
      worksheet.getCell("M8").value = datos.apellido1;
      worksheet.getCell("U8").value = datos.apellido2;
    }
    // Agrega más mapeos según tus plantillas

    // 👈 CAMPOS COMUNES
    const nombreCompleto = `${datos.nombre} ${datos.apellido1} ${datos.apellido2}`.trim();
    if (worksheet.getCell("A1")) {
      worksheet.getCell("A1").value = nombreCompleto;
    }

    await workbook.xlsx.writeFile(salidaPath);
  } catch (err) {
    throw new Error(`❌ Error generando Excel ${plantillaNombre}: ${err}`);
  }
}

// ====================== PROCESAMIENTO INDIVIDUAL ======================
async function procesarPersona(
  persona: Persona,
  indice: number,
  total: number
): Promise<{ exitos: number; errores: number }> {
  let exitos = 0;
  let errores = 0;

  const nombreBase = limpiarNombre(persona.nombre);
  const dirUsuario = path.join(CONFIG.OUTPUT_DIR, nombreBase);
  
  console.log(`📄 (${indice + 1}/${total}) Procesando: ${persona.nombre}`);
  
  try {
    await fs.mkdir(dirUsuario, { recursive: true });

    // 👈 PROCESAR CADA PLANTILLA DE FORMA SECUENCIAL
    for (const plantilla of CONFIG.PLANTILLAS) {
      try {
        const extension = plantilla.tipo === "word" ? "docx" : "xlsx";
        const archivoOriginal = path.join(dirUsuario, `${plantilla.nombre}.${extension}`);
        
        console.log(`  📝 Generando ${plantilla.nombre}...`);

        // 👈 GENERAR DOCUMENTO ORIGINAL
        if (["solo-originals", "ambos"].includes(CONFIG.TIPO_SALIDA)) {
          if (plantilla.tipo === "word") {
            await generarWord(persona, plantilla.nombre, archivoOriginal);
          } else {
            await generarExcel(persona, plantilla.archivo, archivoOriginal, plantilla.nombre);
          }
        }

        // 👈 CONVERTIR A PDF
        if (["solo-pdf", "ambos"].includes(CONFIG.TIPO_SALIDA)) {
          // Si solo queremos PDF, generamos el original temporal
          if (CONFIG.TIPO_SALIDA === "solo-pdf") {
            if (plantilla.tipo === "word") {
              await generarWord(persona, plantilla.nombre, archivoOriginal);
            } else {
              await generarExcel(persona, plantilla.archivo, archivoOriginal, plantilla.nombre);
            }
          }

          console.log(`  🔄 Convirtiendo a PDF...`);
          await convertirAPdf(archivoOriginal, dirUsuario);

          // 👈 LIMPIAR TEMPORAL SI SOLO QUEREMOS PDF
          if (CONFIG.TIPO_SALIDA === "solo-pdf" && CONFIG.LIMPIAR_TEMPORALES) {
            try {
              await fs.unlink(archivoOriginal);
            } catch {} // Ignorar si no existe
          }
        }

        console.log(`  ✅ ${plantilla.nombre} completado`);
        exitos++;

        // 👈 PEQUEÑA PAUSA PARA NO SATURAR LibreOffice
        await new Promise(resolve => setTimeout(resolve, 500));
        
      } catch (err) {
        console.error(`  ❌ Error en ${plantilla.nombre}:`, err);
        errores++;
      }
    }

  } catch (err) {
    console.error(`❌ Error general procesando ${persona.nombre}:`, err);
    errores += CONFIG.PLANTILLAS.length;
  }

  return { exitos, errores };
}

// ====================== FUNCIÓN PRINCIPAL ======================
async function main() {
  console.log("🚀 GENERADOR MASIVO DE DOCUMENTOS v2.0");
  console.log(`📊 Configuración: ${CONFIG.LIMITE_PERSONAS} personas × ${CONFIG.PLANTILLAS.length} plantillas`);
  console.log(`📁 Tipo salida: ${CONFIG.TIPO_SALIDA}`);
  console.log(`🔧 Procesamiento: ${CONFIG.PROCESO_SECUENCIAL ? 'Secuencial' : 'Paralelo'}\n`);

  // 👈 VERIFICACIONES INICIALES
  if (CONFIG.TIPO_SALIDA !== "solo-originals") {
    if (!(await verificarLibreOffice())) {
      console.log("💡 Cambia TIPO_SALIDA a 'solo-originals' si no tienes LibreOffice");
      return;
    }
  }

  await fs.mkdir(CONFIG.OUTPUT_DIR, { recursive: true });
  await precargarPlantillas();

  // 👈 OBTENER DATOS DE GOOGLE SHEETS
  console.log("📡 Conectando con Google Sheets...");
  try {
    const auth = new google.auth.GoogleAuth({
      keyFile: "./generador-docs-31f4b831a196.json",
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });

    const sheets = google.sheets({ version: "v4", auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: CONFIG.RANGE,
    });

    const personas = (response.data.values || [])
      .slice(0, CONFIG.LIMITE_PERSONAS)
      .filter(fila => fila[1]) // Solo filas con nombre
      .map((fila): Persona => ({
        apellido2: fila[0] || "",
        nombre: fila[1] || "",
        apellido1: fila[2] || "",
        // 👈 CAMPOS ADICIONALES DINÁMICOS
        email: fila[3] || "",
        telefono: fila[4] || "",
        documento: fila[5] || "",
        curso: fila[6] || "",
        fecha_inicio: fila[7] || "",
        fecha_fin: fila[8] || "",
      }));

    console.log(`📊 Personas encontradas: ${personas.length}`);
    console.log(`📄 Total documentos a generar: ${personas.length * CONFIG.PLANTILLAS.length}\n`);

    if (personas.length === 0) {
      console.log("❌ No se encontraron personas para procesar");
      return;
    }

    // 👈 PROCESAMIENTO
    const inicioTiempo = Date.now();
    let totalExitos = 0;
    let totalErrores = 0;

    for (let i = 0; i < personas.length; i++) {
      const resultado = await procesarPersona(personas[i], i, personas.length);
      totalExitos += resultado.exitos;
      totalErrores += resultado.errores;
    }

    // 👈 ESTADÍSTICAS FINALES
    const tiempoTotal = (Date.now() - inicioTiempo) / 1000;
    console.log("\n🎉 PROCESO COMPLETADO!");
    console.log(`✅ Documentos generados exitosamente: ${totalExitos}`);
    console.log(`❌ Errores: ${totalErrores}`);
    console.log(`⏱️ Tiempo total: ${Math.round(tiempoTotal)}s`);
    console.log(`📁 Resultados guardados en: ${CONFIG.OUTPUT_DIR}`);
    
    if (totalExitos > 0) {
      console.log(`📊 Velocidad promedio: ${(totalExitos / tiempoTotal).toFixed(1)} documentos/segundo`);
    }

  } catch (err) {
    console.error("❌ Error conectando con Google Sheets:", err);
    console.log("💡 Verifica:");
    console.log("   - El archivo 'generador-docs-31f4b831a196.json' existe");
    console.log("   - El SPREADSHEET_ID es correcto");
    console.log("   - Tienes permisos de lectura en la hoja");
  }
}

// ====================== EJECUCIÓN ======================
main().catch(err => {
  console.error("❌ Error crítico:", err);
  process.exit(1);
});