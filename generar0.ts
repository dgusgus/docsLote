import { google } from "googleapis";
import fs from "fs/promises";
import path from "path";
import { spawn } from "child_process";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import ExcelJS from "exceljs";
import pLimit from "p-limit"; // 👈 control de concurrencia

// ====================== CONFIG ======================
const CONFIG = {
  SPREADSHEET_ID: "1sMu2QaY2kAy1h-V0YiKhZ2jD6VnEbYaCRXm72Fj_r58",
  RANGE: "'Hoja1'!C2:E",
  LIMITE_PERSONAS: 10,
  SOFFICE_PATH: "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
  PLANTILLAS: {
    word: "plantillas/Carta certificado de trabajo TEDO 2025.docx",
    excel: "plantillas/06 FORM 006.xlsx",
  },
  OUTPUT_DIR: "docs",
  CONCURRENCY: 2, // 👈 número de procesos en paralelo (solo podemos hacer 1 por que LibreOffice (soffice.exe) en Windows no soporta bien la ejecución en paralelo.)
};

// ====================== HELPERS ======================
function limpiarNombre(nombre: string) {
  return nombre.replace(/[^\w]/g, "_").substring(0, 50);
}

async function convertirAPdf(inputPath: string, outputDir: string) {
  return new Promise<void>((resolve, reject) => {
    const libre = spawn(CONFIG.SOFFICE_PATH, [
      "--headless",
      "--convert-to",
      "pdf",
      inputPath,
      "--outdir",
      outputDir,
    ]);

    libre.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`❌ LibreOffice falló con código ${code}`));
    });
  });
}

async function generarWord(datos: any, plantillaContent: Buffer, salidaDocx: string) {
  const zip = new PizZip(plantillaContent);
  const doc = new Docxtemplater(zip, {
    delimiters: { start: "[[", end: "]]" },
  });
  doc.render(datos);
  const buf = doc.getZip().generate({ type: "nodebuffer" });
  await fs.writeFile(salidaDocx, buf);
}

async function generarExcel(datos: any, plantillaPath: string, salidaXlsx: string) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(plantillaPath);
  const ws = wb.worksheets[0];

  ws.getCell("C8").value = datos.nombre;
  ws.getCell("M8").value = datos.apellido1;
  ws.getCell("U8").value = datos.apellido2;

  await wb.xlsx.writeFile(salidaXlsx);
}

// ====================== MAIN ======================
async function main() {
  // 🔹 1. Leer datos de Google Sheets
  const auth = new google.auth.GoogleAuth({
    keyFile: "./generador-docs-31f4b831a196.json",
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });

  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: CONFIG.RANGE,
  });

  const personas = (res.data.values || [])
    .slice(0, CONFIG.LIMITE_PERSONAS)
    .map((r) => ({
      apellido2: r[0],
      nombre: r[1],
      apellido1: r[2],
    }));

  console.log(`📊 Total personas a procesar: ${personas.length}`);

  // 🔹 2. Pre-cargar plantillas en memoria (optimización)
  const plantillaWord = await fs.readFile(CONFIG.PLANTILLAS.word);

  // 🔹 3. Concurrencia controlada
  const limit = pLimit(CONFIG.CONCURRENCY);

  await Promise.all(
    personas.map((persona, idx) =>
      limit(async () => {
        try {
          const nombreBase = limpiarNombre(persona.nombre);
          const dirUsuario = path.join(CONFIG.OUTPUT_DIR, nombreBase);
          await fs.mkdir(dirUsuario, { recursive: true });

          // Word → PDF
          const wordTemp = path.join(dirUsuario, "certificado.docx");
          await generarWord(persona, plantillaWord, wordTemp);
          await convertirAPdf(wordTemp, dirUsuario);

          // Excel → PDF
          const excelTemp = path.join(dirUsuario, "credenciales.xlsx");
          await generarExcel(persona, CONFIG.PLANTILLAS.excel, excelTemp);
          await convertirAPdf(excelTemp, dirUsuario);

          console.log(`✅ (${idx + 1}/${personas.length}) ${persona.nombre}`);
        } catch (err) {
          console.error(`⚠️ Error con ${persona.nombre}:`, err);
        }
      })
    )
  );

  console.log("🎉 Proceso completado!");
}

main().catch((err) => console.error("❌ Error global:", err));