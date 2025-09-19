// leer-sheets.js
import { google } from "googleapis";

async function main() {
  const auth = new google.auth.GoogleAuth({
    keyFile: "./generador-docs-31f4b831a196.json", // el JSON que descargaste
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });

  const client = await auth.getClient();
  const sheets = google.sheets({ version: "v4", auth: client });

  const SPREADSHEET_ID = "1sMu2QaY2kAy1h-V0YiKhZ2jD6VnEbYaCRXm72Fj_r58"; // lo obtienes de la URL del Sheet
  const RANGE = "'Hoja1'!C2:E"; // ajusta nombre de hoja y rango

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: RANGE,
  });

  const rows = res.data.values || [];
  console.log("Filas leídas:", rows.length);
  console.log(rows); // procesa como necesites
}

main().catch(err => console.error(err));