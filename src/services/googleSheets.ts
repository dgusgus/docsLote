import { google } from 'googleapis';
import { Persona } from '../types/index.js';
import { CONFIG, PATHS, TIMEOUTS } from '../config/settings.js';
import { Logger } from '../utils/fileUtils.js';

// ====================== SERVICIO GOOGLE SHEETS ======================
export class GoogleSheetsService {
  private auth: any;
  private sheets: any;

  async inicializar(): Promise<void> {
    try {
      this.auth = new google.auth.GoogleAuth({
        keyFile: PATHS.CREDENTIALS,
        scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
      });

      this.sheets = google.sheets({ version: "v4", auth: this.auth });
      Logger.success("Conexión con Google Sheets establecida");
    } catch (error) {
      throw new Error(`Error inicializando Google Sheets: ${error}`);
    }
  }

  async obtenerPersonas(limite?: number, rangoInicio?: number, rangoFin?: number): Promise<Persona[]> {
    if (!this.sheets) {
      await this.inicializar();
    }

    try {
      Logger.progress("Obteniendo datos de Google Sheets...");
      
      const response = await Promise.race([
        this.sheets.spreadsheets.values.get({
          spreadsheetId: CONFIG.SPREADSHEET_ID,
          range: CONFIG.RANGE,
        }),
        new Promise((_, reject) => 
          setTimeout(() => reject(new Error('Timeout en Google Sheets')), TIMEOUTS.GOOGLE_SHEETS)
        )
      ]);

      const filas = response.data.values || [];
      Logger.info(`Filas obtenidas de Google Sheets: ${filas.length}`);

      // Filtrar filas válidas
      let personas = filas
        .filter((fila: any[]) => fila[1] && fila[1].trim()) // Debe tener nombre
        .map((fila: any[], indice: number): Persona => ({
          indice: indice + 1, // Para referencia
          apellido2: fila[0]?.toString().trim() || "",
          nombre: fila[1]?.toString().trim() || "",
          apellido1: fila[2]?.toString().trim() || "",
          email: fila[3]?.toString().trim() || "",
          telefono: fila[4]?.toString().trim() || "",
          documento: fila[5]?.toString().trim() || "",
          curso: fila[6]?.toString().trim() || "",
          fecha_inicio: fila[7]?.toString().trim() || "",
          fecha_fin: fila[8]?.toString().trim() || "",
        }));

      // Aplicar filtros
      if (rangoInicio !== undefined && rangoFin !== undefined) {
        personas = personas.slice(rangoInicio - 1, rangoFin);
        Logger.info(`Aplicado filtro de rango: ${rangoInicio} a ${rangoFin}`);
      } else if (limite) {
        personas = personas.slice(0, limite);
        Logger.info(`Aplicado límite: ${limite} personas`);
      }

      Logger.success(`Personas procesadas: ${personas.length}`);
      return personas;

    } catch (error) {
      throw new Error(`Error obteniendo datos de Google Sheets: ${error}`);
    }
  }

  async obtenerPersonaPorNombre(nombre: string): Promise<Persona | null> {
    const todasLasPersonas = await this.obtenerPersonas();
    
    const personaEncontrada = todasLasPersonas.find(persona => 
      persona.nombre.toLowerCase().includes(nombre.toLowerCase()) ||
      `${persona.nombre} ${persona.apellido1}`.toLowerCase().includes(nombre.toLowerCase())
    );

    if (personaEncontrada) {
      Logger.success(`Persona encontrada: ${personaEncontrada.nombre} ${personaEncontrada.apellido1}`);
    } else {
      Logger.warn(`No se encontró persona con nombre similar a: ${nombre}`);
    }

    return personaEncontrada || null;
  }

  async listarPersonas(limite: number = 10): Promise<void> {
    const personas = await this.obtenerPersonas(limite);
    
    Logger.titulo("Lista de personas disponibles:");
    personas.forEach((persona, index) => {
      console.log(`${index + 1}. ${persona.nombre} ${persona.apellido1} ${persona.apellido2}`.trim());
    });
  }

  // ====================== MÉTODOS DE BÚSQUEDA ======================
  async buscarPersonasPorCriterio(criterio: string, valor: string): Promise<Persona[]> {
    const todasLasPersonas = await this.obtenerPersonas();
    
    return todasLasPersonas.filter(persona => {
      const valorCampo = persona[criterio as keyof Persona];
      return valorCampo?.toString().toLowerCase().includes(valor.toLowerCase());
    });
  }

  async obtenerEstadisticas(): Promise<{
    totalPersonas: number;
    personasConEmail: number;
    personasConTelefono: number;
    cursosUnicos: string[];
  }> {
    const personas = await this.obtenerPersonas();
    
    return {
      totalPersonas: personas.length,
      personasConEmail: personas.filter(p => p.email).length,
      personasConTelefono: personas.filter(p => p.telefono).length,
      cursosUnicos: [...new Set(personas.map(p => p.curso).filter((curso): curso is string => typeof curso === 'string' && !!curso))]
    };
  }
}