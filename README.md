generar0.ts 
es la version mas basica pero funiona 
--------------------------------------------------------------------------------------
# 🚀 Generador Masivo de Documentos v2.0

Genera documentos en masa desde Google Sheets usando plantillas Word/Excel con conversión automática a PDF.

## ⚡ Características Principales

- **📊 Integración Google Sheets**: Lee datos directamente desde tus formularios
- **📄 Plantillas dinámicas**: Soporte para Word (.docx) y Excel (.xlsx)
- **🔄 Conversión PDF automática**: Usando LibreOffice headless
- **🎯 Múltiples modos**: Todos, por rango, o persona específica
- **⚙️ CLI intuitiva**: Modo interactivo y comandos directos
- **🔧 Modular**: Código organizado y extensible
- **📊 Reportes detallados**: Estadísticas y manejo de errores

## 📋 Requisitos Previos

### 1. Software Requerido
```bash
# Node.js 18+
node --version  # Debe ser >= 18.0.0

# LibreOffice (para conversión PDF)
# Windows: Descargar desde https://www.libreoffice.org/
# Linux: sudo apt install libreoffice
# Mac: brew install --cask libreoffice
```

### 2. Credenciales Google Sheets
1. Crear proyecto en [Google Cloud Console](https://console.cloud.google.com/)
2. Habilitar Google Sheets API
3. Crear cuenta de servicio y descargar JSON
4. Renombrar archivo a `generador-docs-31f4b831a196.json`

## 🛠️ Instalación Paso a Paso

### Paso 1: Clonar y configurar
```bash
# Clonar el repositorio
git clone <tu-repo>
cd doc-por-lotes

# Instalar dependencias
pnpm install

# Compilar TypeScript
pnpm run build
```

### Paso 2: Configurar plantillas
```bash
# Crear carpeta de plantillas
mkdir plantillas

# Agregar tus archivos .docx y .xlsx
cp /ruta/a/tus/plantillas/* plantillas/
```

### Paso 3: Configurar settings
Editar `src/config/settings.ts`:

```typescript
export const CONFIG = {
  // Tu ID de Google Sheets
  SPREADSHEET_ID: "tu-spreadsheet-id-aqui",
  
  // Rango de datos (ajustar según tu hoja)
  RANGE: "'Hoja1'!C2:Z",
  
  // Ruta LibreOffice (ajustar si es necesario)
  SOFFICE_PATH: "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
  
  // Configurar tus plantillas
  PLANTILLAS: [
    {
      nombre: "certificado",
      archivo: "plantillas/certificado.docx",
      tipo: "word",
      descripcion: "Certificado de participación"
    },
    // ... más plantillas
  ]
};
```

## 🎮 Modos de Uso

### 1. 🎯 Modo Interactivo (Recomendado para principiantes)
```bash
# Ejecutar sin parámetros abre el modo interactivo
pnpm run dev
# o
tsx generar.ts
```

### 2. ⚡ Comandos Directos (Para usuarios avanzados)

#### Procesar todas las personas
```bash
# Todas las plantillas, ambos formatos
tsx generar.ts todos

# Solo PDFs (más rápido)
tsx generar.ts todos --tipo solo-pdf

# Limitar cantidad
tsx generar.ts todos --limite 10

# Directorio personalizado
tsx generar.ts todos --output ./mis_documentos
```

#### Procesar por rango
```bash
# Personas 1 a 50
tsx generar.ts rango --inicio 1 --fin 50

# Con plantillas específicas
tsx generar.ts rango --inicio 1 --fin 10 --plantillas "certificado,diploma"

# Solo originales (sin PDF)
tsx generar.ts rango --inicio 1 --fin 5 --tipo solo-originals
```

#### Procesar persona específica
```bash
# Por nombre
tsx generar.ts persona --nombre "Juan Pérez"

# Nombre parcial
tsx generar.ts persona --nombre "María" --tipo solo-pdf

# Con plantillas específicas
tsx generar.ts persona --nombre "Ana" --plantillas "certificado,credenciales"
```

### 3. 📊 Comandos de Información

```bash
# Listar personas disponibles
tsx generar.ts listar --limite 20

# Ver plantillas configuradas
tsx generar.ts plantillas

# Verificar configuración del sistema
tsx generar.ts verificar

# Estadísticas de Google Sheets
tsx generar.ts estadisticas

# Ayuda general
tsx generar.ts --help
```

## 🎨 Configuración de Plantillas

### Plantillas Word (.docx)
Usa marcadores con doble corchetes:
```
[[nombre]] [[apellido1]] [[apellido2]]
[[fecha_actual]] [[año_actual]]
[[nombre_completo]] [[email]]
```

### Plantillas Excel (.xlsx)
Configura el mapeo en `settings.ts`:
```typescript
export const MAPEOS_EXCEL = {
  mi_plantilla: {
    nombre: "A1",
    apellido1: "B1", 
    apellido2: "C1",
    email: "D1"
  }
};
```

### Variables Automáticas Disponibles
```typescript
// Datos básicos
[[nombre]] [[apellido1]] [[apellido2]]
[[email]] [[telefono]] [[documento]]

// Combinaciones
[[nombre_completo]]        // "Juan Pérez López"
[[apellidos_completos]]    // "Pérez López"
[[iniciales]]              // "J.P.L."

// Fechas
[[fecha_actual]]           // "15/12/2024"
[[fecha_actual_larga]]     // "viernes, 15 de diciembre de 2024"
[[año_actual]]             // 2024
[[mes_actual]]             // "diciembre"

// Utilidades
[[codigo_unico]]           // "JUPE1234"
[[telefono_formateado]]    // "123-456-7890"
[[email_dominio]]          // "gmail.com"
```

## 📁 Estructura de Salida

```
docs_generados/
├── juan_perez/
│   ├── certificado.docx
│   ├── certificado.pdf
│   ├── credenciales.xlsx
│   └── credenciales.pdf
├── maria_gonzalez/
│   ├── certificado.docx
│   ├── certificado.pdf
│   └── ...
└── ...
```

## 🔧 Configuraciones Avanzadas

### Modo Producción (500+ personas)
```typescript
// En settings.ts
export const CONFIG = {
  LIMITE_PERSONAS: 500,
  TIPO_SALIDA: "solo-pdf",        // Más rápido
  LIMPIAR_TEMPORALES: true,       // Ahorra espacio
  OUTPUT_DIR: "./produccion_masiva"
};
```

### Configuración por Lotes
```bash
# Procesar en chunks de 50
tsx generar.ts rango --inicio 1 --fin 50
tsx generar.ts rango --inicio 51 --fin 100
tsx generar.ts rango --inicio 101 --fin 150
```

## 🐛 Solución de Problemas

### Error: "LibreOffice no encontrado"
```bash
# Windows: Verificar instalación
"C:\Program Files\LibreOffice\program\soffice.exe" --version

# Linux
which libreoffice

# Mac
/Applications/LibreOffice.app/Contents/MacOS/soffice --version
```

### Error: "Google Sheets timeout"
- Verificar conexión a internet
- Revisar permisos del archivo JSON
- Confirmar que el SPREADSHEET_ID es correcto

### Error: "Plantilla no encontrada"
```bash
# Verificar rutas
tsx generar.ts plantillas
tsx generar.ts verificar
```

### Memoria insuficiente (500+ personas)
```typescript
// Reducir lote en settings.ts
LIMITE_PERSONAS: 100,  // Procesar de a 100
```

## 📊 Scripts Útiles

```json
{
  "scripts": {
    "dev": "tsx generar.ts",
    "build": "tsc",
    "start": "node dist/generar.js", 
    "interactivo": "tsx generar.ts interactivo",
    "verificar": "tsx generar.ts verificar",
    "stats": "tsx generar.ts estadisticas"
  }
}
```

## 🚀 Casos de Uso Recomendados

### 1. Primera vez / Pruebas
```bash
# Verificar todo funciona
tsx generar.ts verificar

# Probar con pocas personas
tsx generar.ts rango --inicio 1 --fin 3 --tipo ambos
```

### 2. Producción Normal
```bash
# Procesar todos con PDFs solamente (más rápido)
tsx generar.ts todos --tipo solo-pdf --limite 200
```

### 3. Casos Especiales
```bash
# Regenerar documentos de una persona específica
tsx generar.ts persona --nombre "Juan" --tipo ambos

# Solo certificados para un rango
tsx generar.ts rango --inicio 1 --fin 100 --plantillas "certificado"
```

## 📈 Rendimiento Esperado

- **Documentos Word**: ~5-10 por segundo
- **Documentos Excel**: ~3-8 por segundo  
- **Conversión PDF**: ~2-5 por segundo
- **500 personas × 5 plantillas**: ~15-30 minutos

## 🆘 Soporte

Para problemas específicos:
1. Ejecutar `tsx generar.ts verificar`
2. Revisar logs de errores
3. Verificar configuración en `settings.ts`
4. Probar con una sola persona primero

---

¡Listo para generar documentos en masa! 🎉


para iniciar proyecto
pnpm run dev


