# 📄 Generador Masivo de Documentos

Herramienta de línea de comandos que lee datos desde **Google Sheets** y genera automáticamente documentos Word, Excel, PDF y códigos QR en masa, uno por cada persona registrada en la hoja de cálculo.

---

## ¿Qué hace este proyecto?

Imagina que tenés 200 empleados y necesitás generar para cada uno: un certificado de trabajo, una carta de adjudicación y una declaración jurada — todo en PDF y con sus datos correctos. Este proyecto lo hace automáticamente en minutos.

El flujo completo es:

```
Google Sheets (datos) → Plantillas Word/Excel → Documentos personalizados → PDF
                                                                          → QR code
```

---

## Estructura del proyecto

```
doc-por-lotes/
│
├── generar.ts              ← Punto de entrada principal (documentos Word/Excel/PDF)
├── generarQR.ts            ← Script independiente para generar QR codes
│
├── plantillas/             ← Tus archivos .docx y .xlsx con marcadores [[campo]]
│   ├── certificado de trabajo.docx
│   ├── carta adjudicacion.docx
│   └── Declaracion.xlsx
│
├── docs_generados/         ← Salida: documentos generados (se crea automáticamente)
│   └── 1_grupo_apellido1_apellido2_nombre/
│       ├── certificado_trabajo.docx
│       ├── certificado_trabajo.pdf
│       └── ...
│
├── qrs_generados/          ← Salida: QR codes (se crea automáticamente)
│   └── 1_grupo_apellido1_apellido2_nombre/
│       └── qr.png
│
├── src/
│   ├── cli/
│   │   ├── commands.ts     ← Define todos los comandos CLI (todos, rango, persona...)
│   │   └── processor.ts    ← Orquesta el procesamiento persona por persona
│   │
│   ├── config/
│   │   └── settings.ts     ← ⚙️ CONFIGURACIÓN CENTRAL (editar aquí para personalizar)
│   │
│   ├── services/
│   │   ├── googleSheets.ts     ← Conecta y lee datos de Google Sheets
│   │   ├── documentGenerator.ts← Rellena plantillas Word y Excel con los datos
│   │   ├── pdfConverter.ts     ← Convierte archivos a PDF usando LibreOffice
│   │   └── qrGenerator.ts      ← Genera QR codes PNG por persona
│   │
│   ├── types/
│   │   └── index.ts        ← Tipos TypeScript (Persona, PlantillaConfig, etc.)
│   │
│   └── utils/
│       └── fileUtils.ts    ← Utilidades: Logger, limpiarNombre, crear directorios
│
├── generador-docs-31f4b831a196.json  ← 🔑 Credenciales Google (NO subir a git)
├── package.json
└── tsconfig.json
```

---

## Requisitos previos

### 1. Node.js 18 o superior
```bash
node --version   # debe mostrar v18.x.x o mayor
```

### 2. pnpm
```bash
npm install -g pnpm
```

### 3. LibreOffice (solo si vas a generar PDFs)
- **Windows:** Descargar desde [libreoffice.org](https://www.libreoffice.org/)
- **Linux:** `sudo apt install libreoffice`
- **Mac:** `brew install --cask libreoffice`

### 4. Credenciales de Google Sheets
1. Ir a [Google Cloud Console](https://console.cloud.google.com/)
2. Crear un proyecto nuevo
3. Habilitar la **Google Sheets API**
4. Crear una **cuenta de servicio** y descargar el archivo JSON
5. Renombrar el archivo a `generador-docs-31f4b831a196.json` y ponerlo en la raíz del proyecto
6. En tu Google Sheet, compartir la hoja con el email de la cuenta de servicio (aparece en el JSON como `client_email`)

---

## Instalación

```bash
# 1. Clonar el repositorio
git clone <url-del-repo>
cd doc-por-lotes

# 2. Instalar dependencias
pnpm install

# 3. Verificar que todo esté configurado
pnpm run verificar
```

---

## Configuración

Todo se configura en **`src/config/settings.ts`**. Es el único archivo que necesitás tocar para adaptar el proyecto a tus datos.

### ID del Google Sheet

```typescript
SPREADSHEET_ID: "1sMu2QaY2kAy1h-V0YiKhZ2jD6VnEbYaCRXm72Fj_r58",
RANGE: "'Hoja1'!B2:Z",   // columna B fila 2 en adelante
```

Para encontrar el ID: en la URL de tu hoja de cálculo es la parte entre `/d/` y `/edit`.

### Columnas esperadas del Google Sheet

El sistema lee las columnas en este orden a partir de la columna B:

| Columna | Campo | Descripción |
|---------|-------|-------------|
| B | grupo | Grupo o área del empleado |
| C | nombre | Nombre de pila |
| D | apellido1 | Primer apellido |
| E | apellido2 | Segundo apellido |
| F | documento | CI o número de documento |
| G | telefono | Teléfono de contacto |
| H | email | Correo electrónico |
| I | cargo | Cargo o puesto |
| J | fecha_inicio | Fecha de inicio |
| K | fecha_fin | Fecha de fin |

### Ruta de LibreOffice

```typescript
// Windows (por defecto)
SOFFICE_PATH: "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

// Linux / Mac (por defecto si no es Windows)
SOFFICE_PATH: "/usr/bin/libreoffice"
```

### Plantillas

```typescript
PLANTILLAS: [
  {
    nombre: "certificado_trabajo",           // nombre interno (sin espacios)
    archivo: "plantillas/certificado de trabajo.docx",  // ruta al archivo
    tipo: "word",                            // "word" o "excel"
    descripcion: "Certificado de trabajo oficial"
  },
  {
    nombre: "mi_excel",
    archivo: "plantillas/formulario.xlsx",
    tipo: "excel",
    descripcion: "Formulario Excel"
  },
]
```

---

## Cómo hacer las plantillas

### Plantillas Word (.docx)

Usá doble corchete para marcar dónde va cada dato:

```
Yo, [[nombre_completo]], con documento [[documento]], 
trabajo en el cargo de [[cargo]] desde el [[fecha_inicio]].

Firmado el [[fecha_actual_larga]].
```

**Variables disponibles automáticamente:**

| Marcador | Resultado ejemplo |
|----------|-------------------|
| `[[nombre]]` | Juan |
| `[[apellido1]]` | Pérez |
| `[[apellido2]]` | López |
| `[[nombre_completo]]` | Juan Pérez López |
| `[[apellidos_completos]]` | Pérez López |
| `[[iniciales]]` | J.P.L. |
| `[[documento]]` | 12345678 |
| `[[cargo]]` | Técnico |
| `[[telefono]]` | 70123456 |
| `[[telefono_formateado]]` | 701-234-56 |
| `[[email]]` | juan@correo.com |
| `[[email_dominio]]` | correo.com |
| `[[fecha_inicio]]` | 01/01/2024 |
| `[[fecha_fin]]` | 31/12/2024 |
| `[[fecha_actual]]` | 04/03/2026 |
| `[[fecha_actual_larga]]` | miércoles, 4 de marzo de 2026 |
| `[[año_actual]]` | 2026 |
| `[[mes_actual]]` | marzo |
| `[[codigo_unico]]` | JUPE4521 |
| `[[grupo]]` | Área Técnica |

### Plantillas Excel (.xlsx)

Para Excel, los datos se escriben en celdas específicas. Configurás el mapeo en `settings.ts`:

```typescript
export const MAPEOS_EXCEL: Record<string, Record<string, string>> = {
  mi_plantilla: {       // debe coincidir con el "nombre" de la plantilla
    nombre: "C8",       // el campo "nombre" va en la celda C8
    apellido1: "M8",
    apellido2: "U8",
    documento: "H10",
    cargo: "H16",
  }
};
```

Si tu plantilla Excel no tiene un mapeo configurado, el sistema usa un mapeo genérico por defecto.

---

## Uso — Generador de documentos

> **Nota:** Como `tsx` no está en el PATH global, todos los comandos se ejecutan con `pnpm run`.

### Comandos disponibles

```bash
# Modo interactivo (recomendado para empezar)
pnpm run dev

# Ver ayuda
pnpm run tipo
```

Los demás comandos pasan argumentos con `--` después del nombre del script:

```bash
# Procesar TODAS las personas
pnpm run dev -- todos

# Procesar un rango (personas 1 a 50)
pnpm run dev -- rango --inicio 1 --fin 50

# Procesar una persona específica por nombre
pnpm run dev -- persona --nombre "Juan Perez"

# Solo generar PDFs (más rápido, requiere LibreOffice)
pnpm run dev -- todos --tipo solo-pdf

# Solo documentos originales Word/Excel (sin PDF, no requiere LibreOffice)
pnpm run dev -- todos --tipo solo-originals

# Usar plantillas específicas solamente
pnpm run dev -- rango --inicio 1 --fin 10 --plantillas "certificado_trabajo,carta adjudicacion"

# Guardar en una carpeta personalizada
pnpm run dev -- todos --output ./mis_documentos

# Listar personas disponibles en el Sheet
pnpm run dev -- listar --limite 20

# Ver plantillas configuradas
pnpm run dev -- plantillas

# Verificar que todo esté bien configurado
pnpm run verificar

# Ver estadísticas del Sheet
pnpm run dev -- estadisticas
```

### Estructura de salida

```
docs_generados/
└── 1_area_tecnica_perez_lopez_juan_12345678/
    ├── certificado_trabajo.docx
    ├── certificado_trabajo.pdf
    ├── carta adjudicacion.docx
    ├── carta adjudicacion.pdf
    ├── Declaracion Jurada de Imcompatibilidad.xlsx
    └── Declaracion Jurada de Imcompatibilidad.pdf
```

El nombre de la carpeta se arma automáticamente con: `índice_grupo_apellido1_apellido2_nombre_documento`.

---

## Uso — Generador de QR

Script completamente independiente. No requiere LibreOffice ni plantillas.

```bash
# Generar QRs para TODOS
pnpm run qr

# Generar QRs para un rango
pnpm run qr -- --rango 1 5

# Generar QR para una persona
pnpm run qr -- --nombre "Juan Perez"

# Guardar en carpeta personalizada
pnpm run qr -- --rango 1 50 --output ./qrs_evento

# Ver ayuda
pnpm run qr:ayuda
```

### ¿Qué contiene cada QR?

Cada QR codifica texto plano con todos los datos de la persona:

```
GRUPO: Área Técnica
NOMBRE: Juan
APELLIDO 1: Pérez
APELLIDO 2: López
DOCUMENTO: 12345678
CARGO: Técnico Especialista
TELEFONO: 70123456
EMAIL: juan@correo.com
FECHA INICIO: 01/01/2024
FECHA FIN: 31/12/2024
```

Es texto plano para que cualquier app de cámara o lector de QR pueda leerlo sin instalar nada especial.

### Estructura de salida QR

```
qrs_generados/
└── 1_area_tecnica_perez_lopez_juan_12345678/
    └── qr.png   ← imagen PNG 400×400 px, lista para imprimir
```

---

## Scripts de package.json

| Comando | Qué hace |
|---------|----------|
| `pnpm run dev` | Modo interactivo del generador de documentos |
| `pnpm run verificar` | Verifica Google Sheets, LibreOffice y plantillas |
| `pnpm run qr` | Genera QRs para todas las personas |
| `pnpm run qr:todos` | Alias de `qr` |
| `pnpm run qr:ayuda` | Muestra ayuda del generador de QR |
| `pnpm run build` | Compila TypeScript a JavaScript en `/dist` |
| `pnpm run start` | Ejecuta la versión compilada |
| `pnpm run clean` | Elimina la carpeta `/dist` |

---

## Rendimiento estimado

| Tarea | Velocidad aproximada |
|-------|---------------------|
| Generar Word | 5–10 documentos/seg |
| Generar Excel | 3–8 documentos/seg |
| Convertir a PDF | 2–5 por segundo |
| Generar QR | 20–50 por segundo |
| 100 personas × 3 plantillas + PDF | ~5–10 minutos |

---

## Solución de problemas

### `tsx` no reconocido
Usá siempre `pnpm run <script>` en lugar de `tsx` directamente. El `tsx` está instalado localmente, no globalmente.

### Error: Google Sheets — credenciales inválidas
- Verificar que el archivo `.json` de credenciales esté en la raíz del proyecto
- Confirmar que el `SPREADSHEET_ID` en `settings.ts` es correcto
- Asegurarse de que la hoja está compartida con el email `client_email` del JSON

### Error: LibreOffice no encontrado
```bash
# Windows — verificar instalación
"C:\Program Files\LibreOffice\program\soffice.exe" --version

# Linux
which libreoffice

# Mac
/Applications/LibreOffice.app/Contents/MacOS/soffice --version
```
Luego ajustar `SOFFICE_PATH` en `settings.ts`. Si no necesitás PDFs, usá `--tipo solo-originals`.

### Error: Plantilla no encontrada
```bash
pnpm run dev -- plantillas    # muestra las rutas configuradas
pnpm run verificar            # verifica cuáles existen realmente
```
Verificar que las rutas en `settings.ts` coincidan exactamente con los archivos en la carpeta `plantillas/`.

### Marcadores `[[campo]]` no se reemplazan en Word
- Asegurarse de que el marcador esté escrito exactamente con doble corchete
- A veces Word parte el marcador en varios "runs" de texto internamente. Para evitarlo: escribir el marcador completo, seleccionarlo y pegarlo como texto sin formato

### Los datos del Excel quedan en celdas incorrectas
Revisar el mapeo en `MAPEOS_EXCEL` dentro de `settings.ts`. El nombre de la clave debe coincidir exactamente con el campo `nombre` de la plantilla en `CONFIG.PLANTILLAS`.

---

## Dependencias principales

| Paquete | Para qué se usa |
|---------|----------------|
| `googleapis` | Leer datos de Google Sheets |
| `docxtemplater` + `pizzip` | Rellenar plantillas Word con datos |
| `exceljs` | Leer y escribir archivos Excel |
| `qrcode` | Generar imágenes QR en PNG |
| `commander` | Parsear argumentos de línea de comandos |
| `inquirer` | Modo interactivo con preguntas en consola |
| `chalk` | Colores en la consola |
| `ora` | Spinners de carga |
| `cli-progress` | Barra de progreso |
| `tsx` | Ejecutar TypeScript directamente sin compilar |

---

## Archivos que NO se suben a git

El `.gitignore` ya excluye:

```
node_modules/
generador-docs-31f4b831a196.json   ← credenciales privadas
plantillas/                         ← pueden tener info sensible
docs_generados/                     ← salida generada
qrs_generados/                      ← salida generada
dist/
temp/
```

---

## Flujo de desarrollo recomendado

Cuando hagas cambios o pruebes por primera vez:

```bash
# 1. Verificar que todo está bien
pnpm run verificar

# 2. Probar con pocas personas primero
pnpm run dev -- rango --inicio 1 --fin 3 --tipo ambos

# 3. Revisar los archivos generados en docs_generados/

# 4. Si todo está bien, procesar todos
pnpm run dev -- todos --tipo solo-pdf
```
