# 📄 Generador Masivo de Documentos

Herramienta de línea de comandos que lee datos desde **Google Sheets** y genera automáticamente documentos Word, Excel, PDF, registros de asistencia y códigos QR en masa, uno por cada persona registrada en la hoja de cálculo.

---

## ¿Qué hace este proyecto?

Imagina que tenés 200 empleados y necesitás generar para cada uno: un certificado de trabajo, una carta de adjudicación y una declaración jurada — todo en PDF y con sus datos correctos. Este proyecto lo hace automáticamente en minutos.

El flujo completo es:

```
Google Sheets (datos) → Plantillas Word/Excel → Documentos personalizados → PDF
                                                                          → QR incrustado en PDF
                      → Hoja "Actividades"    → Registros de asistencia  → PDF por fecha
                                                                          → PDF unificado por grupo
                                                                          → QR standalone
```

---

## Estructura del proyecto

```
doc-por-lotes/
│
├── generar.ts              ← Punto de entrada principal (menú interactivo + CLI)
├── generarQR.ts            ← Script independiente para generar QR codes
├── generarAsistencia.ts    ← Script independiente para registros de asistencia
│
├── plantillas/             ← Tus archivos .docx y .xlsx con marcadores [[campo]]
│   ├── certificado de trabajo.docx
│   ├── carta adjudicacion.docx
│   ├── Declaracion.xlsx
│   └── LISTAS DE ASISTENCIA SIREPRE 2026.xlsx
│
├── docs_generados/         ← Salida: documentos por persona (se crea automáticamente)
│   └── G1_perez_lopez_juan_12345678/
│       ├── certificado_trabajo.docx
│       ├── certificado_trabajo.pdf   ← con QR incrustado si qr: true
│       └── ...
│
├── qrs_generados/          ← Salida: QR codes standalone
│   └── G1_perez_lopez_juan_12345678/
│       └── qr.png
│
├── asistencia_generada/    ← Salida: registros de asistencia por grupo
│   └── grupo_1/
│       ├── 11-03-2026_SIMULACRO.xlsx   ← si tipo incluye excel
│       ├── 11-03-2026_SIMULACRO.pdf    ← si tipo incluye pdf por fecha
│       └── grupo_1.pdf                 ← si tipo incluye merged
│
├── src/
│   ├── cli/
│   │   ├── commands.ts     ← Todos los comandos CLI y menú interactivo
│   │   └── processor.ts    ← Orquesta el procesamiento persona por persona
│   │
│   ├── config/
│   │   └── settings.ts     ← ⚙️ CONFIGURACIÓN CENTRAL
│   │
│   ├── services/
│   │   ├── googleSheets.ts       ← Conecta y lee datos de Google Sheets
│   │   ├── documentGenerator.ts  ← Rellena plantillas Word y Excel
│   │   ├── pdfConverter.ts       ← Convierte archivos a PDF con LibreOffice
│   │   ├── qrGenerator.ts        ← Genera QR codes PNG por persona (standalone)
│   │   ├── qrStamper.ts          ← Incrusta QR en los PDFs generados
│   │   └── grupoExcelGenerator.ts← Genera Excel de asistencia por grupo
│   │
│   ├── types/
│   │   └── index.ts        ← Tipos TypeScript (Persona, PlantillaConfig, etc.)
│   │
│   └── utils/
│       └── fileUtils.ts    ← Logger, limpiarNombre, crear directorios
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
6. En tu Google Sheet, compartir la hoja con el email del campo `client_email` del JSON

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

Todo se configura en **`src/config/settings.ts`**.

### ID del Google Sheet

```typescript
SPREADSHEET_ID: "1sMu2QaY2kAy1h-V0YiKhZ2jD6VnEbYaCRXm72Fj_r58",
RANGE: "'Hoja1'!B2:Z",
```

El ID está en la URL de tu hoja entre `/d/` y `/edit`.

### Columnas del Google Sheet — Hoja1

El sistema lee las columnas en este orden a partir de la columna B:

| Col | Campo        | Descripción                    |
|-----|--------------|--------------------------------|
| B   | grupo        | Número o nombre de grupo       |
| C   | nombre       | Nombre de pila                 |
| D   | apellido1    | Primer apellido                |
| E   | apellido2    | Segundo apellido               |
| F   | documento    | CI o número de documento       |
| G   | telefono     | Teléfono de contacto           |
| H   | email        | Correo electrónico             |
| I   | cargo        | Cargo o puesto                 |
| J   | fecha_inicio | Fecha de inicio                |
| K   | fecha_fin    | Fecha de fin                   |
| L   | tipo         | `URBANO` o `MOVIL` (asistencia)|

### Columnas del Google Sheet — Hoja "Actividades"

Usada por el generador de asistencia:

| Col | Campo     | Descripción                    |
|-----|-----------|--------------------------------|
| A   | Fecha     | Formato DD/MM/YYYY o DD-MM-YYYY|
| B   | Actividad | Nombre de la actividad         |
| C   | Ubicacion | Lugar donde se realiza         |

### Plantillas en settings.ts

```typescript
PLANTILLAS: [
  {
    nombre: "certificado_trabajo",
    archivo: join(projectRoot, "plantillas/certificado de trabajo.docx"),
    tipo: "word",
    descripcion: "Certificado de trabajo oficial",
    qr: true,   // ← este PDF llevará QR incrustado
  },
  {
    nombre: "Declaracion Jurada de Imcompatibilidad",
    archivo: join(projectRoot, "plantillas/Declaracion.xlsx"),
    tipo: "excel",
    descripcion: "Declaracion Jurada",
    qr: true,
  },
  {
    nombre: "Registro de asistencia OPERADORES",
    archivo: join(projectRoot, "plantillas/Registro de asistencia OPERADORES.xlsx"),
    tipo: "excel",
    descripcion: "Registro de asistencia",
    qr: false,  // ← este NO lleva QR
  },
],
```

### Ruta de LibreOffice

```typescript
// Windows (por defecto)
SOFFICE_PATH: "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

// Linux / Mac
SOFFICE_PATH: "/usr/bin/libreoffice"
```

---

## Cómo hacer las plantillas Word

Usá doble corchete para marcar dónde va cada dato:

```
Yo, [[nombre_completo]], con documento [[documento]],
trabajo en el cargo de [[cargo]] desde el [[fecha_inicio]].

Firmado el [[fecha_actual_larga]].
```

**Variables disponibles automáticamente:**

| Marcador                | Resultado ejemplo                        |
|-------------------------|------------------------------------------|
| `[[nombre]]`            | Juan                                     |
| `[[apellido1]]`         | Pérez                                    |
| `[[apellido2]]`         | López                                    |
| `[[nombre_completo]]`   | Juan Pérez López                         |
| `[[apellidos_completos]]` | Pérez López                            |
| `[[iniciales]]`         | J.P.L.                                   |
| `[[documento]]`         | 12345678                                 |
| `[[cargo]]`             | Técnico                                  |
| `[[grupo]]`             | Área Técnica                             |
| `[[telefono]]`          | 70123456                                 |
| `[[telefono_formateado]]` | 701-234-56                             |
| `[[email]]`             | juan@correo.com                          |
| `[[email_dominio]]`     | correo.com                               |
| `[[fecha_inicio]]`      | 01/01/2024                               |
| `[[fecha_fin]]`         | 31/12/2024                               |
| `[[fecha_actual]]`      | 04/03/2026                               |
| `[[fecha_actual_larga]]`| miércoles, 4 de marzo de 2026            |
| `[[año_actual]]`        | 2026                                     |
| `[[mes_actual]]`        | marzo                                    |
| `[[codigo_unico]]`      | JUPE4521                                 |

### Plantillas Excel (.xlsx)

Los datos se escriben en celdas específicas. Configurás el mapeo en `settings.ts`:

```typescript
export const MAPEOS_EXCEL: Record<string, Record<string, string>> = {
  mi_plantilla: {       // debe coincidir con el "nombre" de la plantilla
    nombre:    "C8",
    apellido1: "M8",
    apellido2: "U8",
    documento: "H10",
    cargo:     "H16",
  }
};
```

Si la plantilla no tiene mapeo configurado, el sistema usa un mapeo genérico por defecto.

---

## QR incrustado en PDFs

Cada plantilla con `qr: true` en `settings.ts` recibirá un QR incrustado automáticamente en su PDF generado.

La posición, tamaño y página del QR se configuran en **`src/services/qrStamper.ts`**:

```typescript
export const CONFIG_QR_EN_PDF: ConfigQREnPDF[] = [
  {
    plantilla: '*',               // '*' = aplica a todos, o poner el nombre exacto
    posicion:  'superior-derecha',
    tamanoPt:  170,               // tamaño en puntos PDF (1 cm ≈ 28.35 pt)
    margenPt:  20,                // margen desde el borde
    pagina:    'primera',         // 'primera' | 'ultima' | número
  },
];
```

**Posiciones disponibles:** `superior-derecha`, `superior-izquierda`, `inferior-derecha`, `inferior-izquierda`, o coordenadas manuales `{ x: 450, y: 600 }`.

**Conversión rápida cm → pt:**

| cm | pt  |
|----|-----|
| 2  | 57  |
| 3  | 85  |
| 4  | 113 |
| 5  | 142 |
| 6  | 170 |

---

## Uso — Menú interactivo (recomendado)

```bash
pnpm run dev
```

El menú principal te guía por tres secciones:

```
¿Qué querés hacer?
  📄  Generar documentos (Word/Excel/PDF por persona)
  📋  Generar registros de asistencia
  📱  Generar códigos QR
  🔍  Información y verificación del sistema
```

Cada sección tiene su propio sub-menú con todas las opciones disponibles.

---

## Uso — Generador de documentos (CLI)

```bash
# Ver ayuda
pnpm run tipo

# Procesar TODAS las personas
pnpm run dev -- todos

# Procesar un rango (personas 1 a 50)
pnpm run dev -- rango --inicio 1 --fin 50

# Procesar una persona específica
pnpm run dev -- persona --nombre "Juan Perez"

# Solo generar PDFs
pnpm run dev -- todos --tipo solo-pdf

# Solo documentos originales Word/Excel (sin LibreOffice)
pnpm run dev -- todos --tipo solo-originals

# Plantillas específicas
pnpm run dev -- rango --inicio 1 --fin 10 --plantillas "certificado_trabajo,carta adjudicacion"

# Carpeta de salida personalizada
pnpm run dev -- todos --output ./mis_documentos

# Listar personas
pnpm run dev -- listar --limite 20

# Ver plantillas configuradas
pnpm run dev -- plantillas

# Verificar sistema
pnpm run verificar

# Estadísticas del Sheet
pnpm run dev -- estadisticas
```

### Estructura de salida por persona

```
docs_generados/
└── G1_perez_lopez_juan_12345678/
    ├── certificado_trabajo.docx
    ├── certificado_trabajo.pdf      ← con QR incrustado (si qr: true)
    ├── carta adjudicacion.docx
    ├── carta adjudicacion.pdf
    ├── Declaracion Jurada de Imcompatibilidad.xlsx
    └── Declaracion Jurada de Imcompatibilidad.pdf
```

El nombre de carpeta se arma con: `G{grupo}_{apellido1}_{apellido2}_{nombre}_{documento}`.

---

## Uso — Generador de asistencia

Genera registros de asistencia leyendo personas de **Hoja1** y actividades de **Actividades** del mismo Google Sheet.

Por defecto genera un PDF unificado por grupo con todas las fechas en orden cronológico.

```bash
# Menú interactivo (recomendado)
pnpm run dev
# → seleccionar "Generar registros de asistencia"

# O directamente por CLI:

# Todos los grupos, PDF unificado (comportamiento por defecto)
pnpm run asistencia

# Solo un grupo
pnpm run asistencia -- --grupo 26

# Elegir qué archivos generar con --tipo:
pnpm run asistencia -- --tipo solo-excel      # solo .xlsx por fecha
pnpm run asistencia -- --tipo solo-pdf        # un .pdf por fecha
pnpm run asistencia -- --tipo solo-merged     # un PDF por grupo (todas las fechas) [DEFAULT]
pnpm run asistencia -- --tipo todos           # .xlsx + .pdf por fecha + PDF unificado

# Combinaciones:
pnpm run asistencia -- --grupo 26 --tipo todos
pnpm run asistencia -- --tipo todos --output ./mi_salida

# Atajos de npm scripts:
pnpm run asistencia:excel     # → solo-excel
pnpm run asistencia:pdf       # → solo-pdf
pnpm run asistencia:merged    # → solo-merged
pnpm run asistencia:todos     # → todos
pnpm run asistencia:ayuda     # → muestra ayuda
```

### Modos de --tipo explicados

| Modo          | .xlsx | .pdf por fecha | .pdf unificado |
|---------------|:-----:|:--------------:|:--------------:|
| `solo-excel`  | ✅    | ❌             | ❌             |
| `solo-pdf`    | ❌    | ✅             | ❌             |
| `solo-merged` | ❌    | ❌             | ✅             |
| `todos`       | ✅    | ✅             | ✅             |

### Estructura de salida de asistencia

```
asistencia_generada/
└── grupo_26/
    ├── 11-03-2026_SIMULACRO.xlsx          ← si tipo incluye excel
    ├── 11-03-2026_SIMULACRO.pdf           ← si tipo incluye pdf por fecha
    ├── 25-03-2026_CAPACITACION.xlsx
    ├── 25-03-2026_CAPACITACION.pdf
    └── grupo_26.pdf                       ← si tipo incluye merged
```

La plantilla de asistencia tiene dos bloques en la misma hoja:
- **Bloque superior:** operadores MÓVIL
- **Bloque inferior:** operadores URBANO

El campo `tipo` de cada persona en Google Sheets debe ser `URBANO` o `MOVIL` (acepta variantes con/sin tilde, mayúsculas/minúsculas).

---

## Uso — Generador de QR standalone

Genera un archivo `qr.png` por persona, sin depender de plantillas ni LibreOffice.

```bash
# Menú interactivo (recomendado)
pnpm run dev
# → seleccionar "Generar códigos QR"

# O directamente:

# Todos
pnpm run qr

# Rango
pnpm run qr -- --rango 1 50

# Una persona
pnpm run qr -- --nombre "Juan Perez"

# Carpeta personalizada
pnpm run qr -- --output ./qrs_evento

# Ayuda
pnpm run qr:ayuda
```

### Contenido del QR

Texto plano con todos los datos de la persona:

```
GRUPO: 1
NOMBRE: Juan Pérez López
DOCUMENTO: 12345678
CARGO: Técnico Especialista
CELULAR: 70123456
EMAIL: juan@correo.com
FECHA INICIO: 01/01/2024
FECHA FIN: 31/12/2024
```

### Estructura de salida QR

```
qrs_generados/
└── G1_perez_lopez_juan_12345678/
    └── qr.png   ← PNG 400×400 px, listo para imprimir
```

---

## Scripts de package.json — referencia completa

| Comando                      | Qué hace                                         |
|------------------------------|--------------------------------------------------|
| `pnpm run dev`               | Menú interactivo principal                       |
| `pnpm run verificar`         | Verifica Google Sheets, LibreOffice y plantillas |
| `pnpm run tipo`              | Muestra ayuda del CLI de documentos              |
| `pnpm run qr`                | Genera QRs para todas las personas               |
| `pnpm run qr:todos`          | Alias de `qr`                                    |
| `pnpm run qr:ayuda`          | Muestra ayuda del generador de QR                |
| `pnpm run asistencia`        | Genera asistencia (modo: solo-merged)            |
| `pnpm run asistencia:excel`  | Solo archivos Excel por fecha                    |
| `pnpm run asistencia:pdf`    | Un PDF por fecha (sin conservar Excel)           |
| `pnpm run asistencia:merged` | PDF unificado por grupo (todas las fechas)       |
| `pnpm run asistencia:todos`  | Excel + PDF por fecha + PDF unificado            |
| `pnpm run asistencia:ayuda`  | Muestra ayuda del generador de asistencia        |
| `pnpm run build`             | Compila TypeScript a JavaScript en `/dist`       |
| `pnpm run start`             | Ejecuta la versión compilada                     |
| `pnpm run clean`             | Elimina la carpeta `/dist`                       |

---

## Rendimiento estimado

| Tarea                                  | Velocidad aproximada    |
|----------------------------------------|-------------------------|
| Generar Word                           | 5–10 documentos/seg     |
| Generar Excel                          | 3–8 documentos/seg      |
| Convertir a PDF (LibreOffice)          | 2–5 por segundo         |
| Incrustar QR en PDF                    | 10–20 por segundo       |
| Generar QR standalone                  | 20–50 por segundo       |
| 100 personas × 3 plantillas + PDF + QR | ~5–10 minutos           |
| Asistencia: 10 grupos × 5 fechas       | ~3–5 minutos            |

---

## Flujo de desarrollo recomendado

```bash
# 1. Verificar que todo está bien
pnpm run verificar

# 2. Probar con pocas personas primero
pnpm run dev -- rango --inicio 1 --fin 3 --tipo ambos

# 3. Revisar los archivos generados en docs_generados/

# 4. Si todo está bien, procesar todos
pnpm run dev -- todos --tipo solo-pdf

# Para asistencia:
pnpm run asistencia -- --grupo 26 --tipo todos   # probar con un grupo
pnpm run asistencia -- --tipo solo-merged         # producción
```

---

## Solución de problemas

### `tsx` no reconocido
Usá siempre `pnpm run <script>` en lugar de `tsx` directamente.

### Error: Google Sheets — credenciales inválidas
- Verificar que el archivo `.json` esté en la raíz del proyecto
- Confirmar que el `SPREADSHEET_ID` en `settings.ts` es correcto
- Asegurarse de que la hoja está compartida con el `client_email` del JSON

### Error: LibreOffice no encontrado
```bash
# Windows
"C:\Program Files\LibreOffice\program\soffice.exe" --version

# Linux
which libreoffice

# Mac
/Applications/LibreOffice.app/Contents/MacOS/soffice --version
```
Luego ajustar `SOFFICE_PATH` en `settings.ts`. Si no necesitás PDFs, usá `--tipo solo-excel` o `--tipo solo-originals`.

### Marcadores `[[campo]]` no se reemplazan en Word
- Escribir el marcador completo, seleccionarlo y pegarlo como texto sin formato.
  Word a veces parte el texto internamente en varios "runs" y rompe el marcador.

### Los datos del Excel quedan en celdas incorrectas
Revisar `MAPEOS_EXCEL` en `settings.ts`. El nombre de la clave debe coincidir exactamente con el campo `nombre` de la plantilla.

### Asistencia: personas sin tipo no aparecen en ningún bloque
El campo `tipo` en Google Sheets debe ser exactamente `URBANO` o `MOVIL`. Se aceptan variantes con tilde (MÓVIL, URBANO) y cualquier capitalización, pero si el valor es otro (vacío, `OP`, etc.) la persona aparecerá como advertencia y no se incluirá en ningún bloque.

### El QR se ve mal al escanearlo (caracteres raros)
El proyecto usa encoding NFC + modo byte para garantizar que acentos y ñ se lean correctamente en cualquier app de cámara. Si seguís viendo caracteres extraños, verificá que el campo en Google Sheets esté en UTF-8 sin caracteres especiales inesperados.

### El menú interactivo no muestra la opción de asistencia o QR
Verificá que `commands.ts` sea la versión actualizada (la que incluye `menuAsistencia()` y `menuQR()`). Si actualizaste el archivo, reiniciá el proceso con `pnpm run dev`.

---

## Dependencias principales

| Paquete                  | Para qué se usa                              |
|--------------------------|----------------------------------------------|
| `googleapis`             | Leer datos de Google Sheets                  |
| `docxtemplater` + `pizzip` | Rellenar plantillas Word con datos         |
| `exceljs`                | Leer y escribir archivos Excel               |
| `pdf-lib`                | Mergear PDFs e incrustar QR en PDF           |
| `qrcode`                 | Generar imágenes QR en PNG                   |
| `commander`              | Parsear argumentos de línea de comandos      |
| `inquirer`               | Menú interactivo con preguntas en consola    |
| `chalk`                  | Colores en la consola                        |
| `ora`                    | Spinners de carga                            |
| `cli-progress`           | Barra de progreso                            |
| `tsx`                    | Ejecutar TypeScript directamente             |

---

## Archivos que NO se suben a git

El `.gitignore` excluye:

```
node_modules/
generador-docs-31f4b831a196.json   ← credenciales privadas
plantillas/                         ← pueden tener info sensible
docs_generados/                     ← salida generada
qrs_generados/                      ← salida generada
asistencia_generada/                ← salida generada
dist/
temp/
```