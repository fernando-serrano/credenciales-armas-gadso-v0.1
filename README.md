# 🔐 CREDENCIALES-ARMAS-GADSO
## Flujo Integrado de Validación de Credenciales SUCAMEC

Sistema automatizado de **validación integral** de credenciales que integra tres pipelines en cascada: 

1. **📋 Normalización** → Preparación de datos
2. **🔑 Validación de Login** → Acceso directo SUCAMEC  
3. **📝 Validación de Inscripción** → Acceso público (registro)

---

## 📋 Tabla de Contenidos

- [Descripción General & Flujo](#-descripción-general-del-flujo-completo)
- [Arquitectura del Sistema](#-arquitectura-del-sistema)
- [Requisitos Previos](#-requisitos-previos)
- [Instalación y Configuración](#-instalación-y-configuración)
- [Normalización de Datos (Paso 1)](#-paso-1-normalización-de-datos)
- [Pipeline de Login (Paso 2)](#-paso-2-pipeline-de-validación-de-credenciales-login)
- [Pipeline de Inscripción (Paso 3)](#-paso-3-pipeline-de-validación-de-acceso-inscripción-pública)
- [Estructura de Datos Excel](#-estructura-de-datos-excel)
- [Ejecución Completa Pas del Flujo Completo

Este proyecto **valida credenciales de acceso** a través de un **flujo integrado de tres etapas**:

### El Problema que Resuelve

Necesitas validar masivamente si un conjunto de personas:
1. ✅ Tienen acceso directo al sistema SUCAMEC (login con credenciales)
2. 🔄 Si **no tienen acceso**, puedes determinar si pueden **registrarse públicamente**
3. ⚠️ Conocer diferencias entre "datos invalidos" vs "no registrado" vs "ya existe cuenta"

### Visión General del Flujo

```
       ┌──────────────────────────────┐
       │ credenciales-desnormalizado  │
       │ .xlsx                        │
       │ (Datos crudos del CSV)       │
       └──────────────────┬───────────┘
                          │ Contiene:
                          │ - marca temporal (fecha+hora combinada)
                          │ - dni (puede 7 u 8 dígitos)
                          │ - contraseña
                          │ - apelido paterno / materno / nombres
                          │
                          ▼
       ┌──────────────────────────────┐
       │ PASO 1: NORMALIZACIÓN        │ ← test_normalizacion.py
       │ ✓ Separa fecha/hora          │
       │ ✓ DNI a 8 dígitos            │
       │ ✓ Nombres a mayúsculas       │
       │ ✓ IDs únicos estables        │
       └──────────┬───────────────────┘
                  │
                  ▼
       ┌──────────────────────────────┐
       │ credenciales-normalizado     │ ← input para siguientes pasos
       │ .xlsx (Etapa 1 completa)     │
       │ - Columna "estado" = vacío   │
       │ - Columna "detalle" = vacío  │
       └──────────┬───────────────────┘
                  │
        ┌─────────┴─────────┐
        │                   │
        ▼                   ▼
    ╔─────────────╗     ╔─────────────────╗
    │  PASO 2:    │     │   PASO 3:       │
    │  LOGIN      │     │   INSCRIPCIÓN   │
    │  (TODOS)    │     │ (Si No Activo)  │
    ╚──────┬──────╝     ╚────────┬────────╝
           │         pipeline-validacion-acceso.py
           │         (lee "No Activo" del Paso 2)
           │                    │
           ▼                    ▼
    ┌─────────────┐    ┌──────────────────┐
    │SUCAMEC      │    │ SUCAMEC Público  │
    │ panel.xhtml │    │ inscripcionAcceso│
    │             │    │ (sin login)      │
    │ DNI+Usuario │    │                  │
    │ +Contraseña │    │ DNI+Nombres+     │
    │ +CAPTCHA    │    │ Apellidos        │
    └──────┬──────┘    └────────┬─────────┘
           │                    │
           ▼                    ▼
    ┌─────────────┐    ┌──────────────────┐
    │ Estado      │    │ Estado Inscripción│
    │ "Activo"    │    │ "Puede registrar" │
    │ "No Activo" │    │ "Ya registrado"   │
    │ "Error"     │    │ "No coincide"     │
    └──────┬──────┘    │ "Error"          │
           │           └────────┬─────────┘
           │                    │
           └────────┬───────────┘
                    │
                    ▼
            ┌─────────────────┐
            │ REPORTE FINAL   │
            │ Excel actualizado│
            │ + Dashboard     │
            └─────────────────┘
```

### Flujo de Estados por Registro

```
Para CADA fila del Excel:

REGISTRO
  └─ DNI: 12345678, Contraseña: mipass
     │
     ├─ PASO 2 - LOGIN (SIEMPRE se ejecuta)
     │  │
     │  ├─ Intenta login con credenciales
     │  │
     │  └─ Resultado:
     │     ├─ ✅ "Activo" → Credenciales válidas
     │     │   └─ FINAL: Ya tiene acceso, FIN
     │     │
     │     ├─ ❌ "No Activo" → Credenciales inválidas
     │     │   │
     │     │   └─ PASO 3 - INSCRIPCIÓN (SOLO Si No Activo)
     │     │      │
     │     │      ├─ Intenta validar en formulario público
     │     │      │
     │     │      └─ Resultado:
     │     │         ├─ ✅ "Puede registrarse" → Datos válidos, puede crear cuenta
     │     │         ├─ ⚠️  "Ya existe cuenta" → Datos pertenecen a cuenta existente
     │     │         ├─ ❌ "No coincide" → Datos no hacen match en sistema
     │     │         └─ 🔌 "Error" → Error técnico
     │     │
     │     └─ 🔌 "Error" → Error técnico en login
     │        └─ FINAL: No se pudo validar
     │
     └─ RESUMEN PARA ESTE REGISTRO:
        - ID: 1
        - DNI: 12345678
        - Estado Final: (ej: "Puede registrarse")
        - Detalle: (ej: "Validado en inscripción pública")
        - Tiempo: HH:MM:SS
    │ Credenciales        │ Validación Acceso│
    │ + CAPTCHA          │ (si No Activo)   │
    └────────┬───┘          └────────┬─────────┘
             │                       │
             ├─ Activo               ├─ Puede registrarse
             ├─ No Activo    ──────► ├─ Ya existe cuenta
             ├─ Error CAPTCHA        ├─ Error validación
             └─ Credenciales OK      └─ No puede registrarse
```

---

## � Arquitectura del Sistema

### Estructura de Directorios

```
CREDENCIALES-ARMAS-GADSO/
│
├─ README.md                                      ← ESTE ARCHIVO (guía completa)
├─ QUICKSTART.md                                  ← Inicio rápido (5 minutos)
├─ .gitignore                                     ← Archivos ignorados por Git
│
├─ credenciales-armas-gadso/                      ← Scripts principales
│  ├─ test_normalizacion.py                       ← Paso 1: Normaliza Excel
│  ├─ pipeline-credenciales.py                    ← Paso 2: Valida login SUCAMEC
│  └─ pipeline-validacion-acceso.py               ← Paso 3: Valida inscripción pública
│
├─ data/                                          ← Archivos Excel (entrada/salida)
│  ├─ credenciales-desnormalizado.xlsx            ← INPUT: Tu CSV/Excel original
│  ├─ credenciales-normalizado.xlsx               ← OUTPUT: Estados + resultados
│  └─ dashboard_validacion.png                    ← Gráfica de resultados
│
└─ .env                                           ← Variables de entorno (crear si necesitas)
```

### Dependencias & Flujo de Datos

```
Pipeline 1: NORMALIZACIÓN
   Input:  credenciales-desnormalizado.xlsx
   Output: credenciales-normalizado.xlsx
   ├─ Lee datos crudos
   ├─ Separa marca temporal
   ├─ Normaliza DNI
   ├─ Convierte nombres a mayúsculas
   ├─ Crea IDs únicos
   └─ Inicializa columnas de estado

Pipeline 2: LOGIN (SUCAMEC Panel)
   Input:  credenciales-normalizado.xlsx (estado vacío)
   Output: credenciales-normalizado.xlsx (con estado="Activo"/"No Activo"/"Error")
   ├─ Lee TODOS los registros
   ├─ Para cada registro:
   │  ├─ Abre navegador (Playwright)
   │  ├─ Intenta login
   │  ├─ Resuelve CAPTCHA (automático o manual)
   │  ├─ Clasifica resultado
   │  └─ Guarda en Excel
   └─ Genera dashboard de validación

Pipeline 3: INSCRIPCIÓN (SUCAMEC Público)
   Input:  credenciales-normalizado.xlsx (filtra solo "No Activo")
   Output: Consola + Excel opcional
   ├─ Lee registros con estado="No Activo"
   ├─ Para cada registro:
   │  ├─ Abre formulario público
   │  ├─ Valida datos
   │  ├─ Clasifica resultado inscripción
   │  └─ Reporta en consola
   └─ Genera resumen final
```

---

## 🔧 Requisitos Previos

### Requerimientos Mínimos

| Componente | Versión Mínima | Descripción |
|---|---|---|
| **Python** | 3.10+ | Lenguaje de programación |
| **pip** | 20.0+ | Gestor de paquetes Python |
| **Navegador** | Chrome/Edge | Requerido por Playwright (se descarga automáticamente) |
| **SO** | Windows 10+, macOS 10.15+, Linux | Sistema operativo compatible |

### Dependencias Python del Proyecto

```bash
# Instalar todas de una vez:
pip install -r requirements.txt

# O instalación manual:
pip install pandas>=1.3.0          # Lectura/escritura de Excel
pip install openpyxl>=3.6.0        # Engine para Excel (.xlsx)
pip install python-dotenv>=0.19.0  # Carga de variables de entorno
pip install playwright>=1.40.0     # Automatización web
pip install pillow>=9.0.0          # Procesamiento de imágenes
pip install pytesseract>=0.3.10    # OCR para CAPTCHA (opcional pero recomendado)
```

### Verificar Instalación

```bash
# Verifica que Python esté instalado
python --version  # Debe mostrar Python 3.10+

# Verifica que pip esté instalado
pip --version     # Debe mostrar pip 20.0+

# Después de instalar dependencias, prueba import:
python -c "import pandas, openpyxl, playwright, PIL, pytesseract; print('✅ Todas las dependencias OK')"
```

### Software Adicional (Altamente Recomendado)

#### **Tesseract OCR** (Para resolver CAPTCHA automáticamente)

**¿Por qué es necesario?**
- Sin OCR: Tendrás que resolver CAPTCHAs manualmente en cada validación
- Con OCR: Se resuelven automáticamente (6 intentos inteligentes)

**Instalación por SO:**

**Windows:**
1. Descarga el instalador: https://github.com/UB-Mannheim/tesseract/wiki
2. Ejecuta: `tesseract-ocr-w64-setup-v5.x.x.exe`
3. Durante la instalación, atenta la ruta (ej: `C:\Program Files\Tesseract-OCR`)
4. El script detectará automáticamente si Tesseract está disponible

**macOS:**
```bash
brew install tesseract
# La ruta automáticamente será: /usr/local/bin/tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt install tesseract-ocr
# La ruta automáticamente será: /usr/bin/tesseract
```

**Verificar Instalación:**
```bash
tesseract --version  # Debe mostrar versión (ej: tesseract 5.3.0)
```

---

## 🚀 Instalación y Configuración

### Paso 1: Clonar/Descargar el Proyecto
```bash
cd CREDENCIALES-ARMAS-GADSO
```

### Paso 2: Crear Entorno Virtual (Recomendado)
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Mac/Linux
source venv/bin/activate
```

### Paso 3: Instalar Dependencias
```bash
pip install -r requirements.txt
```

**O instalar manualmente:**
```bash
pip install pandas openpyxl python-dotenv playwright pillow pytesseract
playwright install
```

### Paso 4: Configurar Variables de Entorno
Crear archivo `.env` en la raíz del proyecto:
```env
# (Actualmente no requiere variables, pero deja espacio para futuras credenciales)
```

### Paso 5: Instalar Tesseract OCR (Opcional)
Para resolver CAPTCHAs automáticamente:
- Descargar desde: https://github.com/UB-Mannheim/tesseract/wiki
- Durante la instalación, anota la ruta
- La ruta por defecto se configura en `pipeline-credenciales.py`:
  ```python
  pytesseract.pytesseract.tesseract_cmd = r'C:\Users\fserrano\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
  ```

### Paso 6: Preparar Excel de Entrada
1. Crear carpeta `data/` en la raíz si no existe
2. Colocar `credenciales-desnormalizado.xlsx` con estructura:
   ```
   | marca temporal    | dni | contraseña | apelido paterno | apellido materno | nombres |
   | 2024-03-15 10:30  | 1234567 | pass123 | GONZALEZ | RODRIGUEZ | JUAN |
   ```

---

## 🔄 Flujo de Trabajo Completo

### ⚠️ ORDEN CRÍTICO DE EJECUCIÓN

```
PASO 1: Normalización (SIEMPRE PRIMERO)
        └─ Script: test_normalizacion.py o integrado en pipeline-credenciales.py

PASO 2: Validación de Credenciales (Login SUCAMEC)
        └─ Script: pipeline-credenciales.py

PASO 3: Validación de Acceso (Inscripción Pública)
        └─ Script: pipeline-validacion-acceso.py
        └─ Solo procesa registros con estado "No Activo" del Paso 2
```

**⚠️ IMPORTANTE:**
- NO saltar pasos
- Cada paso depende de la salida del anterior
- La normalización es obligatoria

---

## 📊 Componentes Detallados

### 1. NORMALIZACIÓN DE DATOS

**Archivo:** `test_normalizacion.py`

**Propósito:** Prepara el Excel desnormalizado para procesamiento

**Operaciones:**
1. Lee `credenciales-desnormalizado.xlsx`
2. Separa columna "marca temporal" en:
   - `fecha` (YYYY-MM-DD)
   - `hora` (HH:MM:SS)
3. Normaliza DNI:
   - Si tiene 7 dígitos → Agrega 0 adelante
   - Ejemplo: `1234567` → `01234567`
4. Convierte nombres a MAYÚSCULAS
   - `apelido paterno`
   - `apellido materno`
   - `nombres`
5. Crea columna `nombre_completo` unificada
6. Inicializa columnas de estado:
   - `estado` (vacío inicialmente)
   - `detalle_validacion` (vacío inicialmente)
7. Genera ID único estable para cada registro
8. Guarda resultado en `credenciales-normalizado.xlsx`

**Ejecución:**
```bash
python credenciales-armas-gadso/test_normalizacion.py
```

**Salida Interior de Datos:**
```
Excel Normalizado:
┌────┬──────────┬────────────┬────────────┬────────────┬─────────────┬────┬───────────────────┐
│ id │ dni      │ contraseña │ apelido... │ apellido.. │ nombres     │ 🔄 │ nombre_completo   │
├────┼──────────┼────────────┼────────────┼────────────┼─────────────┼────┼───────────────────┤
│ 1  │ 01234567 │ pass123    │ GONZALEZ   │ RODRIGUEZ  │ JUAN        │    │ GONZALEZ RODRIGUEZ│
└────┴──────────┴────────────┴────────────┴────────────┴─────────────┴────┴───────────────────┘
                                                           ↓
                                    Agrega columnas de estado/detalle
                                    Genera IDs estables
```

---

### 2. PIPELINE DE CREDENCIALES (VALIDACIÓN DE LOGIN)

**Archivo:** `pipeline-credenciales.py`

**Propósito:** Valida credenciales intentando login en SUCAMEC

**Flujo Interno:**
```
┌─ Leer Excel Normalizado
│
├─ Para cada registro:
│  │
│  ├─ Extraer DNI, contraseña, nombres
│  │
│  ├─ Abre navegador Playwright
│  │
│  ├─ Navega a: https://www.sucamec.gob.pe/sel/faces/login.xhtml
│  │
│  ├─ Tab "Tradicional" (DNI + Usuario + Contraseña)
│  │
│  ├─ Ingresa datos:
│  │  ├─ Tipo Doc: DNI
│  │  ├─ Número: DNI
│  │  ├─ Usuario: DNI
│  │  └─ Contraseña: contraseña del Excel
│  │
│  ├─ Resuelve CAPTCHA:
│  │  ├─ Intenta OCR automático (si Tesseract disponible)
│  │  └─ Si falla → Solicita resolución manual (usuario lo resuelve visualmente)
│  │
│  ├─ Envía formulario
│  │
│  ├─ Espera resultado (3000ms):
│  │  ├─ ✅ Exitoso → Estado: "Activo"
│  │  ├─ ❌ Credenciales inválidas → Estado: "No Activo"
│  │  ├─ ⚠️  CAPTCHA inválido → Reintentar
│  │  └─ 🔌 Error conexión → Estado: "Error"
│  │
│  └─ Guarda resultado + detalle en Excel
│
└─ Genera Dashboard de Validación
```

**Selectores PrimeFaces:**
```python
SEL = {
    "tab_tradicional": 'a[href="#tabViewLogin:j_idt33"]',
    "tipo_doc_trigger": "#tabViewLogin\\:tradicionalForm\\:tipoDoc .ui-selectonemenu-trigger",
    "numero_documento": "#tabViewLogin\\:tradicionalForm\\:documento",
    "usuario": "#tabViewLogin\\:tradicionalForm\\:usuario",
    "clave": "#tabViewLogin\\:tradicionalForm\\:clave",
    "captcha_img": "#tabViewLogin\\:tradicionalForm\\:imgCaptcha",
    "captcha_input": "#tabViewLogin\\:tradicionalForm\\:textoCaptcha",
    "ingresar": "#tabViewLogin\\:tradicionalForm\\:ingresar",
}
```

**Configuración Interna:**
```python
URL_LOGIN = "https://www.sucamec.gob.pe/sel/faces/login.xhtml?faces-redirect=true"
GUARDAR_CADA_REGISTRO = True    # Guarda progreso en cada iteración
HEADLESS_BROWSER = False         # False = ver navegador (necesario para CAPTCHA manual)
OCR_AVAILABLE = True/False       # Auto-detectado si Tesseract está instalado
```

**Resolución de CAPTCHA:**
1. **Automática (si OCR disponible):**
   - Descarga imagen CAPTCHA
   - Preprocesa: contraste, brightness, erosión
   - 6 intentos OCR con diferentes PSM (Page Segmentation Modes)
   - Si todas fallan → modo manual

2. **Manual:**
   - Muestra navegador visual
   - Usuario resuelve manualmente
   - Presiona Enter cuando complete

**Estados de Salida:**
- `Activo` → Credenciales válidas, login exitoso
- `No Activo` → Credenciales inválidas, usuario no tiene acceso
- `Error` → Error técnico (conexión, timeout, etc.)

**Ejecución:**
```bash
python credenciales-armas-gadso/pipeline-credenciales.py
```

**Salida:**
- `data/credenciales-normalizado.xlsx` actualizado con:
  - Columna `estado` = Activo / No Activo / Error
  - Columna `detalle_validacion` = Mensaje específico
  - Dashboard gráfico de resultados

---

### 3. PIPELINE DE VALIDACIÓN DE ACCESO (INSCRIPCIÓN PÚBLICA)

**Archivo:** `pipeline-validacion-acceso.py`

**Propósito:** Para usuarios con estado "No Activo", determina si pueden registrarse

**Lógica:**
```
Lee registros con estado "No Activo"
        │
        ├─ Abre formulario público: https://www.sucamec.gob.pe/sel/faces/pub/inscripcionAcceso.xhtml
        │
        ├─ Intenta validar datos:
        │  ├─ Tipo Doc: DNI
        │  ├─ Número: DNI del registro
        │  ├─ Nombres: nombres del registro
        │  ├─ Apellido Paterno: apellido paterno
        │  └─ Apellido Materno: apellido materno
        │
        ├─ Hace click en "Validar"
        │
        ├─ Espera respuesta (interpreta payload AJAX + DOM):
        │  ├─ "Ya existe una cuenta activa" → Ya está registrado
        │  ├─ "Ya existe un turno registrado" → Cita ya existe
        │  ├─ "No coincide" → Datos no coinciden
        │  ├─ "Se ha validado" → Puede registrarse
        │  ├─ "Genero" campo habilitado → Puede completar registro
        │  └─ (vacío) → Datos no encontrados
        │
        └─ Actualiza estado en Excel
```

**Selectores del Formulario Público:**
```python
SEL_INSCRIPCION = {
    "tipo_doc_label": "#formInscAcceso\\:cbTipoDoc_label",
    "tipo_doc_trigger": "#formInscAcceso\\:cbTipoDoc .ui-selectonemenu-trigger",
    "numero_doc": "#formInscAcceso\\:numDoc",
    "nombres": "#formInscAcceso\\:nomb",
    "apellido_paterno": "#formInscAcceso\\:appat",
    "apellido_materno": "#formInscAcceso\\:apmat",
    "btn_validar": "#formInscAcceso\\:btnValidar",
}
```

**Configuración:**
```python
URL_INSCRIPCION = "https://www.sucamec.gob.pe/sel/faces/pub/inscripcionAcceso.xhtml"
EXCEL_NORMALIZADO = os.path.join("data", "credenciales-normalizado.xlsx")
HEADLESS_BROWSER = False
ESCRIBIR_EXCEL = False  # False = solo lectura, True = actualiza Excel
```

**Clasificación de Resultados:**
```
├─ "CUENTA_ACTIVA": "Ya existe una cuenta activa"
├─ "TURNO_EXISTE": "Ya existe un turno registrado"
├─ "NO_COINCIDE": "Datos NO coinciden con el sistema"
├─ "PUEDE_REGISTRARSE": "Puede completar el registro"
├─ "ERROR": "Error técnico en validación"
└─ "NO_ENCONTRADO": "Datos no encontrados"
```

**Ejecución:**
```bash
python credenciales-armas-gadso/pipeline-validacion-acceso.py
```

**Salida:**
```
Consola:
- Reporte de registros procesados
- Resumen de estados encontrados
- Cuentas activas detectadas
- Errores/pestaña cerrada

Excel (opcional con ESCRIBIR_EXCEL=True):
- Columna "estado_inscripcion" con resultados
```

---

## 📑 Estructura de Datos Excel

### Excel Entrada: `credenciales-desnormalizado.xlsx`

**Columnas Requeridas:**
```
┌─────────────────┬──────────────┬─────────────┬──────────────────┬──────────────────┬──────────┐
│ marca temporal  │ dni          │ contraseña  │ apelido paterno  │ apellido materno │ nombres  │
├─────────────────┼──────────────┼─────────────┼──────────────────┼──────────────────┼──────────┤
│ 2024-03-15... │ 1234567      │ pass123     │ gonzalez         │ rodriguez        │ juan     │
│ 2024-03-16... │ 87654321     │ mypass      │ perez            │ garcia           │ maria    │
└─────────────────┴──────────────┴─────────────┴──────────────────┴──────────────────┴──────────┘
```

**Validaciones:**
- `marca temporal` → Convertible a datetime
- `dni` → 7-8 dígitos (si 7, se completa con 0)
- `contraseña` → No vacía
- Nombres → Convertir a mayúsculas

---

### Excel Salida: `credenciales-normalizado.xlsx`

**Columnas Generadas:**
```
┌────┬──────────┬────────────┬────────────┬────────────┬─────────────┬────────┬────────┐
│ id │ dni      │ contraseña │ apelido... │ apellido.. │ nombres     │ estado │ detalle│
├────┼──────────┼────────────┼────────────┼────────────┼─────────────┼────────┼────────┤
│ 1  │ 01234567 │ pass123    │ GONZALEZ   │ RODRIGUEZ  │ JUAN        │Activo  │ OK     │
│ 2  │ 87654321 │ mypass     │ PEREZ      │ GARCIA     │ MARIA       │ No Act...│ Error │
└────┴──────────┴────────────┴────────────┴────────────┴─────────────┴────────┴────────┘

Columnas Adicionales Después:
├─ fecha                    → Convertida de "marca temporal"
├─ hora                     → Convertida de "marca temporal"  
├─ nombre_completo          → Unificado de apelidos + nombres
├─ estado                   → Activo / No Activo / Error
├─ detalle_validacion       → Descripción del resultado
└─ (futuro) estado_inscripcion → Resultado de validación de acceso
```

**Estados Válidos:**
```
"Activo"       → Credenciales válidas
"No Activo"    → Credenciales inválidas
"Error"        → Error técnico (timeout, conexión, etc.)
```

**Ejemplos de Detalle:**
```
✅ Activo
   - "Login exitoso"
   - "Credenciales válidas"

❌ No Activo
   - "Credenciales incorrectas"
   - "Usuario no encontrado"
   - "Contraseña inválida"

⚠️  Error
   - "Timeout en validación de CAPTCHA"
   - "No se pudo conectar al servidor"
   - "Error al procesar CAPTCHA"
   - "Navegador cerrado durante ejecución"
```

---

## 🎬 Ejecución Paso a Paso

### Escenario Completo: Primero a Último

#### **PASO 1: Normalización de Datos**

1. Coloca `credenciales-desnormalizado.xlsx` en carpeta `data/`

2. Abre terminal en la raíz del proyecto:
   ```bash
   cd CREDENCIALES-ARMAS-GADSO
   ```

3. Ejecuta normalización:
   ```bash
   python credenciales-armas-gadso/test_normalizacion.py
   ```

4. Salida esperada:
   ```
   ========================================================================
             PRUEBA DE NORMALIZACIÓN
   ========================================================================
   
   🔄 Leyendo Excel desnormalizado...
   📋 Columnas encontradas: [...lista...]
   📊 Total de registros: 145
   
   ✅ Marca temporal separada en fecha y hora
   ✅ DNI normalizado (completado con 0 si necesario)
   ✅ Nombres convertidos a mayúsculas
   ✅ Excel normalizado guardado en data/credenciales-normalizado.xlsx
   
   📊 Registros procesados: 145
   ========================================================================
   ```

5. Verifica: Debe existir `data/credenciales-normalizado.xlsx`

---

#### **PASO 2: Validación de Credenciales (Login)**

1. Terminal en raíz del proyecto:
   ```bash
   python credenciales-armas-gadso/pipeline-credenciales.py
   ```

2. Se abrirá navegador Playwright automáticamente

3. Flujo esperado:
   ```
   ======================================================================
        VALIDADOR DE CREDENCIALES
   ======================================================================
   Modo navegador: VISIBLE
   
   Leyendo Excel normalizado...
   Excel cargado: 145 registros totales
   
   📊 Procesando registro 1 de 145: DNI 01234567...
      ✓ Tipo documento seleccionado: DNI
      ✓ Número de documento ingresado: 01234567
      ✓ Usuario ingresado: 01234567
      ✓ Contraseña ingresada: ***
      
      🔍 Detectado CAPTCHA...
      
      [OCR automático intenta resolver...]
      ✅ CAPTCHA resuelto automáticamente
      
      ✓ Validando login...
      ✅ Login exitoso - Estado: Activo
   
   ✅ Progreso guardado (registro 1 de 145)
   
   Siguiendo con registro 2 de 145...
   ```

4. **Si OCR falla o Tesseract no instalado:**
   ```
   🔍 Detectado CAPTCHA...
   ⚠️  pytesseract NO está instalado → modo MANUAL
   
   🛑 === RESUELVE EL CAPTCHA MANUALMENTE ===
   Código CAPTCHA visible en la imagen
   Escribe el código en la consola:
   > [usuario escribe manualmente]
   ```

5. **Cancelación en cualquier momento:**
   - Presiona **Enter** en terminal
   - Script finaliza de forma segura

6. **Resultado esperado:**
   - Excel actualizado con columnas:
     - `estado` (Activo / No Activo / Error)
     - `detalle_validacion` (descripción)
   - Dashboard gráfico en consola

---

#### **PASO 3: Validación de Acceso (Inscripción)**

1. Asegúrate que PASO 2 completó exitosamente

2. Terminal en raíz del proyecto:
   ```bash
   python credenciales-armas-gadso/pipeline-validacion-acceso.py
   ```

3. Se abrirá formulario público automáticamente

4. Flujo esperado:
   ```
   ======================================================================
        VALIDADOR DE ACCESO (INSCRIPCION/REGISTRO)
   ======================================================================
   Modo navegador: VISIBLE
   
   Leyendo Excel normalizado...
   Excel cargado: 145 registros totales
   
   Registros No Activos encontrados: 42
   Candidatos para validar acceso: 42
   
   ⏳ ADVERTENCIA: actualizar Excel deshabilitado (ESCRIBIR_EXCEL=False)
   
    1/42 DNI 01234567 (GONZALEZ, RODRIGUEZ, JUAN)...
       ✓ Tipo documento: DNI
       ✓ Número: 01234567
       ✓ Nombres válidos
       
       🔍 Validando...
       ✅ CUENTA_ACTIVA → "Ya existe una cuenta activa"
   
    2/42 DNI 87654321 (PEREZ, GARCIA, MARIA)...
       🔍 Validando...
       ✅ PUEDE_REGISTRARSE → "Se ha validado los datos"
   
   ======================================================================
        RESUMEN DE VALIDACION DE ACCESO
   ======================================================================
   Registros No Activos totales: 42
   Candidatos para validar: 42
   Procesados exitosamente: 42
   Cuentas activas encontradas: 15
   
   ✅ Validación completada exitosamente
   ======================================================================
   ```

5. **Resultado:**
   - Informe por consola (sin modificar Excel por defecto)
   - Si `ESCRIBIR_EXCEL=True` → Actualiza Excel con resultados

---

## 🐛 Troubleshooting

### Problema: "ModuleNotFoundError: No module named 'playwright'"

**Solución:**
```bash
pip install playwright
playwright install
```

---

### Problema: "pytesseract NO está instalado"

**Contexto:** El script intenta resolver CAPTCHAs automáticamente
- Si Tesseract no está instalado: **se solicita resolución manual** (NORMAL)
- Usuario resuelve CAPTCHA visualmente en el navegador

**Si quieres OCR automático:**
1. Descargar: https://github.com/UB-Mannheim/tesseract/wiki
2. Instalar con ruta por defecto
3. O actualizar ruta en `pipeline-credenciales.py`:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Tu\Ruta\tesseract.exe'
   ```

---

### Problema: "Event loop is closed" después de Playwright cerrar

**Causa:** Cerrar browser fuera del contexto `with sync_playwright()`

**Solución:** ✅ Ya implementada en ambos scripts
- Browser se cierra dentro del contexto `with`
- No se ejecutan comandos después de su cierre

---

### Problema: "Target page, context or browser has been closed"

**Causa:** Usuario cierra pestaña del navegador durante validación

**Manejo:**
- Script detecta esta condición: `es_error_pestana_cerrada()`
- Registra como "Error" en Excel
- Continúa con siguiente registro

---

### Problema: Timeout en validación de credenciales

**Posibles causas:**
1. **SUCAMEC muy lento** → Aumentar timeout (por defecto 3000ms)
2. **Conexión intermitente** → Reintentar
3. **Firewall bloqueando** → Revisar conectividad

**Solución:**
```python
# En pipeline-credenciales.py, aumentar timeouts globales:
TIMEOUT_LOGIN = 5000  # ms
```

---

### Problema: Excel no se actualiza después de ejecutar scripts

**Verificar:**
1. ¿PASO 1 completó? → Debe existir `data/credenciales-normalizado.xlsx`
2. ¿El archivo está abierto en Excel?** → Cerrar y reintentar
3. ¿Script ejecutó sin errores?** → Revisar consola para mensajes de error

**Solución:**
```bash
# Elimina Excel anterior si está corrupto
rm data/credenciales-normalizado.xlsx

# Reejecutar normalización
python credenciales-armas-gadso/test_normalizacion.py
```

---

### Problema: No se resuelve CAPTCHA manualmente

**Situación:** Se solicita entrada manual, pero el navegador no muestra CAPTCHA

**Causas:**
1. `HEADLESS_BROWSER = True` → Deshabilita navegador visual
2. Navegador cubierto por otra ventana
3. CAPTCHA no cargó correctamente

**Solución:**
```python
# En pipeline-credenciales.py:
HEADLESS_BROWSER = False  # Debe ser False para ver navegador
```

---

### Problema: "FileNotFoundError: data/credenciales-desnormalizado.xlsx"

**Causa:** El archivo de entrada no existe

**Solución:**
1. Crea carpeta `data/` en raíz:
   ```bash
   mkdir data
   ```
2. Coloca `credenciales-desnormalizado.xlsx` en esa carpeta
3. Verifica nombre exacto: debe ser `credenciales-desnormalizado.xlsx`

---

### Problema: Caracteres extraños en nombres o DNI

**Causa:** Encoding incorrecto del Excel

**Solución:**
```python
# Asegurar que openpyxl está actualizado:
pip install --upgrade openpyxl
```

---

## 📌 Notas Importantes

### ⚠️ Restricciones Técnicas

1. **Playwright solo es headless correctamente si NO necesita entrada manual**
   - Para CAPTCHA manual: `HEADLESS_BROWSER = False` obligatorio
   - Modo visual (no headless) es correcto y previsto

2. **Cancelación por Enter**
   - Presionar Enter en la terminal cancela completamente el flujo
   - Script finaliza de forma segura (cierra navegadores, guarda progreso)

3. **Excel: No abrir mientras se ejecutan scripts**
   - Si está abierto: puede no guardar cambios
   - Resolver: cerrar archivo, ejecutar script, abrir resultado

### ✅ Buenas Prácticas

1. **Orden de ejecución:**
   ```
   1. test_normalizacion.py (SIEMPRE primero)
   2. pipeline-credenciales.py
   3. pipeline-validacion-acceso.py
   ```

2. **Configurable según necesidad:**
   ```python
   # En cada script, ajusta según tu contexto:
   
   # pipeline-credenciales.py
   GUARDAR_CADA_REGISTRO = True    # Guardar progreso en cada iteración
   HEADLESS_BROWSER = False         # False = ver navegador
   
   # pipeline-validacion-acceso.py
   ESCRIBIR_EXCEL = False           # False = solo inspeccionar
   HEADLESS_BROWSER = False         # False = ver navegador
   ```

3. **Monitorear progreso:**
   - Consola imprime cada paso
   - Excel se actualiza en tiempo real (si `GUARDAR_CADA_REGISTRO=True`)
   - Puedes abrir Excel en otra ventana para ver progreso

### 📈 Manejo de Reintentos

- **CAPTCHA fallido:** Reintentos automáticos (6 intentos OCR + manual)
- **Timeout temporal:** Reintentar conexión
- **Credenciales inválidas:** Se registra y continúa (no reintentar)

### 🔒 Privacidad y Seguridad

- Las credenciales se usan solo para validación
- NO se almacenan contraseñas en otra parte
- Excel es local (no se sube a servidor)
- Usar `.env` si en futuro se agregan credenciales de sistema

---

## 🚧 Roadmap Futuro

- [ ] Agregar reintentos configurables
- [ ] Exportar resultados a CSV/JSON
- [ ] Crear interfaz gráfica (Tkinter/PySimpleGUI)
- [ ] Integración con bases de datos
- [ ] Reportes automáticos por email
- [ ] Validación de patrones de DNI
- [ ] Caché de resultados conocidos

---

## 📞 Contacto / Soporte

Para problemas específicos:
1. Revisar **Troubleshooting** arriba
2. Verificar logs de consola
3. Revisar estado del Excel intermediario

---

## 📄 Licencia

Proyecto interno GADSO - Todos los derechos reservados

---

**Versión:** 1.0  
**Última actualización:** 26-03-2026  
**Autor:** Equipo GADSO

