# GUÍA DE INICIO RÁPIDO

## Ejecución en 3 Pasos

### 1️⃣ NORMALIZACIÓN
```bash
python credenciales-armas-gadso/test_normalizacion.py
```
✅ Genera: `data/credenciales-normalizado.xlsx`

### 2️⃣ VALIDACIÓN DE CREDENCIALES (LOGIN)
```bash
python credenciales-armas-gadso/pipeline-credenciales.py
```
✅ Actualiza: `estado`, `detalle_validacion` en Excel

### 3️⃣ VALIDACIÓN DE ACCESO (INSCRIPCIÓN)
```bash
python credenciales-armas-gadso/pipeline-validacion-acceso.py
```
✅ Inspecciona registros "No Activo" para acceso público

---

## Instalación Rápida

```bash
# Crear entorno virtual
python -m venv venv
venv\Scripts\activate  # Windows

# Instalar dependencias
pip install pandas openpyxl python-dotenv playwright pillow pytesseract
playwright install
```

---

## Flujo Visual

```
credenciales-desnormalizado.xlsx
    ↓ (test_normalizacion.py)
credenciales-normalizado.xlsx
    ↓ (pipeline-credenciales.py)
+ estado: Activo/No Activo/Error
+ detalle_validacion
    ↓ (pipeline-validacion-acceso.py)
Reporte de acceso público (inscripción)
```

---

## Configuración por Uso

| Script | Opción | Valor | Efecto |
|--------|--------|-------|--------|
| pipeline-credenciales.py | `GUARDAR_CADA_REGISTRO` | True | Guarda Excel cada registro |
| pipeline-credenciales.py | `HEADLESS_BROWSER` | False | Ver navegador (necesario para CAPTCHA manual) |
| pipeline-validacion-acceso.py | `ESCRIBIR_EXCEL` | False | Solo inspeccionar, no modificar |

---

## Cancelación Anytime
**Presiona Enter en la terminal** → Finaliza flujo de forma segura

---

Consulta **README.md** para documentación completa
