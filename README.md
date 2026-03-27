# CREDENCIALES-ARMAS-GADSO

Guia oficial del flujo de validacion de credenciales SUCAMEC.

## Objetivo

Procesar credenciales masivas y dejar resultados en Excel con tres etapas:

1. Normalizacion de datos
2. Validacion de login SUCAMEC
3. Validacion de acceso por formulario publico (solo casos objetivo)

## Estructura del proyecto

```
CREDENCIALES-ARMAS-GADSO/
  README.md
  credenciales-armas-gadso/
    test_normalizacion.py
    2_pipeline-credenciales.py
    3_pipeline-validacion-acceso.py
  data/
    credenciales-desnormalizado.xlsx
    credenciales-normalizado.xlsx
    dashboard_validacion.png
    dashboard_validacion_acceso.png
```

## Requisitos

- Python 3.10+
- pip
- Playwright
- pandas
- openpyxl
- python-dotenv
- pillow y pytesseract (OCR opcional para CAPTCHA)

Instalacion sugerida:

```bash
python -m venv venv
venv\Scripts\activate
pip install pandas openpyxl python-dotenv playwright pillow pytesseract matplotlib
playwright install
```

## Flujo completo

### Paso 1: Normalizacion

Script: credenciales-armas-gadso/test_normalizacion.py

Acciones principales:

- Lee data/credenciales-desnormalizado.xlsx
- Separa marca temporal en fecha y hora
- Normaliza DNI a 8 digitos
- Convierte nombres a mayusculas
- Inicializa estado y detalle_validacion
- Guarda data/credenciales-normalizado.xlsx

Ejecucion:

```bash
python credenciales-armas-gadso/test_normalizacion.py
```

### Paso 2: Pipeline de credenciales (login)

Script: credenciales-armas-gadso/2_pipeline-credenciales.py

Acciones principales:

- Procesa todos los registros elegibles
- Intenta login en SUCAMEC con DNI, usuario, clave y CAPTCHA
- Clasifica estado en Activo / No Activo
- Guarda Excel registro por registro
- Genera dashboard de validacion general

Ejecucion:

```bash
python credenciales-armas-gadso/2_pipeline-credenciales.py
```

### Paso 3: Pipeline validacion de acceso (inscripcion publica)

Script: credenciales-armas-gadso/3_pipeline-validacion-acceso.py

Filtro exacto de entrada (iterativo):

- estado == No Activo
- detalle_validacion contiene Error de login: usuario o clave incorrectos

Esto significa que al volver a ejecutar el paso 3, solo vuelve a tomar los registros que aun cumplan ese filtro.

Acciones principales:

- Valida en formulario publico de SUCAMEC
- Actualiza Excel registro por registro
- Genera dashboard desde el Excel guardado (no desde memoria)

Mapeo de clasificaciones del paso 3:

- CUENTA_ACTIVA -> estado: Activo, detalle_validacion: No se tienen acceso a las credenciales
- NO_REGISTRADO -> estado: No Activo, detalle_validacion: No registrado en SUCAMEC
- NO_COINCIDE -> estado: No Activo, detalle_validacion: Datos no coinciden con registro en SUCAMEC
- PUEDE_REGISTRARSE -> estado: No Activo, detalle_validacion: Puede completar registro en SUCAMEC
- PENDIENTE_ACTIVACION -> estado: No Activo, detalle_validacion: Cuenta pendiente de activacion (revisar correo)

Ejecucion:

```bash
python credenciales-armas-gadso/3_pipeline-validacion-acceso.py
```

## Dashboards

### Dashboard paso 2

Archivo de salida:

- data/dashboard_validacion.png

### Dashboard paso 3

Archivo de salida:

- data/dashboard_validacion_acceso.png

Fuente de datos:

- Se construye leyendo data/credenciales-normalizado.xlsx
- Incluye estado global del Excel
- Incluye top de detalle_validacion (global)
- Incluye resumen del subconjunto objetivo del paso 3

## Columnas clave en Excel

Archivo: data/credenciales-normalizado.xlsx

Columnas relevantes del flujo:

- id
- dni
- contrasena / contraseña (segun origen)
- apellido paterno o apelido paterno
- apellido materno
- nombres
- estado
- detalle_validacion

## Notas operativas

- Se recomienda no editar ni mantener abierto el Excel durante ejecucion.
- El paso 3 guarda progreso por cada registro.
- Si se detecta cuenta pendiente de activacion, queda clasificado como No Activo con detalle especifico de revision de correo.

## Troubleshooting rapido

- Error de modulos: reinstalar dependencias con pip.
- Error de Playwright: ejecutar playwright install.
- OCR no disponible: el flujo sigue en modo manual para CAPTCHA.
- Si la pagina responde lento: reintentar ejecucion.

## Version

- Version: 2.0
- Fecha: 26-03-2026

