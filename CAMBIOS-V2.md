# Cambios en Pipeline Validación Acceso - v2

## Resumen de Modificaciones

El script `3_pipeline-validacion-acceso.py` ha sido modificado para **insertar/actualizar automáticamente el Excel** con los resultados de la validación.

---

## 🎯 Cambios Principales

### 1. **Filtrado Específico**
**ANTES:** Procesaba TODOS los registros con estado "No Activo"

**AHORA:** Solo procesa registros con:
- Estado: `"No Activo"`
- Detalle: `"Error de login: usuario o clave incorrectos"`

```python
# Filtro en procesar_validacion_acceso():
df_no_activos = df[
    (df["estado"].astype(str).str.strip() == "No Activo") &
    (df["detalle_validacion"].astype(str).str.contains(
        "Error de login: usuario o clave incorrectos", 
        case=False, 
        na=False
    ))
].copy()
```

---

### 2. **Mapeo de Resultados a Excel**

Se agregó diccionario `MAPEO_RESULTADOS` que traduce clasificaciones a estados/detalles:

```python
MAPEO_RESULTADOS = {
    "CUENTA_ACTIVA": {
        "estado": "Activo",
        "detalle": "No se tienen acceso a las credenciales",
    },
    "NO_REGISTRADO": {
        "estado": "No Activo",
        "detalle": "No registrado en SUCAMEC",
    },
    "NO_COINCIDE": {
        "estado": "No Activo",
        "detalle": "Datos no coinciden con registro en SUCAMEC",
    },
    "PUEDE_REGISTRARSE": {
        "estado": "No Activo",
        "detalle": "Puede completar registro en SUCAMEC",
    },
    "TURNO_EXISTE": {
        "estado": "No Activo",
        "detalle": "Ya existe turno/cita registrado",
    },
}
```

---

### 3. **Clasificaciones Retornadas**

Las funciones ahora retornan **(clave_clasificación, mensaje)** en lugar de booleanos:

| Clasificación | Resultado en Sistema | Estado Excel | Detalle Excel |
|---|---|---|---|
| `CUENTA_ACTIVA` | Ya existe una cuenta activa | **Activo** | No se tienen acceso a las credenciales |
| `NO_REGISTRADO` | Formulario habilitado | No Activo | No registrado en SUCAMEC |
| `NO_COINCIDE` | Datos no coinciden | No Activo | Datos no coinciden con registro en SUCAMEC |
| `PUEDE_REGISTRARSE` | Validación exitosa | No Activo | Puede completar registro en SUCAMEC |
| `TURNO_EXISTE` | Turno ya existe | No Activo | Ya existe turno/cita registrado |

**Funciones actualizadas:**
- `clasificar_texto_resultado()` → Retorna `(clave, mensaje)`
- `clasificar_payload_ajax()` → Retorna `(clave, mensaje)`
- `validar_resultado_inscripcion_por_ui()` → Retorna `(clave, mensaje)`
- `validar_acceso_inscripcion()` → Retorna `(clave, mensaje)`

---

### 4. **Escritura en Excel**

**CAMBIO CRÍTICO:**
```python
# ANTES:
ESCRIBIR_EXCEL = False  # Solo inspeccionaba

# AHORA:
ESCRIBIR_EXCEL = True   # ✅ Escribe automáticamente cambios
```

**Comportamiento:**
- Lee registro con estado "No Activo" + error credenciales
- Valida acceso en formulario público
- Obtiene clasificación (CUENTA_ACTIVA, NO_REGISTRADO, etc.)
- Busca en MAPEO_RESULTADOS
- **Actualiza Excel:** `df.at[idx, "estado"]` y `df.at[idx, "detalle_validacion"]`
- Al final: `df.to_excel(EXCEL_NORMALIZADO, index=False)`

**Ejemplo de Actualización:**
```
ANTES:
  estado: No Activo
  detalle_validacion: Error de login: usuario o clave incorrectos

DESPUÉS (si CUENTA_ACTIVA):
  estado: Activo
  detalle_validacion: No se tienen acceso a las credenciales

O DESPUÉS (si NO_REGISTRADO):
  estado: No Activo
  detalle_validacion: No registrado en SUCAMEC
```

---

### 5. **Mejoras Visuales**

Consola ahora muestra de forma clara los resultados:
```
✅ ACTIVO POTENCIAL: No se tienen acceso a las credenciales
ℹ️  No registrado en SUCAMEC
⚠️  Datos no coinciden con registro en SUCAMEC
```

Resumen final actualizado:
```
======================================================================
  RESUMEN DE VALIDACION DE ACCESO
======================================================================
Registros No Activos totales: 42
Candidatos para validar: 42
Procesados exitosamente: 42
Cuentas activas encontradas: 8
Estado final: Flujo completado.

✅ Excel actualizado correctamente: data/credenciales-normalizado.xlsx
   Estados y detalles validación se han inscrito en el archivo
======================================================================
```

---

## 📝 Ejemplos de Uso

### Ejemplo 1: Persona con Cuenta Activa

**Entrada Excel:**
| dni | nombres | estado | detalle_validacion |
|---|---|---|---|
| 12345678 | JUAN | No Activo | Error de login: usuario o clave incorrectos |

**Proceso:**
1. Script abre formulario público
2. Completa: 12345678, JUAN, etc.
3. Sistema detecta: "Ya existe una cuenta activa"
4. Clasificación: `CUENTA_ACTIVA`
5. MAPEO: → Estado: Activo, Detalle: "No se tienen acceso a las credenciales"

**Salida Excel:**
| dni | nombres | estado | detalle_validacion |
|---|---|---|---|
| 12345678 | JUAN | **Activo** | **No se tienen acceso a las credenciales** |

---

### Ejemplo 2: Persona No Registrada

**Entrada Excel:**
| dni | nombres | estado | detalle_validacion |
|---|---|---|---|
| 87654321 | MARIA | No Activo | Error de login: usuario o clave incorrectos |

**Proceso:**
1. Script abre formulario público
2. Completa: 87654321, MARIA, etc.
3. Sistema habilita formulario de registro (significa: no existe en el sistema)
4. Clasificación: `NO_REGISTRADO`
5. MAPEO: → Estado: No Activo, Detalle: "No registrado en SUCAMEC"

**Salida Excel:**
| dni | nombres | estado | detalle_validacion |
|---|---|---|---|
| 87654321 | MARIA | No Activo | **No registrado en SUCAMEC** |

---

### Ejemplo 3: Datos No Coinciden

**Entrada Excel:**
| dni | nombres | estado | detalle_validacion |
|---|---|---|---|
| 11111111 | CARLOS | No Activo | Error de login: usuario o clave incorrectos |

**Proceso:**
1. Script abre formulario público
2. Intenta completar pero nombres/apellidos diferentes en sistema
3. Payload AJAX vacío (sin mensajes visibles) pero datos no validados
4. Clasificación: `NO_COINCIDE`
5. MAPEO: → Estado: No Activo, Detalle: "Datos no coinciden con registro en SUCAMEC"

**Salida Excel:**
| dni | nombres | estado | detalle_validacion |
|---|---|---|---|
| 11111111 | CARLOS | No Activo | **Datos no coinciden con registro en SUCAMEC** |

---

## 🚀 Ejecución

### Comando (igual que antes):
```bash
python credenciales-armas-gadso/3_pipeline-validacion-acceso.py
```

### Diferencia:
- **Antes:** Solo mostraba resultados en consola
- **Ahora:** Actualiza Excel automáticamente + muestra en consola

### Cancelación (igual que antes):
Presiona **Enter** en cualquier momento para detener y guardar progreso

---

## ⚙️ Configuración

Se pueden ajustar estos valores en el script:

```python
# Línea 11-14
HEADLESS_BROWSER = False      # False = ver navegador
ESCRIBIR_EXCEL = True         # True = guardar cambios en Excel (ACTIVADO)

# Diccionario MAPEO_RESULTADOS (líneas 16-32)
# Puedes personalizar los detalles de cada classificación aquí
```

---

## 🔍 Verificación

Después de ejecutar, verifica el Excel:
1. Abre `data/credenciales-normalizado.xlsx`
2. Observa que registros "No Activo" con error credenciales ahora tienen:
   - Nueva columna `estado` (puede cambiar a "Activo" o mantenerse "No Activo")
   - Nueva columna `detalle_validacion` con texto específico del mapeo

**Registros que NO fueron procesados:**
- Otros "No Activo" con detales diferentes (ej: timeout, otro error)
- "Activo" (no se reprosesan)
- Vacíos

---

## 📌 Notas Importantes

1. **Solo filtra No Activo + Error Credenciales:**
   - Si un registro "No Activo" tiene otro detalle (ej: "Timeout"), NO se procesa
   - Si un registro es "Activo", NO se procesa

2. **Guarda progreso automáticamente:**
   - Si el script se interrumpe, el Excel se actualiza con los cambios hasta ese punto
   - Reintenta sin perder cambios previos

3. **Detalles personalizables:**
   - Modifica `MAPEO_RESULTADOS` para cambiar los textos que se escriben en Excel

---

## 🆚 Resumen de Cambios vs Versión Anterior

| Aspecto | v1 | v2 |
|---|---|---|
| Filtrado | Todos "No Activo" | "No Activo" + error credenciales |
| Retorna | Booleano (True/False) | Clasificación (clave) |
| Excel | NO se actualiza | ✅ SÍ se actualiza |
| ESCRIBIR_EXCEL | False (solo lectura) | True (lectura+escritura) |
| Mapeo resultados | Manual | Automático (MAPEO_RESULTADOS) |
| Detalles | Mensaje del sistema | Personalizado + mensaje |

---

**Versión:** 2.0  
**Fecha:** 26-03-2026  
**Estado:** Producción Lista  
