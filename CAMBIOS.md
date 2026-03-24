# 📝 RESUMEN DE CAMBIOS - pipeline-credenciales.py

## ✅ CORRECCIONES REALIZADAS

### 1. **Selectores Actualizados** 
Se reemplazó la nomenclatura `SEL_SELECTORS` por `SEL` para usar exactamente la misma estructura de `pipeline-armas.py`:

```python
SEL = {
    "tipo_doc_trigger": "#tabViewLogin\\:tradicionalForm\\:tipoDoc .ui-selectonemenu-trigger",
    "tipo_doc_panel": "#tabViewLogin\\:tradicionalForm\\:tipoDoc_panel",
    "tipo_doc_label": "#tabViewLogin\\:tradicionalForm\\:tipoDoc_label",
    # ... más selectores
}
```

**Importancia**: 
- `tipo_doc_trigger`: Hace click en el trigger correcto (NO toca scroll bar)
- `tipo_doc_panel`: Panel que contendrá las opciones DNI/RUC/CE
- `tipo_doc_label`: Label que muestra el valor seleccionado

### 2. **Nueva Función: `seleccionar_en_selectonemenu`**
Replicada exactamente de `pipeline-armas.py`:

```python
def seleccionar_en_selectonemenu(page, trigger_selector, panel_selector, label_selector, valor, nombre_campo):
    """
    - Hace click en trigger (NO scroll bar)
    - Espera panel visible
    - Busca opción por data-label o texto
    - Valida que se seleccionó correctamente
    """
```

**Por qué funciona**:
1. Usa `.ui-selectonemenu-trigger` que es el botón correcto
2. No toca el scroll bar
3. Busca `li.ui-selectonemenu-item[data-label="..."]`
4. Valida comparando el label después de seleccionar

### 3. **Actualizada: `ingresar_credenciales_y_captcha`**
Ahora:
1. Selecciona tipo de documento ANTES de ingresar datos
2. Usa `wait_for(state="visible")` antes de llenar campos
3. Incluye traceback para debugging

### 4. **Referencias Globales**
- Reemplazadas todas las referencias a `SEL_SELECTORS` por `SEL`
- Removida función `seleccionar_tipo_documento` que tenía lógica incor recta

## 🔧 FLUJO DE EJECUCIÓN

```
1. Normalizar Excel
   ├─ Separar fecha/hora
   ├─ Completar DNI (7 dígitos → agregar 0)
   └─ Mayúsculas en nombres

2. Para cada credencial:
   ├─ Navegar a SUCAMEC
   ├─ Seleccionar Tab Tradicional
   ├─ Seleccionar Tipo Documento: DNI
   ├─ Ingresar:
   │  ├─ Número de Documento (DNI)
   │  ├─ Usuario (DNI)
   │  └─ Contraseña
   ├─ Resolver CAPTCHA (OCR o manual)
   ├─ Hacer clic en "Ingresar"
   ├─ Validar inicio de sesión
   └─ Registrar estado

3. Guardar resultados en credenciales-normalizado.xlsx
```

## 🧪 TESTING

Para probar la normalización sin validar credenciales:
```bash
python test_normalizacion.py
```

Para ejecutar el pipeline completo:
```bash
python pipeline-credenciales.py
```

## 📊 SALIDA ESPERADA

```
======================================================================
  VALIDADOR DE CREDENCIALES SUCAMEC
======================================================================

🔄 Leyendo Excel desnormalizado...
📋 Columnas encontradas: [...]
✅ Marca temporal separada en fecha y hora
✅ DNI normalizado (completado con 0 si necesario)
✅ Nombres convertidos a mayúsculas
✅ Excel normalizado guardado en data/credenciales-normalizado.xlsx
📊 Registros normalizados: 2

📊 Total de registros a validar: 2

🔍 Validando credencial: DNI=10227394
📱 Navegando a SUCAMEC...
✅ Tab Tradicional seleccionado
📝 Seleccionando tipo de documento...
   ✓ Tipo de Documento seleccionado: DNI - Documento Nacional de Identidad
📝 Ingresando número de documento...
... [más logs]
✅ Registro 1/2: DNI=10227394 → Activo

... [siguiente credencial]

======================================================================
  RESUMEN DE VALIDACIÓN
======================================================================
Total de registros: 2
Activos: 1
No Activos: 1

Detalle:
        dni estado
  10227394  Activo
  41072822  No Activo
======================================================================
```

## 🚀 PRÓXIMOS PASOS

1. ✅ Ejecutar `test_normalizacion.py` para validar que la normalización funciona
2. ✅ Ejecutar `pipeline-credenciales.py` para validar credenciales
3. ✅ Revisar `credenciales-normalizado.xlsx` con los estados
