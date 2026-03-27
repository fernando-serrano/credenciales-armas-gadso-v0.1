from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
import os
from datetime import datetime
import threading
import time
import numpy as np
import unicodedata
import re


def quitar_tildes(texto: str) -> str:
    t = str(texto or "")
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def normalizar_nombre(texto: str) -> str:
    t = quitar_tildes(texto)
    t = re.sub(r"\s+", " ", t).strip().upper()
    return t


def limpiar_tildes_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Quita tildes en todas las columnas de texto antes de guardar en Excel."""
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(lambda v: quitar_tildes(v) if pd.notna(v) else v)
    return df


def columna_documento(df: pd.DataFrame) -> str:
    if "nro_documento" in df.columns:
        return "nro_documento"
    if "dni" in df.columns:
        return "dni"
    return "nro_documento"


def tipo_doc_por_nro(nro_documento: str) -> str:
    n = len(str(nro_documento or ""))
    if n == 8:
        return "DNI"
    if n == 9:
        return "CARNET EXTRANJERIA"
    return ""


def deduplicar_por_fecha_cercana(df: pd.DataFrame, col_doc: str = "nro_documento", col_fecha_ref: str = "__fecha_ref") -> pd.DataFrame:
    """Para documentos repetidos, conserva la fila con fecha mas cercana a hoy."""
    if col_doc not in df.columns or col_fecha_ref not in df.columns:
        return df

    df_aux = df.copy()
    doc_norm = df_aux[col_doc].fillna("").astype(str).str.strip()

    sin_doc = df_aux[doc_norm == ""].copy()
    con_doc = df_aux[doc_norm != ""].copy()
    if con_doc.empty:
        return df

    hoy = pd.Timestamp.today().normalize()
    con_doc[col_fecha_ref] = pd.to_datetime(con_doc[col_fecha_ref], errors="coerce")
    con_doc["__sin_fecha"] = con_doc[col_fecha_ref].isna().astype(int)
    con_doc["__dist_dias"] = (con_doc[col_fecha_ref] - hoy).abs().dt.days.fillna(10**9)

    con_doc = con_doc.sort_values(
        by=[col_doc, "__sin_fecha", "__dist_dias", col_fecha_ref],
        ascending=[True, True, True, False],
        kind="stable",
    )
    con_doc = con_doc.drop_duplicates(subset=[col_doc], keep="first")
    con_doc = con_doc.drop(columns=["__sin_fecha", "__dist_dias"], errors="ignore")

    return pd.concat([con_doc, sin_doc], ignore_index=True)


def ordenar_por_fecha_asc(df: pd.DataFrame, col_fecha: str = "fecha") -> pd.DataFrame:
    """Ordena por fecha de la mas lejana a la mas reciente (ascendente)."""
    if col_fecha not in df.columns:
        return df
    df_aux = df.copy()
    fecha_dt = pd.to_datetime(df_aux[col_fecha], format="%d/%m/%y", errors="coerce")
    df_aux["__sin_fecha"] = fecha_dt.isna().astype(int)
    df_aux["__fecha_sort"] = fecha_dt

    sort_cols = ["__sin_fecha", "__fecha_sort"]
    sort_asc = [True, True]
    if "id" in df_aux.columns:
        df_aux["__id_sort"] = pd.to_numeric(df_aux["id"], errors="coerce").fillna(10**9)
        sort_cols.append("__id_sort")
        sort_asc.append(True)

    df_aux = df_aux.sort_values(by=sort_cols, ascending=sort_asc, kind="stable").reset_index(drop=True)
    return df_aux.drop(columns=["__sin_fecha", "__fecha_sort", "__id_sort"], errors="ignore")

# ====================== INTENTO DE IMPORTAR OCR (opcional) ======================
OCR_AVAILABLE = False
try:
    from PIL import Image, ImageFilter, ImageEnhance, ImageOps
    from io import BytesIO
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\fserrano\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
    OCR_AVAILABLE = True
    print("✅ OCR (pytesseract) cargado correctamente")
except ImportError:
    print("⚠️  pytesseract NO está instalado → se usará modo MANUAL (CAPTCHA a mano)")
except Exception as e:
    print(f"⚠️  Error al cargar OCR: {e} → modo MANUAL")

load_dotenv()

# ============================================================
# CONFIGURACIÓN
# ============================================================

URL_LOGIN = "https://www.sucamec.gob.pe/sel/faces/login.xhtml?faces-redirect=true"
EXCEL_DESNORMALIZADO = os.path.join("data", "credenciales-desnormalizado.xlsx")
EXCEL_NORMALIZADO = os.path.join("data", "credenciales-normalizado.xlsx")
GUARDAR_CADA_REGISTRO = True  # True: guarda progreso por cada iteración
HEADLESS_BROWSER = False  # True acelera ejecución, pero desactiva CAPTCHA manual visual

# Cancelacion global por consola (Enter)
CANCEL_EVENT = threading.Event()


def iniciar_listener_cancelacion():
    """Inicia un listener en background para cancelar todo el flujo con Enter."""
    if CANCEL_EVENT.is_set():
        return

    def _esperar_enter_cancelacion():
        try:
            input("\n🛑 Presiona Enter en cualquier momento para cancelar todo el flujo...\n")
            CANCEL_EVENT.set()
            print("\n🛑 Cancelación solicitada. Finalizando de forma segura...")
        except EOFError:
            # Entornos sin stdin interactivo
            return
        except Exception:
            return

    hilo = threading.Thread(target=_esperar_enter_cancelacion, daemon=True)
    hilo.start()


def cancelacion_solicitada() -> bool:
    return CANCEL_EVENT.is_set()


def verificar_cancelacion():
    if cancelacion_solicitada():
        raise KeyboardInterrupt("Cancelado por usuario")

# Selectores - Replicar exactamente de pipeline-armas
SEL = {
    "tab_tradicional": 'a[href="#tabViewLogin:j_idt33"]',
    "tipo_doc_trigger": "#tabViewLogin\\:tradicionalForm\\:tipoDoc .ui-selectonemenu-trigger",
    "tipo_doc_panel": "#tabViewLogin\\:tradicionalForm\\:tipoDoc_panel",
    "tipo_doc_label": "#tabViewLogin\\:tradicionalForm\\:tipoDoc_label",
    "numero_documento": "#tabViewLogin\\:tradicionalForm\\:documento",
    "usuario": "#tabViewLogin\\:tradicionalForm\\:usuario",
    "clave": "#tabViewLogin\\:tradicionalForm\\:clave",
    "captcha_img": "#tabViewLogin\\:tradicionalForm\\:imgCaptcha",
    "captcha_input": "#tabViewLogin\\:tradicionalForm\\:textoCaptcha",
    "boton_refresh": "#tabViewLogin\\:tradicionalForm\\:botonCaptcha",
    "ingresar": "#tabViewLogin\\:tradicionalForm\\:ingresar",
}

# ============================================================
# FUNCIONES OCR PARA RESOLVER CAPTCHA
# ============================================================

def corregir_captcha_ocr(texto_raw: str) -> str:
    """Normaliza y limpia el texto OCR del CAPTCHA"""
    if not texto_raw:
        return ""
    texto = texto_raw.strip().upper().replace(" ", "").replace("\n", "").replace("\r", "")
    texto = ''.join(c for c in texto if c.isalnum())
    return texto[:5]  # Solo 5 caracteres


def validar_captcha_texto(texto: str) -> bool:
    """Valida que el CAPTCHA tenga exactamente 5 caracteres alfanuméricos"""
    if not texto or len(texto) != 5:
        return False
    return texto.isalnum()


def preprocesar_imagen_captcha(img_bytes: bytes, variante: int = 0) -> 'Image':
    """Preprocesa la imagen del CAPTCHA para mejorar OCR"""
    if not OCR_AVAILABLE:
        return None
    
    img = Image.open(BytesIO(img_bytes))
    
    if variante == 0:
        # Aumento simple de contraste
        img = ImageEnhance.Contrast(img).enhance(2.0)
    elif variante == 1:
        # Threshold adaptativo
        img = img.convert('L')
        img = ImageEnhance.Contrast(img).enhance(3.0)
        img = img.point(lambda p: 255 if p > 128 else 0)
    else:  # variante == 2
        # Blur + inversión
        img = img.filter(ImageFilter.GaussianBlur(radius=0.5))
        img = ImageOps.invert(img)
        img = img.point(lambda p: 255 if p > 110 else 0)
        img = ImageEnhance.Sharpness(img).enhance(4.0)
    
    return img


def solve_captcha_ocr(page, contexto: str = "CAPTCHA", max_intentos: int = 6):
    """
    Resuelve CAPTCHA con OCR automático o solicita resolución manual
    """
    if not OCR_AVAILABLE:
        print(f"⚠️  OCR no disponible → {contexto} manual")
        return resolver_captcha_manual(page)
    
    PSM_MODES = [7, 8, 13]
    NUM_VARIANTES = 3
    
    intento = 0
    while intento < max_intentos:
        verificar_cancelacion()
        intento += 1
        try:
            print(f"🔍 Intentando resolver {contexto} (intento {intento}/{max_intentos})...")
            page.wait_for_timeout(200)
            
            # Capturar imagen del CAPTCHA
            img_bytes = page.locator(SEL["captcha_img"]).screenshot(type="png")
            
            mejor_texto = None
            
            # Probar diferentes variantes y PSM
            for variante in range(NUM_VARIANTES):
                verificar_cancelacion()
                img = preprocesar_imagen_captcha(img_bytes, variante=variante)
                
                for psm in PSM_MODES:
                    verificar_cancelacion()
                    config = f'--psm {psm} --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ --dpi 300'
                    texto_raw = pytesseract.image_to_string(img, config=config, lang='eng').strip()
                    texto = corregir_captcha_ocr(texto_raw)
                    
                    if validar_captcha_texto(texto):
                        print(f"   ✓ Variante {variante}, PSM {psm}: '{texto}' válido")
                        mejor_texto = texto
                        break
                    else:
                        print(f"   ✗ Variante {variante}, PSM {psm}: '{texto}' (len={len(texto)})")
                
                if mejor_texto:
                    break
            
            if mejor_texto:
                # Ingresar CAPTCHA
                campo_captcha = page.locator(SEL["captcha_input"])
                campo_captcha.fill(mejor_texto)
                page.wait_for_timeout(300)
                print(f"   ✅ CAPTCHA ingresado: {mejor_texto}")
                return mejor_texto
            
            # Si falló, refrescar CAPTCHA
            if intento < max_intentos:
                print(f"   ↻ Refrescando {contexto}...")
                page.locator(SEL["boton_refresh"]).click(force=True)
                page.wait_for_timeout(500)
        
        except KeyboardInterrupt:
            print("🛑 OCR cancelado por usuario")
            return None
        except Exception as e:
            print(f"   ❌ Error en intento {intento}: {e}")
            if intento < max_intentos:
                page.locator(SEL["boton_refresh"]).click(force=True)
                page.wait_for_timeout(500)
    
    print(f"❌ No se pudo resolver {contexto} automáticamente → requerirá resolución manual")
    return resolver_captcha_manual(page)

# ============================================================
# FUNCIONES DE NORMALIZACIÓN
# ============================================================

def normalizar_excel():
    """
    Normaliza el Excel desnormalizado:
    - Estandariza nro_documento
    - Calcula tipo_doc por longitud (8 DNI, 9 CARNET EXTRANJERIA)
    - Formatea fecha en dd/mm/aa
    - Elimina marca temporal y hora
    - Convierte apellidos y nombre a mayúsculas
    - Unifica nombre completo
    """
    print("🔄 Leyendo Excel desnormalizado...")
    df = pd.read_excel(EXCEL_DESNORMALIZADO, dtype=str)
    
    # Hacer copia para no afectar original
    df_normalizado = df.copy()
    
    print(f"📋 Columnas encontradas: {df_normalizado.columns.tolist()}")
    
    # 1. Estandarizar columna nro_documento
    if "nro_documento" not in df_normalizado.columns:
        if "nro_doc" in df_normalizado.columns:
            df_normalizado["nro_documento"] = df_normalizado["nro_doc"]
        elif "dni" in df_normalizado.columns:
            df_normalizado["nro_documento"] = df_normalizado["dni"]

    columnas_alias_doc = [c for c in ["nro_doc", "dni"] if c in df_normalizado.columns]
    if columnas_alias_doc:
        df_normalizado = df_normalizado.drop(columns=columnas_alias_doc)

    # 2. Normalizar nro_documento y tipo_doc (preservar tipo_doc ya informado)
    if "nro_documento" in df_normalizado.columns:
        df_normalizado["nro_documento"] = df_normalizado["nro_documento"].apply(normalizar_dni)
        if "tipo_doc" not in df_normalizado.columns:
            df_normalizado["tipo_doc"] = ""
        tipo_actual = df_normalizado["tipo_doc"].fillna("").astype(str).str.strip()
        tipo_calculado = df_normalizado["nro_documento"].apply(tipo_doc_por_nro)
        df_normalizado["tipo_doc"] = tipo_actual.where(tipo_actual != "", tipo_calculado)
        print("✅ nro_documento normalizado y tipo_doc calculado")

    # 3. Formatear fecha y eliminar hora/marca temporal
    if "fecha" in df_normalizado.columns:
        fecha_dt = pd.to_datetime(df_normalizado["fecha"], errors="coerce")
    elif "marca temporal" in df_normalizado.columns:
        fecha_dt = pd.to_datetime(df_normalizado["marca temporal"], errors="coerce")
    else:
        fecha_dt = pd.Series([pd.NaT] * len(df_normalizado), index=df_normalizado.index)

    df_normalizado["__fecha_ref"] = fecha_dt
    df_normalizado["fecha"] = fecha_dt.dt.strftime("%d/%m/%y")
    for col in ["hora", "marca temporal"]:
        if col in df_normalizado.columns:
            df_normalizado = df_normalizado.drop(columns=[col])

    # Si hay documentos repetidos, conservar el mas cercano a hoy por fecha.
    total_antes = len(df_normalizado)
    df_normalizado = deduplicar_por_fecha_cercana(df_normalizado)
    total_despues = len(df_normalizado)
    if total_despues < total_antes:
        print(f"✅ Duplicados por documento depurados: {total_antes - total_despues} (criterio: fecha mas cercana a hoy)")

    if "__fecha_ref" in df_normalizado.columns:
        df_normalizado = df_normalizado.drop(columns=["__fecha_ref"])
    
    # 4. Convertir apellidos y nombres a mayúsculas (sin tildes)
    nombre_cols = ['apelido paterno', 'apellido materno', 'nombres']
    for col in nombre_cols:
        if col in df_normalizado.columns:
            df_normalizado[col] = df_normalizado[col].apply(normalizar_nombre)
    
    # 4. Crear columna de nombre completo unificado
    if all(col in df_normalizado.columns for col in nombre_cols):
        df_normalizado['nombre_completo'] = (
            df_normalizado['apelido paterno'].fillna('') + ' ' +
            df_normalizado['apellido materno'].fillna('') + ' ' +
            df_normalizado['nombres'].fillna('')
        ).str.replace(r'\s+', ' ', regex=True).str.strip()
    
    print("✅ Nombres convertidos a mayúsculas")
    
    # 5. Inicializar columnas de resultado
    if 'estado' not in df_normalizado.columns:
        df_normalizado['estado'] = ''
    if 'detalle_validacion' not in df_normalizado.columns:
        df_normalizado['detalle_validacion'] = ''
    
    # Verificar si el Excel normalizado ya existe
    df_existente = None
    if os.path.exists(EXCEL_NORMALIZADO):
        try:
            df_existente = pd.read_excel(EXCEL_NORMALIZADO, dtype=str)
            if "nro_documento" not in df_existente.columns:
                if "nro_doc" in df_existente.columns:
                    df_existente["nro_documento"] = df_existente["nro_doc"]
                elif "dni" in df_existente.columns:
                    df_existente["nro_documento"] = df_existente["dni"]
            if "nro_documento" in df_existente.columns:
                df_existente["nro_documento"] = df_existente["nro_documento"].apply(normalizar_dni)
            print(f"✅ Archivo normalizado existente encontrado con {len(df_existente)} registros")
        except Exception as e:
            print(f"⚠️  No se pudo leer archivo existente: {e}")

    # 6. Asegurar columna ID estable para trazabilidad entre ejecuciones
    if 'id' not in df_normalizado.columns:
        df_normalizado.insert(0, 'id', '')

    if df_existente is not None and 'id' in df_existente.columns and 'nro_documento' in df_existente.columns and 'nro_documento' in df_normalizado.columns:
        mapa_id = (
            df_existente[['nro_documento', 'id']]
            .dropna(subset=['nro_documento'])
            .drop_duplicates(subset=['nro_documento'], keep='first')
            .set_index('nro_documento')['id']
        )
        if not mapa_id.empty:
            df_normalizado['id'] = df_normalizado.apply(
                lambda r: mapa_id.get(r['nro_documento'], r['id']), axis=1
            )

    ids_ocupados = pd.to_numeric(df_normalizado['id'], errors='coerce').dropna().astype(int)
    siguiente_id = int(ids_ocupados.max()) + 1 if len(ids_ocupados) > 0 else 1
    for idx in df_normalizado.index:
        if es_valor_vacio(df_normalizado.at[idx, 'id']):
            df_normalizado.at[idx, 'id'] = siguiente_id
            siguiente_id += 1
    
    # PRESERVACIÓN: Si existe archivo anterior, copiar estado/detalle/tipo_doc válidos
    if df_existente is not None:
        print(f"\n🔄 PRESERVANDO datos existentes...")
        mapa_id = None
        if 'id' in df_existente.columns:
            mapa_id = df_existente.drop_duplicates(subset=['id'], keep='first').set_index('id')

        for idx, row in df_normalizado.iterrows():
            nro_doc_actual = normalizar_dni(row.get('nro_documento', ''))
            # Buscar este nro_documento en el archivo existente
            filas_existentes = df_existente[df_existente['nro_documento'] == nro_doc_actual]
            if not filas_existentes.empty:
                fila_existente = filas_existentes.iloc[0]
                # Copiar estado y detalle/tipo_doc sin perder validaciones previas.
                estado_prev = str(fila_existente.get('estado', '')).strip() if pd.notna(fila_existente.get('estado')) else ''
                detalle_prev = str(fila_existente.get('detalle_validacion', '')).strip() if pd.notna(fila_existente.get('detalle_validacion')) else ''
                tipo_prev = str(fila_existente.get('tipo_doc', '')).strip() if pd.notna(fila_existente.get('tipo_doc')) else ''

                id_actual = row.get('id', '')
                if mapa_id is not None and str(id_actual) in mapa_id.index.astype(str):
                    fila_id = mapa_id.loc[mapa_id.index.astype(str) == str(id_actual)].iloc[0]
                    estado_prev = str(fila_id.get('estado', '') or '').strip() or estado_prev
                    detalle_prev = str(fila_id.get('detalle_validacion', '') or '').strip() or detalle_prev
                    tipo_prev = str(fila_id.get('tipo_doc', '') or '').strip() or tipo_prev

                if estado_prev:
                    df_normalizado.at[idx, 'estado'] = estado_prev
                detalle_final = detalle_prev
                if detalle_final:
                    df_normalizado.at[idx, 'detalle_validacion'] = detalle_final
                if tipo_prev and es_valor_vacio(df_normalizado.at[idx, 'tipo_doc']):
                    df_normalizado.at[idx, 'tipo_doc'] = tipo_prev

                if estado_prev or detalle_final or tipo_prev:
                    print(f"   ✓ Documento {nro_doc_actual}: datos de validación preservados")

        # Evita "pérdida" de filas si el desnormalizado trae menos DNIs que el histórico normalizado.
        try:
            docs_actuales = set(df_normalizado['nro_documento'].apply(normalizar_dni))
            df_existente_aux = df_existente.copy()
            df_existente_aux['__doc_norm__'] = df_existente_aux['nro_documento'].apply(normalizar_dni)
            faltantes = df_existente_aux[~df_existente_aux['__doc_norm__'].isin(docs_actuales)].drop(columns=['__doc_norm__'])

            if not faltantes.empty:
                # Alinea columnas para concatenar sin perder estructura
                for col in df_normalizado.columns:
                    if col not in faltantes.columns:
                        faltantes[col] = ''
                for col in faltantes.columns:
                    if col not in df_normalizado.columns:
                        df_normalizado[col] = ''

                faltantes = faltantes[df_normalizado.columns]
                df_normalizado = pd.concat([df_normalizado, faltantes], ignore_index=True)
                print(f"   ✓ Se conservaron {len(faltantes)} registros históricos no presentes en el desnormalizado actual")
        except Exception as e:
            print(f"⚠️  No se pudo conservar registros históricos faltantes: {e}")

    if 'id' in df_normalizado.columns:
        df_normalizado['id'] = pd.to_numeric(df_normalizado['id'], errors='coerce').fillna(0).astype(int)

    df_normalizado = ordenar_por_fecha_asc(df_normalizado)
    
    # Guardar Excel normalizado sin tildes en columnas de texto
    df_normalizado = limpiar_tildes_dataframe(df_normalizado)
    df_normalizado.to_excel(EXCEL_NORMALIZADO, index=False)
    print(f"\n✅ Excel normalizado guardado en {EXCEL_NORMALIZADO}")
    print(f"📊 Registros totales: {len(df_normalizado)}\n")
    
    return df_normalizado


# ============================================================
# FUNCIONES DE VALIDACIÓN DE CREDENCIALES
# ============================================================

def seleccionar_en_selectonemenu(page, trigger_selector: str, panel_selector: str, label_selector: str, valor: str, nombre_campo: str):
    """
    Selecciona una opción PrimeFaces SelectOneMenu por data-label o texto visible.
    Idéntico a pipeline-armas.py - NO TOCA EL SCROLL BAR
    """
    trigger = page.locator(trigger_selector)
    trigger.wait_for(state="visible", timeout=12000)
    trigger.click()

    panel = page.locator(panel_selector)
    panel.wait_for(state="visible", timeout=7000)

    opcion = panel.locator(f'li.ui-selectonemenu-item[data-label="{valor}"]')
    try:
        opcion.wait_for(state="visible", timeout=2000)
    except PlaywrightTimeoutError:
        opcion = panel.locator("li.ui-selectonemenu-item").filter(has_text=valor)
        opcion.wait_for(state="visible", timeout=5000)

    opcion.first.click()
    page.wait_for_timeout(250)

    texto_label = page.locator(label_selector).inner_text().strip()
    if texto_label.upper() != valor.upper():
        raise Exception(
            f"No se confirmó la selección de {nombre_campo}. Esperado: '{valor}' | Actual: '{texto_label}'"
        )
    print(f"   ✓ {nombre_campo} seleccionado: {texto_label}")


def ingresar_credenciales_y_captcha(page, nro_documento: str, contrasena: str, tipo_doc: str = "DNI") -> bool:
    """
    Ingresa credenciales y resuelve CAPTCHA
    """
    try:
        verificar_cancelacion()
        # 1. Seleccionar tipo de documento segun tipo_doc
        print(f"📝 Seleccionando tipo de documento...")
        tipo_doc_norm = str(tipo_doc or "DNI").strip().upper()
        valor_tipo_doc = "CARNET" if "CARNET" in tipo_doc_norm else "DNI - Documento Nacional de Identidad"
        seleccionar_en_selectonemenu(
            page,
            trigger_selector=SEL["tipo_doc_trigger"],
            panel_selector=SEL["tipo_doc_panel"],
            label_selector=SEL["tipo_doc_label"],
            valor=valor_tipo_doc,
            nombre_campo="Tipo de Documento"
        )
        
        # 2. Número de Documento (es el DNI)
        verificar_cancelacion()
        print(f"📝 Ingresando número de documento...")
        campo_numero = page.locator(SEL["numero_documento"])
        campo_numero.wait_for(state="visible", timeout=5000)
        campo_numero.fill(nro_documento)
        page.wait_for_timeout(300)
        
        # 3. Usuario (también es el DNI)
        verificar_cancelacion()
        print(f"📝 Ingresando usuario...")
        campo_usuario = page.locator(SEL["usuario"])
        campo_usuario.wait_for(state="visible", timeout=5000)
        campo_usuario.fill(nro_documento)
        page.wait_for_timeout(300)
        
        # 4. Contraseña
        verificar_cancelacion()
        print(f"📝 Ingresando contraseña...")
        campo_clave = page.locator(SEL["clave"])
        campo_clave.wait_for(state="visible", timeout=5000)
        campo_clave.fill(contrasena)
        page.wait_for_timeout(300)
        
        print(f"✅ Credenciales ingresadas para DOC: {nro_documento}")
        
        # 5. Resolver CAPTCHA
        verificar_cancelacion()
        captcha_resuelto = solve_captcha_ocr(page, contexto="CAPTCHA Login")
        if not captcha_resuelto:
            print("❌ No se pudo resolver CAPTCHA")
            return False
        
        # 6. Hacer clic en Ingresar
        verificar_cancelacion()
        print("🔘 Haciendo clic en 'Ingresar'...")
        boton_ingresar = page.locator(SEL["ingresar"])
        boton_ingresar.click()
        page.wait_for_timeout(2000)
        
        return True
        
    except KeyboardInterrupt:
        print("🛑 Cancelado durante ingreso de credenciales")
        return False
    except Exception as e:
        print(f"❌ Error al ingresar credenciales: {e}")
        import traceback
        traceback.print_exc()
        return False


def resolver_captcha_manual(page):
    """
    Solicita resolución manual del CAPTCHA
    Espera a que el usuario lo resuelva antes de continuar
    """
    try:
        print("\n🔐 CAPTCHA MANUAL REQUERIDO")
        print("⏳ Esperando resolución manual del CAPTCHA en el navegador...")
        print("   Resuelve el CAPTCHA y haz clic en 'Ingresar'")
        
        # Esperar a que el formulario se envíe (indicador de que CAPTCHA fue resuelto)
        # Esperar a cambio de URL o elemento que indica éxito/fallo
        page.wait_for_timeout(3000)  # Espera inicial
        
        # Esperar hasta 120 segundos para que el usuario resuelva
        inicio = datetime.now()
        while (datetime.now() - inicio).seconds < 120:
            if cancelacion_solicitada():
                print("🛑 Cancelación detectada en CAPTCHA manual")
                return False
            try:
                # Verificar si la página ha cambiado (indicador de envío)
                page.wait_for_load_state("networkidle", timeout=500)
                print("✅ CAPTCHA resuelto por usuario")
                return True
            except:
                page.wait_for_timeout(500)
        
        print("⏱️  Timeout esperando resolución de CAPTCHA")
        return False
        
    except Exception as e:
        print(f"❌ Error en resolución manual: {e}")
        return False


def validar_resultado_login_por_ui(page, timeout_ms: int = 3000):
    """
    Determina resultado de login por señales de UI (igual que pipeline-armas.py).
    Devuelve: (login_ok: bool, mensaje_error: str|None, tiempo_segundos: float)
    """
    inicio = time.time()

    selectores_exito = [
        "#j_idt11\\:menuPrincipal",
        "#j_idt11\\:j_idt18",  # botón/cabecera de sesión autenticada
        "form#gestionCitasForm",
    ]
    selectores_error = [
        ".ui-messages-error",
        ".ui-message-error",
        ".ui-growl-message-error",
        ".mensajeError",
        "[class*='error']",
        "[class*='Error']",
    ]

    while (time.time() - inicio) * 1000 < timeout_ms:
        verificar_cancelacion()

        for sel in selectores_exito:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    return True, None, (time.time() - inicio)
            except Exception:
                pass

        for sel in selectores_error:
            try:
                loc = page.locator(sel)
                total = min(loc.count(), 3)
                for i in range(total):
                    txt = (loc.nth(i).inner_text() or "").strip()
                    if txt:
                        return False, txt, (time.time() - inicio)
            except Exception:
                pass

        page.wait_for_timeout(120)

    for sel in selectores_exito:
        try:
            if page.locator(sel).count() > 0:
                return True, None, (time.time() - inicio)
        except Exception:
            pass

    mensaje_error = None
    for sel in selectores_error:
        try:
            loc = page.locator(sel)
            total = min(loc.count(), 3)
            for i in range(total):
                txt = (loc.nth(i).inner_text() or "").strip()
                if txt:
                    mensaje_error = txt
                    break
            if mensaje_error:
                break
        except Exception:
            pass

    return False, mensaje_error, (time.time() - inicio)


def captcha_incorrecto_en_pagina(page) -> bool:
    """Detecta si el intento de login falló por código CAPTCHA inválido."""
    try:
        contenido = page.content().lower()
        patrones = [
            "captcha incorrect",
            "captcha invalido",
            "captcha inválido",
            "código de validación incorrect",
            "codigo de validacion incorrect",
            "texto captcha incorrect",
            "ingrese correctamente el código",
            "ingrese correctamente el codigo",
        ]
        return any(p in contenido for p in patrones)
    except Exception:
        return False


def es_error_captcha(texto: str) -> bool:
    t = str(texto or "").lower()
    patrones = [
        "captcha",
        "código de validación",
        "codigo de validacion",
        "texto captcha",
    ]
    return any(p in t for p in patrones)


def obtener_motivo_no_activo(page) -> str:
    """Clasifica el motivo principal cuando no se logra iniciar sesión."""
    try:
        contenido = page.content().lower()
        if any(p in contenido for p in ["usuario o contraseña", "credenciales", "clave incorrect", "datos incorrect"]):
            return "Credenciales incorrectas"
        if any(p in contenido for p in ["código de validación", "codigo de validacion", "captcha incorrect", "captcha inválido", "captcha invalido"]):
            return "CAPTCHA incorrecto"
        if any(p in contenido for p in ["servicio no disponible", "intente más tarde", "error en el sistema", "ha ocurrido un error"]):
            return "Error de plataforma"
        return "No se pudo confirmar inicio de sesión"
    except Exception:
        return "No se pudo leer mensaje de error"


def limpiar_texto_regla(texto: str) -> str:
    """Normaliza texto para comparaciones robustas de reglas."""
    t = str(texto or "").strip().lower()
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.split())


def es_valor_vacio(valor) -> bool:
    """Detecta vacíos reales y placeholders comunes leídos desde Excel."""
    if not pd.notna(valor):
        return True
    t = str(valor).strip().lower()
    return t in {"", "nan", "none", "null", "nat"}


def normalizar_dni(valor) -> str:
    """Normaliza nro_documento preservando ceros a la izquierda cuando aplique."""
    if es_valor_vacio(valor):
        return ""

    texto = str(valor).strip()

    # Si viene como número exportado por Excel (p.ej. "518206.0"),
    # conservar la parte entera para no inventar dígitos.
    if re.fullmatch(r"\d+(\.0+)?", texto):
        solo_digitos = texto.split(".")[0]
    else:
        solo_digitos = re.sub(r"\D", "", texto)

    if not solo_digitos:
        return ""

    # Para 8 dígitos (DNI) completar con ceros a la izquierda; para 9 (CE) conservar longitud.
    if len(solo_digitos) <= 8:
        solo_digitos = solo_digitos.zfill(8)

    return solo_digitos

def debe_reintentar_registro(dni: str, estado: str, detalle: str) -> bool:
    """
    Determina si un registro debe ser reintentado.
    NO reintentar si el detalle marca error de acceso definitivo.
    """
    if es_valor_vacio(estado) or es_valor_vacio(detalle):
        return True  # Si está vacío, reintentar

    estado_norm = limpiar_texto_regla(estado)
    detalle_norm = limpiar_texto_regla(detalle)

    # NO reintentar estos casos (comparación tolerante a tildes/espacios/mayúsculas)
    no_reintentar = [
        "error de login: usuario o clave incorrectos",
        "usuario o clave incorrect",
        "usuario o contrasena incorrect",
    ]

    for patron in no_reintentar:
        if patron in estado_norm or patron in detalle_norm:
            return False

    return True  # Reintentar por defecto


def debe_procesar_registro(estado: str, detalle: str) -> bool:
    """
    Regla operativa de procesamiento:
    - Procesar si estado está vacío.
    - Procesar si estado es "No Activo" y el detalle NO es error de acceso definitivo.
    - No procesar "Activo" ni otros estados cerrados.
    """
    estado_limpio = '' if es_valor_vacio(estado) else str(estado).strip()
    detalle_limpio = str(detalle).strip() if pd.notna(detalle) else ''

    if not estado_limpio:
        return True

    if estado_limpio == "No Activo":
        return debe_reintentar_registro("", estado_limpio, detalle_limpio)

    return False


def obtener_prioridad_registro(estado: str, detalle: str) -> tuple:
    """
    Retorna (nivel_prioridad, descripcion) para ordenar registros.
    Prioridad 1: No Activo con Número de Documento/CAPTCHA
    Prioridad 2: Vacíos
    Prioridad 3: No Activos (otros)
    Prioridad 4: Activos
    """
    estado_norm = limpiar_texto_regla(estado)
    detalle_norm = limpiar_texto_regla(detalle)

    if estado_norm == "no activo" and (
        "numero de documento" in detalle_norm
        or "captcha" in detalle_norm
        or "codigo de validacion" in detalle_norm
    ):
        return (1, "No Activo Prioritario")

    if es_valor_vacio(estado):
        return (2, "Vacío")
    
    if "CAPTCHA incorrecto (reintentos agotados)" in estado or "CAPTCHA incorrecto (reintentos agotados)" in detalle:
        return (1, "No Activo Prioritario")
    
    if estado == "No Activo":
        return (3, "No Activo")
    
    if estado == "Activo":
        return (4, "Activo")
    
    return (5, "Otro")


def guardar_progreso_excel(df_normalizado, idx_registro: int):
    """Persistencia incremental para no perder avance durante la ejecución."""
    try:
        df_normalizado = limpiar_tildes_dataframe(df_normalizado)
        df_normalizado.to_excel(EXCEL_NORMALIZADO, index=False)
        print(f"   💾 Progreso guardado tras registro {idx_registro + 1}")
    except Exception as e:
        print(f"   ⚠️ No se pudo guardar progreso incremental: {e}")


def formatear_duracion(segundos: float) -> str:
    """Formatea una duración en HH:MM:SS."""
    total = int(max(0, round(segundos)))
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def generar_dashboard_validacion(df_normalizado):
    """
    Genera un dashboard con gráficos de pie y barras.
    Retorna ruta de la imagen guardada.
    """
    try:
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        from matplotlib.patches import Rectangle
        
        # Configurar matplotlib para usar backend sin GUI
        plt.switch_backend('Agg')
        
        # Crear figura con subplots
        fig = plt.figure(figsize=(16, 10))
        fig.suptitle('Dashboard de Validación de Credenciales SUCAMEC', fontsize=18, fontweight='bold', y=0.98)
        
        # ===== GRÁFICO 1: PIE CHART DE ESTADOS =====
        ax1 = plt.subplot(2, 2, 1)
        
        # Contar estados
        conteo_estados = df_normalizado['estado'].value_counts()
        if '' in conteo_estados.index:
            conteo_estados = conteo_estados.rename({'': 'Vacío'})
        
        colores_pie = {
            'Activo': '#2ecc71',
            'No Activo': '#e74c3c',
            'Vacío': '#95a5a6',
            'CAPTCHA incorrecto (reintentos agotados)': '#f39c12'
        }
        colores = [colores_pie.get(str(estado), '#3498db') for estado in conteo_estados.index]
        
        wedges, texts, autotexts = ax1.pie(conteo_estados.values, 
                                            labels=conteo_estados.index, 
                                            autopct='%1.1f%%',
                                            colors=colores,
                                            startangle=90,
                                            textprops={'fontsize': 11, 'weight': 'bold'})
        
        ax1.set_title('Distribución de Estados', fontsize=14, fontweight='bold', pad=15)
        
        # Mejorar formato de porcentajes
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(10)
            autotext.set_weight('bold')
        
        # ===== GRÁFICO 2: BARRAS DE DETALLES =====
        ax2 = plt.subplot(2, 2, 2)
        
        # Contar detalles (top 10)
        conteo_detalles = df_normalizado['detalle_validacion'].value_counts().head(10)
        
        # Crear colores gradientes
        colores_barras = plt.cm.viridis(np.linspace(0, 1, len(conteo_detalles)))
        
        barras = ax2.barh(range(len(conteo_detalles)), conteo_detalles.values, color=colores_barras)
        ax2.set_yticks(range(len(conteo_detalles)))
        ax2.set_yticklabels([str(label)[:50] for label in conteo_detalles.index], fontsize=10)
        ax2.set_xlabel('Cantidad', fontsize=11, fontweight='bold')
        ax2.set_title('Top 10 Motivos de No Login', fontsize=14, fontweight='bold', pad=15)
        ax2.invert_yaxis()
        
        # Agregar valor en barras
        for i, barra in enumerate(barras):
            ancho = barra.get_width()
            ax2.text(ancho, barra.get_y() + barra.get_height()/2, 
                    f' {int(ancho)}', ha='left', va='center', fontweight='bold', fontsize=9)
        
        ax2.grid(axis='x', alpha=0.3, linestyle='--')
        
        # ===== GRÁFICO 3: MÉTRICAS RESUMIDAS =====
        ax3 = plt.subplot(2, 2, 3)
        ax3.axis('off')
        
        # Calcular estadísticas
        total = len(df_normalizado)
        activos = len(df_normalizado[df_normalizado['estado'] == 'Activo'])
        no_activos = len(df_normalizado[df_normalizado['estado'] == 'No Activo'])
        vacios = len(df_normalizado[df_normalizado['estado'].isna() | (df_normalizado['estado'] == '')])
        captcha_agotados = len(df_normalizado[df_normalizado['detalle_validacion'].str.contains('CAPTCHA incorrecto', na=False)])
        error_credenciales = len(df_normalizado[df_normalizado['detalle_validacion'].str.contains('Error de login', na=False)])
        
        # Crear tabla de info
        info_text = f"""RESUMEN GENERAL

Total de registros: {total}

ESTADOS:
  • Activos: {activos} ({100*activos/total:.1f}%)
  • No Activos: {no_activos} ({100*no_activos/total:.1f}%)
  • Vacíos: {vacios} ({100*vacios/total:.1f}%)

DETALLES CRÍTICOS:
  • CAPTCHA Reintentos Agotados: {captcha_agotados}
  • Error de Credenciales: {error_credenciales}"""
        
        ax3.text(0.1, 0.9, info_text, 
                fontsize=12, 
                verticalalignment='top',
                family='monospace',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3),
                fontweight='bold')
        
        # ===== GRÁFICO 4: DISTRIBUCIÓN TEMPORAL (si hay columna hora) =====
        ax4 = plt.subplot(2, 2, 4)
        
        if 'hora' in df_normalizado.columns:
            # Extraer hora
            df_normalizado['hora_procesamiento'] = pd.to_datetime(df_normalizado['hora'], format='%H:%M:%S', errors='coerce').dt.hour
            distribucion_hora = df_normalizado['hora_procesamiento'].value_counts().sort_index()
            
            ax4.bar(distribucion_hora.index, distribucion_hora.values, color='#3498db', edgecolor='black', linewidth=1.5)
            ax4.set_xlabel('Hora del Día', fontsize=11, fontweight='bold')
            ax4.set_ylabel('Cantidad de Registros', fontsize=11, fontweight='bold')
            ax4.set_title('Distribución por Hora de Procesamiento', fontsize=14, fontweight='bold', pad=15)
            ax4.grid(axis='y', alpha=0.3, linestyle='--')
            ax4.set_xticks(range(0, 24, 2))
        else:
            ax4.text(0.5, 0.5, 'No hay datos de hora disponibles', 
                    ha='center', va='center', fontsize=12, fontweight='bold',
                    transform=ax4.transAxes)
            ax4.axis('off')
        
        # Ajustar espaciado
        plt.tight_layout()
        
        # Guardar imagen
        ruta_dashboard = os.path.join("data", "dashboard_validacion.png")
        plt.savefig(ruta_dashboard, dpi=150, bbox_inches='tight')
        print(f"\n✅ Dashboard guardado en: {ruta_dashboard}")
        plt.close()
        
        return ruta_dashboard
        
    except Exception as e:
        print(f"⚠️  Error generando dashboard: {e}")
        return None

def validar_credencial(nro_documento: str, contrasena: str, tipo_doc: str = "DNI", max_reintentos_captcha: int = 3, playwright_instance=None) -> tuple:
    """
    Valida una credencial completa en SUCAMEC.
    Retorna (estado, detalle_validacion)
    """
    print(f"\n🔍 Validando credencial: DOC={nro_documento} ({tipo_doc})")
    
    browser = None
    own_playwright = False
    p = playwright_instance
    try:
        verificar_cancelacion()
        if p is None:
            p = sync_playwright().start()
            own_playwright = True

        browser = p.chromium.launch(headless=HEADLESS_BROWSER)
        for intento_login in range(1, max_reintentos_captcha + 1):
            verificar_cancelacion()
            print(f"🔁 Intento de login {intento_login}/{max_reintentos_captcha}")

            context = browser.new_context()
            page = context.new_page()
            try:
                print("📱 Navegando a SUCAMEC...")
                page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(800)

                try:
                    tab = page.locator(SEL["tab_tradicional"])
                    tab.wait_for(state="visible", timeout=8000)
                    tab.click()
                    page.wait_for_timeout(400)
                    print("✅ Tab Tradicional seleccionado")
                except Exception as e:
                    print(f"ℹ️  Tab Tradicional: {e}")

                if not ingresar_credenciales_y_captcha(page, nro_documento, contrasena, tipo_doc=tipo_doc):
                    return "No Activo", "No se pudo completar ingreso de datos/CAPTCHA"

                print("🔍 Validando inicio de sesión...")
                login_ok, mensaje_error_ui, tiempo_espera = validar_resultado_login_por_ui(page, timeout_ms=3500)

                if login_ok:
                    print("✅ CREDENCIAL ACTIVA")
                    return "Activo", "Inicio de sesión correcto"

                motivo = (mensaje_error_ui or "").strip()
                if not motivo:
                    motivo = obtener_motivo_no_activo(page)

                if es_error_captcha(motivo) or captcha_incorrecto_en_pagina(page):
                    print(f"⚠️ CAPTCHA incorrecto detectado. ({motivo or 'sin mensaje'})")
                    if intento_login < max_reintentos_captcha:
                        print("↻ Reintentando login...")
                        continue
                    print("❌ Se agotaron los reintentos por CAPTCHA incorrecto")
                    return "No Activo", "CAPTCHA incorrecto (reintentos agotados)"

                if not motivo:
                    motivo = f"No se detectó sesión autenticada (validado en {tiempo_espera:.2f}s)"
                print(f"❌ Login fallido: {motivo}")
                return "No Activo", motivo
            finally:
                try:
                    context.close()
                except Exception:
                    pass

        return "No Activo", "No se logró iniciar sesión"
    except KeyboardInterrupt:
        print("🛑 Validación cancelada por usuario")
        return "No Activo", "Cancelado por usuario"
    except Exception as e:
        print(f"❌ Error durante validación: {e}")
        import traceback
        traceback.print_exc()
        return "No Activo", f"Error técnico: {e}"
    finally:
        if browser:
            try:
                browser.close()
            except Exception:
                # Evita falsos negativos por cierre tardío del event loop de Playwright
                pass
        if own_playwright and p is not None:
            try:
                p.stop()
            except Exception:
                pass


# ============================================================
# FUNCIÓN PRINCIPAL - ITERAR REGISTROS
# ============================================================

def procesar_todas_credenciales():
    """
    Lee el Excel normalizado e itera sobre todos los registros
    para validar cada credencial.
    
    PRIORIDAD DE PROCESAMIENTO (dentro de los elegibles):
    1. CAPTCHA incorrecto (reintentos agotados)
    2. Registros con estado vacío
    3. Registros con "No Activo"
    4. Registros con "Activo" (se muestran en prioridad, pero no se reprocesan)
    
    NO REINTENTA:
    - Número de Documento:*
    - Error de login: usuario o clave incorrectos
    """
    print("\n" + "="*70)
    print("  VALIDADOR DE CREDENCIALES SUCAMEC")
    print("="*70)
    iniciar_listener_cancelacion()
    inicio_flujo = time.perf_counter()
    
    # Normalizar Excel primero (preserva datos existentes)
    df_normalizado = normalizar_excel()
    
    # Validar que existan las columnas necesarias
    required_cols = ['nro_documento', 'contraseña']
    for col in required_cols:
        if col not in df_normalizado.columns:
            print(f"❌ ERROR: Columna '{col}' no encontrada en Excel")
            print(f"Columnas disponibles: {df_normalizado.columns.tolist()}")
            return
    
    # Inicializar columnas de resultado si no existen
    if 'estado' not in df_normalizado.columns:
        df_normalizado['estado'] = ''
    if 'detalle_validacion' not in df_normalizado.columns:
        df_normalizado['detalle_validacion'] = ''
    
    # ===== PRIORIZACIÓN DE REGISTROS =====
    print("\n🔄 Priorizando registros para procesamiento...")
    
    # Crear índice de prioridad
    prioridades = []
    for idx, row in df_normalizado.iterrows():
        nivel, desc = obtener_prioridad_registro(row['estado'], row['detalle_validacion'])
        prioridades.append((idx, nivel, desc))
    
    # Ordenar por prioridad
    prioridades.sort(key=lambda x: x[1])
    
    print(f"\n📊 Registros a procesar por prioridad:")
    for nivel in [1, 2, 3, 4, 5]:
        registros_nivel = [p for p in prioridades if p[1] == nivel]
        if registros_nivel:
            print(f"   Prioridad {nivel}: {len(registros_nivel)} registros ({registros_nivel[0][2]})")
    
    # ===== PROCESAMIENTO CON PRIORIDAD =====
    registros_activos = 0
    registros_inactivos = 0
    registros_saltados = 0
    contador_procesados = 0
    total_a_procesar = len([
        p for p in prioridades
        if debe_procesar_registro(
            df_normalizado.loc[p[0], 'estado'],
            df_normalizado.loc[p[0], 'detalle_validacion']
        )
    ])
    
    with sync_playwright() as p:
        for idx, nivel_prioridad, desc_prioridad in prioridades:
            if cancelacion_solicitada():
                print("🛑 Cancelación detectada. Se detiene el procesamiento.")
                break
            
            try:
                verificar_cancelacion()
                row = df_normalizado.loc[idx]
                nro_documento = normalizar_dni(row.get('nro_documento', row.get('nro_doc', row.get('dni', ''))))
                contrasena = str(row['contraseña']).strip()
                estado_actual = '' if es_valor_vacio(row['estado']) else str(row['estado']).strip()
                detalle_actual = '' if es_valor_vacio(row['detalle_validacion']) else str(row['detalle_validacion']).strip()
                
                # Validar que no estén vacíos
                if not nro_documento or not contrasena:
                    print(f"⚠️  Registro {idx+1}: nro_documento o contraseña vacíos")
                    df_normalizado.at[idx, 'estado'] = "No Activo"
                    df_normalizado.at[idx, 'detalle_validacion'] = "nro_documento o contraseña vacíos"
                    registros_inactivos += 1
                    if GUARDAR_CADA_REGISTRO:
                        guardar_progreso_excel(df_normalizado, idx)
                    continue
                
                # DECISIÓN: procesar solo vacíos o "No Activo" reintentable
                if not debe_procesar_registro(estado_actual, detalle_actual):
                    print(f"⏭️  Registro {idx+1}: SALTADO (no elegible para reproceso) - DOC={nro_documento} - Estado: {estado_actual}")
                    registros_saltados += 1
                    if estado_actual == "Activo":
                        registros_activos += 1
                    else:
                        registros_inactivos += 1
                    continue
                
                # PROCESAR: Validar nuevamente
                print(f"\n🔍 Procesando Prioridad {nivel_prioridad}: Registro {idx+1}")
                contador_procesados += 1
                print(f"   [{contador_procesados}/{total_a_procesar}] Validando DOC={nro_documento}...")
                
                tipo_doc = str(row.get('tipo_doc', tipo_doc_por_nro(nro_documento))).strip().upper()
                estado, detalle = validar_credencial(nro_documento, contrasena, tipo_doc=tipo_doc, playwright_instance=p)
                
                # Registrar resultado
                df_normalizado.at[idx, 'estado'] = estado
                df_normalizado.at[idx, 'detalle_validacion'] = detalle
                
                if estado == "Activo":
                    registros_activos += 1
                    print(f"   ✅ RESULTADO: {estado}")
                else:
                    registros_inactivos += 1
                    print(f"   ❌ RESULTADO: {estado}")
                    print(f"      Motivo: {detalle}")

                if GUARDAR_CADA_REGISTRO:
                    guardar_progreso_excel(df_normalizado, idx)
                
            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"❌ Error procesando registro {idx+1}: {e}")
                df_normalizado.at[idx, 'estado'] = "No Activo"
                df_normalizado.at[idx, 'detalle_validacion'] = f"Error en procesamiento: {str(e)[:100]}"
                registros_inactivos += 1
                if GUARDAR_CADA_REGISTRO:
                    guardar_progreso_excel(df_normalizado, idx)
    
    # Guardar Excel con estados actualizados
    df_normalizado.to_excel(EXCEL_NORMALIZADO, index=False)
    print(f"\n✅ Proceso completado. Resultados guardados en {EXCEL_NORMALIZADO}")
    
    # Generar dashboard
    print("\n📊 Generando dashboard...")
    generar_dashboard_validacion(df_normalizado)
    
    # Mostrar resumen final
    print("\n" + "="*70)
    print("  RESUMEN FINAL DE VALIDACIÓN")
    print("="*70)
    total = len(df_normalizado)
    print(f"\nTotal de registros: {total}")
    print(f"  ✅ Activos: {registros_activos} ({100*registros_activos/total:.1f}%)")
    print(f"  ❌ No Activos: {registros_inactivos} ({100*registros_inactivos/total:.1f}%)")
    print(f"  ⏭️  Saltados (ya validados): {registros_saltados} ({100*registros_saltados/total:.1f}%)")
    print(f"  🔄 Procesados en esta ejecución: {contador_procesados}")
    duracion_total = time.perf_counter() - inicio_flujo
    print(f"  ⏱️  Tiempo total de flujo: {formatear_duracion(duracion_total)} ({duracion_total:.1f}s)")
    
    # Desglose por detalle
    print(f"\n📌 Detalles de No Activos (top 5):")
    detalles_no_activos = df_normalizado[df_normalizado['estado'] != 'Activo']['detalle_validacion'].value_counts().head(5)
    for detalle, count in detalles_no_activos.items():
        print(f"   • {detalle}: {count}")
    
    print("\n" + "="*70)


if __name__ == "__main__":
    procesar_todas_credenciales()
 