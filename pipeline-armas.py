import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time
import re
import unicodedata
import itertools
from datetime import date

try:
    import pandas as pd
except ImportError:
    pd = None

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

URL_LOGIN = "https://www.sucamec.gob.pe/sel/faces/login.xhtml?faces-redirect=true"
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join("data", "programaciones-armas.xlsx"))

CREDENCIALES = {
    "tipo_documento_valor": os.getenv("TIPO_DOC", "RUC"),
    "numero_documento": os.getenv("NUMERO_DOCUMENTO", ""),
    "usuario": os.getenv("USUARIO_SEL", ""),
    "contrasena": os.getenv("CLAVE_SEL", ""),
}

CREDENCIALES_SELVA = {
    "tipo_documento_valor": os.getenv("SELVA_TIPO_DOC", "RUC"),
    "numero_documento": os.getenv("SELVA_NUMERO_DOCUMENTO", ""),
    "usuario": os.getenv("SELVA_USUARIO_SEL", ""),
    "contrasena": os.getenv("SELVA_CLAVE_SEL", ""),
}

SEL = {
    "tab_tradicional": 'a[href="#tabViewLogin:j_idt33"]',
    "tipo_doc_select": "#tabViewLogin\\:tradicionalForm\\:tipoDoc_input",
    "numero_documento": "#tabViewLogin\\:tradicionalForm\\:documento",
    "usuario": "#tabViewLogin\\:tradicionalForm\\:usuario",
    "clave": "#tabViewLogin\\:tradicionalForm\\:clave",
    "captcha_img": "#tabViewLogin\\:tradicionalForm\\:imgCaptcha",
    "captcha_input": "#tabViewLogin\\:tradicionalForm\\:textoCaptcha",
    "boton_refresh": "#tabViewLogin\\:tradicionalForm\\:botonCaptcha",
    "ingresar": "#tabViewLogin\\:tradicionalForm\\:ingresar",

    # ── Menú PanelMenu PrimeFaces ─────────────────────────────────────────────
    # Header del acordeón CITAS  →  el <h3> que contiene el <a>CITAS</a>
    # Hacemos clic en él para expandir/colapsar el panel
    "menu_citas_header": '#j_idt11\\:menuPrincipal .ui-panelmenu-header:has(a:text-is("CITAS"))',

    # Panel de contenido que se despliega al hacer clic en el header CITAS
    # id fijo según el HTML: j_idt11:menuPrincipal_7
    "menu_citas_panel": '#j_idt11\\:menuPrincipal_7',

    # Ítem "RESERVAS DE CITAS" — usa el onclick con menuid='7_1'
    # Selector más robusto: busca dentro del panel CITAS el span con ese texto
    "submenu_reservas": '#j_idt11\\:menuPrincipal_7 span.ui-menuitem-text:text-is("RESERVAS DE CITAS")',

    # ── SelectOneMenu: tipo de cita en Gestión de Citas ──────────────────────
    "tipo_cita_trigger": '#gestionCitasForm\\:j_idt32 .ui-selectonemenu-trigger',
    "tipo_cita_panel": '#gestionCitasForm\\:j_idt32_panel',
    "tipo_cita_label": '#gestionCitasForm\\:j_idt32_label',
    "tipo_cita_opcion_poligono": '#gestionCitasForm\\:j_idt32_panel li[data-label="EXAMEN PARA POLÍGONO DE TIRO"]',

    # ── Reserva de Cupos (tabGestion:creaCitaPolJurForm) ───────────────────
    "reserva_form": '#tabGestion\\:creaCitaPolJurForm',
    "sede_trigger": '#tabGestion\\:creaCitaPolJurForm\\:sedeId .ui-selectonemenu-trigger',
    "sede_panel": '#tabGestion\\:creaCitaPolJurForm\\:sedeId_panel',
    "sede_label": '#tabGestion\\:creaCitaPolJurForm\\:sedeId_label',
    "fecha_trigger": '#tabGestion\\:creaCitaPolJurForm\\:listaDiasId .ui-selectonemenu-trigger',
    "fecha_panel": '#tabGestion\\:creaCitaPolJurForm\\:listaDiasId_panel',
    "fecha_label": '#tabGestion\\:creaCitaPolJurForm\\:listaDiasId_label',

    # ── Tabla de programación de cupos ──────────────────────────────────────
    "tabla_programacion": '#tabGestion\\:creaCitaPolJurForm\\:dtProgramacion',
    "tabla_programacion_rows": '#tabGestion\\:creaCitaPolJurForm\\:dtProgramacion_data tr',
    "boton_siguiente": '#tabGestion\\:creaCitaPolJurForm button:has-text("Siguiente")',
    "boton_limpiar": '#tabGestion\\:creaCitaPolJurForm\\:botonLimpiar',

    # ── Paso 2 del Wizard ───────────────────────────────────────────────────
    "tipo_operacion_trigger": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe .ui-selectonemenu-trigger',
    "tipo_operacion_panel": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe_panel',
    "tipo_operacion_items": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe_panel li.ui-selectonemenu-item',
    "tipo_operacion_label": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe_label',
    "tipo_tramite_trigger": '#tabGestion\\:creaCitaPolJurForm\\:tipoTramite .ui-selectonemenu-trigger',
    "tipo_tramite_panel": '#tabGestion\\:creaCitaPolJurForm\\:tipoTramite_panel',
    "tipo_tramite_label": '#tabGestion\\:creaCitaPolJurForm\\:tipoTramite_label',
    "tipo_tramite_seg_priv": '#tabGestion\\:creaCitaPolJurForm\\:tipoTramite_panel li[data-label="SEGURIDAD PRIVADA"]',
    "doc_vig_input": '#tabGestion\\:creaCitaPolJurForm\\:nroDocVig_input',
    "doc_vig_panel": '#tabGestion\\:creaCitaPolJurForm\\:nroDocVig_panel',
    "doc_vig_items": '#tabGestion\\:creaCitaPolJurForm\\:nroDocVig_panel li.ui-autocomplete-item',
    "seleccione_solicitud_trigger": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud .ui-selectonemenu-trigger',
    "seleccione_solicitud_panel": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud_panel',
    "seleccione_solicitud_si": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud_panel li[id$="_1"]',
    "seleccione_solicitud_label": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud_label',
    "nro_solicitud_trigger": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud .ui-selectonemenu-trigger',
    "nro_solicitud_panel": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud_panel',
    "nro_solicitud_items": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud_panel li.ui-selectonemenu-item',
    "nro_solicitud_label": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud_label',

    # ── Paso 3 del Wizard (Resumen de Cita) ───────────────────────────────
    "fase3_panel": '#tabGestion\\:creaCitaPolJurForm\\:panelPaso4',
    "fase3_captcha_img": '#tabGestion\\:creaCitaPolJurForm\\:imgCaptcha',
    "fase3_captcha_input": '#tabGestion\\:creaCitaPolJurForm\\:textoCaptcha',
    "fase3_boton_refresh": '#tabGestion\\:creaCitaPolJurForm\\:botonCaptcha',
    "fase3_terminos_box": '#tabGestion\\:creaCitaPolJurForm\\:terminos .ui-chkbox-box',
    "fase3_terminos_input": '#tabGestion\\:creaCitaPolJurForm\\:terminos_input',
    "fase3_boton_generar_cita": '#tabGestion\\:creaCitaPolJurForm\\:j_idt561',
}


class SinCupoError(Exception):
    """Se lanza cuando la hora objetivo existe pero no tiene cupos libres."""


# ============================================================
# OCR helpers  (sin cambios)
# ============================================================

def corregir_captcha_ocr(texto_raw: str) -> str:
    if not texto_raw:
        return ""
    texto = texto_raw.strip().upper().replace(" ", "").replace("\n", "").replace("\r", "")
    texto = ''.join(c for c in texto if c.isalnum())
    return texto


def validar_captcha_texto(texto: str) -> bool:
    if not texto or len(texto) != 5:
        return False
    return texto.isalnum()


def captcha_fuzzy_normalize(texto: str) -> str:
    """
    Normalización suave para comparar candidatos OCR de CAPTCHA.
    No se usa directamente como resultado final, solo para puntuar consenso.
    """
    mapa = {
        "O": "0", "Q": "0", "D": "0",
        "I": "1", "L": "1",
        "Z": "2",
        "S": "5",  # para consenso, S suele confundirse con 5/8
        "T": "7",  # en este captcha T suele confundirse con 7
        "B": "8",
        "G": "6",
    }
    base = ''.join(c for c in str(texto or "").upper() if c.isalnum())
    return ''.join(mapa.get(c, c) for c in base)


def generar_candidatos_len5(texto: str) -> set:
    """Genera candidatos de longitud 5 desde una lectura OCR cruda."""
    limpio = ''.join(c for c in str(texto or "").upper() if c.isalnum())
    candidatos = set()

    if len(limpio) == 5:
        candidatos.add(limpio)

    if 6 <= len(limpio) <= 8:
        # Si OCR mete caracteres extra, probamos podas hasta len=5.
        quitar = len(limpio) - 5
        for idxs in itertools.combinations(range(len(limpio)), quitar):
            rec = ''.join(ch for i, ch in enumerate(limpio) if i not in idxs)
            if len(rec) == 5 and rec.isalnum():
                candidatos.add(rec)

    # Expansión por confusiones frecuentes (solo para casos ya len=5).
    expandidos = set(candidatos)
    swaps = {
        "0": ["O", "Q", "D"],
        "1": ["I", "L"],
        "2": ["Z"],
        "3": ["E"],
        "6": ["G"],
        "7": ["T"],
        "8": ["B", "S"],
        "5": ["S"],
        "E": ["3"],
        "B": ["8"],
    }
    for c in list(candidatos):
        for i, ch in enumerate(c):
            for alt in swaps.get(ch, []):
                expandidos.add(c[:i] + alt + c[i+1:])

    return expandidos


def seleccionar_mejor_captcha_por_consenso(observaciones: list) -> str:
    """Elige el mejor candidato len=5 por consenso entre varias lecturas OCR."""
    if not observaciones:
        return ""

    sets_obs = []
    for obs in observaciones:
        candidatos = generar_candidatos_len5(obs)
        if candidatos:
            sets_obs.append(candidatos)

    if not sets_obs:
        return ""

    universo = set().union(*sets_obs)
    mejor = ""
    mejor_score = -1
    mejor_exact = -1

    for cand in universo:
        cand_fuzzy = captcha_fuzzy_normalize(cand)
        score = 0
        exact = 0
        for cands_obs in sets_obs:
            fuzzy_obs = {captcha_fuzzy_normalize(x) for x in cands_obs}
            if cand_fuzzy in fuzzy_obs:
                score += 1
            if cand in cands_obs:
                exact += 1

        if (score > mejor_score) or (score == mejor_score and exact > mejor_exact):
            mejor = cand
            mejor_score = score
            mejor_exact = exact

    return mejor if validar_captcha_texto(mejor) else ""


def medir_consenso_captcha(candidato: str, observaciones: list) -> tuple:
    """Devuelve (fuzzy_hits, exact_hits, total_observaciones_validas)."""
    if not candidato:
        return 0, 0, 0

    sets_obs = []
    for obs in observaciones:
        candidatos = generar_candidatos_len5(obs)
        if candidatos:
            sets_obs.append(candidatos)

    if not sets_obs:
        return 0, 0, 0

    cand_fuzzy = captcha_fuzzy_normalize(candidato)
    fuzzy_hits = 0
    exact_hits = 0
    for cands_obs in sets_obs:
        fuzzy_obs = {captcha_fuzzy_normalize(x) for x in cands_obs}
        if cand_fuzzy in fuzzy_obs:
            fuzzy_hits += 1
        if candidato in cands_obs:
            exact_hits += 1

    return fuzzy_hits, exact_hits, len(sets_obs)


def captcha_tiene_ambiguedad(texto: str) -> bool:
    """Detecta caracteres con alta confusión visual para decidir refresh de captcha."""
    t = ''.join(c for c in str(texto or "").upper() if c.isalnum())
    if len(t) != 5:
        return True

    grupos_ambiguos = [
        set("A4"),
        set("1I"),
        set("I7"),
        set("S8"),
        set("S5"),
    ]

    for ch in t:
        for grupo in grupos_ambiguos:
            if ch in grupo:
                return True
    return False


def escribir_input_jsf(page, selector: str, valor: str):
    for intento in range(4):
        campo = page.locator(selector)
        campo.wait_for(state="visible", timeout=12000)

        # Intento principal: tipeo humano con delay alto para evitar pérdida de dígitos.
        campo.click()
        campo.press("Control+A")
        campo.press("Backspace")
        campo.type(valor, delay=65)
        campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
        page.wait_for_timeout(140)

        actual = campo.input_value().strip()
        if actual != valor:
            # Fallback fuerte: asignación directa del value y eventos JSF.
            campo.evaluate(
                '''(el, val) => {
                    el.focus();
                    el.value = val;
                    el.setAttribute("value", val);
                    el.dispatchEvent(new Event("input", { bubbles: true }));
                    el.dispatchEvent(new Event("change", { bubbles: true }));
                }''',
                valor
            )
            page.wait_for_timeout(120)
            actual = campo.input_value().strip()

        if actual == valor:
            # Dispara blur al final para el comportamiento JSF de validación.
            campo.evaluate('el => el.blur()')
            page.wait_for_timeout(220)
            try:
                confirmado = page.locator(selector).input_value().strip()
            except Exception:
                confirmado = ""
            if confirmado == valor:
                return
            actual = confirmado

        print(f"   ⚠️ Campo {selector}: esperado '{valor}', tiene '{actual}' → reintentando ({intento+1}/4)")
        page.wait_for_timeout(260)

    raise Exception(f"No se pudo fijar correctamente el valor del campo {selector}")


def escribir_input_rapido(page, selector: str, valor: str):
    campo = page.locator(selector)
    campo.wait_for(state="visible", timeout=10000)
    campo.click()
    campo.fill(valor)
    campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
    campo.blur()
    if campo.input_value() != valor:
        campo.click()
        campo.press("Control+A")
        campo.press("Backspace")
        campo.type(valor, delay=10)
        campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
        campo.blur()


def solve_captcha_manual(page):
    print("\n🔴 MODO MANUAL ACTIVADO")
    print("Llena el código de verificación en la ventana del navegador")
    input("✅ Cuando hayas escrito el CAPTCHA → presiona ENTER aquí para continuar...")


def preprocesar_imagen_captcha(img_bytes: bytes, variante: int = 0) -> 'Image':
    img = Image.open(BytesIO(img_bytes))
    img = img.convert('L')
    if variante == 0:
        img = ImageEnhance.Contrast(img).enhance(3.5)
        w, h = img.size
        img = img.resize((w * 4, h * 4), Image.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=3))
        img = ImageOps.invert(img)
        img = img.point(lambda p: 255 if p > 130 else 0)
        img = ImageEnhance.Sharpness(img).enhance(3.0)
    elif variante == 1:
        img = ImageEnhance.Contrast(img).enhance(2.5)
        w, h = img.size
        img = img.resize((w * 3, h * 3), Image.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=5))
        img = img.point(lambda p: 255 if p > 160 else 0)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
    else:
        img = ImageEnhance.Contrast(img).enhance(4.0)
        w, h = img.size
        img = img.resize((w * 5, h * 5), Image.LANCZOS)
        img = img.filter(ImageFilter.GaussianBlur(radius=0.5))
        img = ImageOps.invert(img)
        img = img.point(lambda p: 255 if p > 110 else 0)
        img = ImageEnhance.Sharpness(img).enhance(4.0)
    return img


def solve_captcha_ocr_base(
    page,
    captcha_img_selector: str,
    boton_refresh_selector: str = None,
    contexto: str = "CAPTCHA",
    evitar_ambiguos: bool = False,
    min_fuzzy_hits: int = 0,
    max_intentos=6,
):
    """
    Motor OCR estilo login: acepta la primera lectura válida por intento.
    Si evitar_ambiguos=True, aplica un filtro adicional antes de aceptar.
    """
    if not OCR_AVAILABLE:
        return None

    PSM_MODES = [7, 8, 13]
    NUM_VARIANTES = 3

    intento = 0
    while True:
        if max_intentos is not None and max_intentos > 0 and intento >= max_intentos:
            break
        intento += 1
        try:
            total_txt = str(max_intentos) if (max_intentos is not None and max_intentos > 0) else "∞"
            print(f"🔍 Intentando resolver {contexto} (intento {intento}/{total_txt})...")
            page.wait_for_timeout(200)
            img_bytes = page.locator(captcha_img_selector).screenshot(type="png")

            mejor_texto = None
            observaciones = []
            for variante in range(NUM_VARIANTES):
                img = preprocesar_imagen_captcha(img_bytes, variante=variante)
                for psm in PSM_MODES:
                    config = f'--psm {psm} --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ --dpi 300'
                    texto_raw = pytesseract.image_to_string(img, config=config, lang='eng').strip()
                    texto = corregir_captcha_ocr(texto_raw)
                    observaciones.append(texto)

                    if validar_captcha_texto(texto):
                        print(f"   → Variante {variante}, PSM {psm}: '{texto_raw}' → '{texto}' ✓")
                        mejor_texto = texto
                        break
                    else:
                        print(f"   → Variante {variante}, PSM {psm}: '{texto_raw}' → '{texto}' (len={len(texto)}) ✗")
                if mejor_texto:
                    break

            if mejor_texto:
                if evitar_ambiguos:
                    fuzzy_hits, exact_hits, total_hits = medir_consenso_captcha(mejor_texto, observaciones)
                    print(f"   ℹ️ Consenso OCR: fuzzy={fuzzy_hits}/{total_hits}, exacto={exact_hits}/{total_hits}")

                    es_ambiguo = captcha_tiene_ambiguedad(mejor_texto)
                    consenso_debil = total_hits > 0 and fuzzy_hits < min_fuzzy_hits

                    if es_ambiguo or consenso_debil:
                        motivo = "ambiguo" if es_ambiguo else "consenso débil"
                        print(f"   ⚠️ CAPTCHA {motivo} detectado ('{mejor_texto}') → se solicitará uno nuevo")
                        if boton_refresh_selector:
                            page.locator(boton_refresh_selector).click(force=True)
                            page.wait_for_timeout(500)
                            continue

                print(f"   ✓ CAPTCHA válido → Usando: {mejor_texto}")
                return mejor_texto

            if boton_refresh_selector:
                print("   ✗ Ninguna combinación dio resultado → Refrescando CAPTCHA...")
                print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                page.locator(boton_refresh_selector).click(force=True)
                page.wait_for_timeout(500)
            else:
                print("   ✗ Ninguna combinación dio resultado (sin botón refresh configurado)")

        except Exception as e:
            print(f"   Error en intento {intento}: {str(e)}")
            page.wait_for_timeout(300)

    if max_intentos is None or max_intentos <= 0:
        print(f"❌ No se pudo resolver {contexto} automáticamente (modo sin límite agotado por salida externa) → modo manual")
    else:
        print(f"❌ No se pudo resolver {contexto} automáticamente después de {max_intentos} intentos → modo manual")
    return None


def solve_captcha_ocr_generico(
    page,
    captcha_img_selector: str,
    boton_refresh_selector: str = None,
    contexto: str = "CAPTCHA",
    evitar_ambiguos: bool = False,
):
    return solve_captcha_ocr_base(
        page,
        captcha_img_selector=captcha_img_selector,
        boton_refresh_selector=boton_refresh_selector,
        contexto=contexto,
        evitar_ambiguos=evitar_ambiguos,
        min_fuzzy_hits=6,
    )


def solve_captcha_ocr(page):
    """Lógica original estable del login: primera lectura válida por intento."""
    return solve_captcha_ocr_base(
        page,
        captcha_img_selector=SEL["captcha_img"],
        boton_refresh_selector=SEL["boton_refresh"],
        contexto="CAPTCHA",
        evitar_ambiguos=False,
        min_fuzzy_hits=0,
    )


def completar_fase_3_resumen(page):
    """Paso 3: resolver captcha del resumen y aceptar términos y condiciones."""
    print("\n🧾 Completando Fase 3 (Resumen de cita)...")

    page.locator(SEL["fase3_panel"]).wait_for(state="visible", timeout=12000)

    captcha_text = solve_captcha_ocr_base(
        page,
        captcha_img_selector=SEL["fase3_captcha_img"],
        boton_refresh_selector=None,
        contexto="CAPTCHA Fase 3",
        evitar_ambiguos=False,
        min_fuzzy_hits=0,
        max_intentos=None,
    )

    if captcha_text and len(captcha_text) == 5:
        escribir_input_rapido(page, SEL["fase3_captcha_input"], captcha_text)
        print(f"   ✓ CAPTCHA Fase 3 escrito: {captcha_text}")
    else:
        print("   ⚠️ OCR no resolvió CAPTCHA Fase 3; usa ingreso manual en el navegador")
        solve_captcha_manual(page)

    checkbox_input = page.locator(SEL["fase3_terminos_input"])
    checkbox_box = page.locator(SEL["fase3_terminos_box"])
    checkbox_box.wait_for(state="visible", timeout=7000)

    marcado = False
    try:
        marcado = checkbox_input.is_checked()
    except Exception:
        marcado = False

    if not marcado:
        checkbox_box.click()
        page.wait_for_timeout(180)

    try:
        marcado = checkbox_input.is_checked()
    except Exception:
        marcado = False

    if not marcado:
        clase_box = checkbox_box.get_attribute("class") or ""
        if "ui-state-active" in clase_box:
            marcado = True

    if not marcado:
        raise Exception("No se pudo marcar 'Acepto los términos y condiciones de Sucamec'")

    print("   ✓ Términos y condiciones marcados")


def limpiar_para_siguiente_registro(page, motivo: str = ""):
    """Pulsa botón Limpiar para reiniciar el wizard y seguir con el siguiente registro."""
    boton_limpiar = page.locator(SEL["boton_limpiar"])
    boton_limpiar.wait_for(state="visible", timeout=8000)
    boton_limpiar.first.click(timeout=8000)
    page.wait_for_timeout(180)
    if motivo:
        print(f"   ✓ Click en 'Limpiar' ({motivo})")
    else:
        print("   ✓ Click en 'Limpiar'")


def generar_cita_final_con_reintento_rapido(page, max_intentos: int = 3):
    """
    Paso final opcional (desactivado por ahora en el flujo principal).
    Hace click en 'Generar Cita' y, si detecta error de captcha/validación,
    reintenta rápido regenerando el captcha de Fase 3.
    """
    print("\n🧪 Paso final opcional: Generar Cita (reintento rápido)")

    boton_generar = page.locator(SEL["fase3_boton_generar_cita"])
    boton_generar.wait_for(state="visible", timeout=10000)

    for intento in range(1, max_intentos + 1):
        inicio_validacion = time.time()
        print(f"   ↻ Intento generar cita {intento}/{max_intentos}")
        boton_generar.click(timeout=10000)

        url_ok = False
        try:
            # Validación rápida: si avanza de vista, debe ocurrir casi inmediato.
            page.wait_for_url("**/aplicacion/**", timeout=1000)
            if "GestionCitas.xhtml" not in page.url:
                url_ok = True
        except PlaywrightTimeoutError:
            url_ok = False

        if url_ok:
            tiempo = time.time() - inicio_validacion
            print(f"   ✓ Generar Cita confirmado en {tiempo:.2f}s")
            print(f"   → URL: {page.url}")
            return True

        mensaje_error = ""
        for selector in [
            ".ui-messages-error",
            ".ui-message-error",
            ".ui-growl-message-error",
            "[class*='error']",
            ".mensajeError",
        ]:
            try:
                loc = page.locator(selector)
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

        tiempo = time.time() - inicio_validacion
        if mensaje_error:
            print(f"   ⚠️ Mensaje detectado: {mensaje_error}")
        print(f"   ⏱️ Validación final: {tiempo:.2f}s")

        error_captcha = bool(re.search(r"captcha|c[oó]digo|validaci[oó]n", mensaje_error, flags=re.IGNORECASE))
        if not error_captcha:
            raise Exception("No se pudo confirmar la generación de cita (sin error de captcha detectable)")

        # Reintento rápido: resolver captcha de Fase 3 y remarcado de términos si aplica.
        nuevo_captcha = solve_captcha_ocr_base(
            page,
            captcha_img_selector=SEL["fase3_captcha_img"],
            boton_refresh_selector=SEL["fase3_boton_refresh"],
            contexto="CAPTCHA Fase 3 (reintento final)",
            evitar_ambiguos=False,
            min_fuzzy_hits=0,
            max_intentos=3,
        )

        if nuevo_captcha and len(nuevo_captcha) == 5:
            escribir_input_rapido(page, SEL["fase3_captcha_input"], nuevo_captcha)
            print(f"   ✓ CAPTCHA reintento escrito: {nuevo_captcha}")
        else:
            print("   ⚠️ OCR no resolvió captcha en reintento final; pasar a ingreso manual")
            solve_captcha_manual(page)

        try:
            if not page.locator(SEL["fase3_terminos_input"]).is_checked():
                page.locator(SEL["fase3_terminos_box"]).click()
                page.wait_for_timeout(150)
        except Exception:
            pass

    raise Exception("No se pudo generar cita tras reintentos rápidos")


def validar_resultado_login_por_ui(page, timeout_ms: int = 3000):
    """
    Determina resultado de login por señales de UI (no por URL):
    - Éxito: aparece menú principal/controles de sesión autenticada.
    - Falla: aparece mensaje de error de validación/captcha.
    Devuelve: (login_ok: bool, mensaje_error: str|None, tiempo_segundos: float)
    """
    inicio = time.time()

    selectores_exito = [
        "#j_idt11\\:menuPrincipal",
        "#j_idt11\\:j_idt18",  # botón "Cerrar Sesión"
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

    # Última comprobación rápida al vencer el timeout.
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


def pagina_muestra_servicio_no_disponible(page) -> bool:
    """Detecta HTML de caída del servicio (HTTP 503 / Service Unavailable)."""
    # Señales de estado saludable: si aparecen, no hay caída.
    selectores_ok = [
        SEL["tab_tradicional"],
        SEL["numero_documento"],
        "#j_idt11\\:menuPrincipal",
        "form#gestionCitasForm",
        SEL["reserva_form"],
    ]
    for sel in selectores_ok:
        try:
            loc = page.locator(sel)
            if loc.count() > 0 and loc.first.is_visible():
                return False
        except Exception:
            pass

    # Título de pestaña en estados 503.
    try:
        titulo = (page.title() or "").strip().upper()
        if "SERVICE UNAVAILABLE" in titulo:
            return True
    except Exception:
        pass

    # h1 explícito del error de Apache/Proxy.
    try:
        h1 = (page.locator("h1").first.inner_text() or "").strip().upper()
        if "SERVICE UNAVAILABLE" in h1:
            return True
    except Exception:
        pass

    # Fallback textual ligero (evita leer todo el HTML para no ralentizar iteraciones).
    try:
        body_text = (page.locator("body").inner_text() or "").upper()
        if "SERVICE UNAVAILABLE" in body_text and "AUTENTICACION TRADICIONAL" not in body_text:
            return True
    except Exception:
        pass

    return False


def esperar_hasta_servicio_disponible(page, url_objetivo: str, espera_segundos: int = 8):
    """
    Si SUCAMEC responde con Service Unavailable, espera y reintenta hasta recuperar servicio.
    Se mantiene en bucle indefinido por requerimiento operativo.
    """
    intento = 0
    while pagina_muestra_servicio_no_disponible(page):
        intento += 1
        print(f"⚠️ SUCAMEC no disponible (Service Unavailable). Reintento {intento} en {espera_segundos}s...")
        time.sleep(espera_segundos)
        try:
            page.goto(url_objetivo, wait_until="domcontentloaded", timeout=45000)
        except Exception as e:
            print(f"   ↳ Error al reintentar acceso: {e}")


def normalizar_fecha_excel(valor_fecha: str) -> str:
    """Convierte fechas de Excel al formato dd/mm/yyyy esperado por SEL."""
    texto = str(valor_fecha or "").strip()
    if not texto:
        return ""

    # Si ya viene en formato con barras, asumimos entrada local dd/mm/yyyy.
    if "/" in texto:
        dt = pd.to_datetime(texto, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")

    # Caso típico de Excel ISO: 2026-03-31 00:00:00 / 2026-03-31
    dt = pd.to_datetime(texto, errors="coerce", dayfirst=False)
    if pd.notna(dt):
        return dt.strftime("%d/%m/%Y")

    # Si ya viene como dd/mm/yyyy o similar, lo conservamos sin hora
    texto = texto.split(" ")[0]
    return texto


def normalizar_hora_fragmento(valor_hora: str) -> str:
    """Normaliza una hora a HH:MM (ej: 8:5 -> 08:05)."""
    texto = str(valor_hora or "").strip().replace(".", ":")
    if ":" not in texto:
        return texto
    partes = texto.split(":")
    if len(partes) != 2:
        return texto
    try:
        hh = int(partes[0])
        mm = int(partes[1])
    except ValueError:
        return texto
    return f"{hh:02d}:{mm:02d}"


def normalizar_hora_rango(valor_rango: str) -> str:
    """Normaliza rango de hora a HH:MM-HH:MM."""
    texto = str(valor_rango or "").strip()
    if not texto:
        return ""
    texto = texto.replace("–", "-").replace("—", "-").replace(" a ", "-").replace(" ", "")
    partes = texto.split("-")
    if len(partes) != 2:
        return texto
    inicio = normalizar_hora_fragmento(partes[0])
    fin = normalizar_hora_fragmento(partes[1])
    return f"{inicio}-{fin}"


def convertir_a_entero(texto: str) -> int:
    numeros = re.findall(r"\d+", str(texto or ""))
    return int(numeros[0]) if numeros else 0


def normalizar_texto_comparable(texto: str) -> str:
    base = str(texto or "").strip().upper()
    base = unicodedata.normalize("NFKD", base)
    base = "".join(c for c in base if not unicodedata.combining(c))
    base = re.sub(r"\s+", " ", base)
    return base


def limpiar_valor_excel(valor: str) -> str:
    """Limpia artefactos comunes de celdas Excel exportadas a texto."""
    t = str(valor or "")
    t = re.sub(r"_x[0-9A-Fa-f]{4}_", "", t)
    t = t.replace("\r", " ").replace("\n", " ")
    t = re.sub(r"\s+", " ", t).strip()
    return t


def extraer_token_solicitud(valor: str) -> str:
    """Obtiene el número principal de solicitud para comparar dentro del label del combo."""
    texto = str(valor or "")
    grupos = re.findall(r"\d+", texto)
    if not grupos:
        return ""
    token = grupos[0].lstrip("0")
    return token if token else "0"


def normalizar_tipo_arma_excel(valor: str) -> str:
    """Normaliza valor de tipo_arma del Excel para comparaciones."""
    base = normalizar_texto_comparable(valor)
    equivalencias = {
        "LARG": "LARGA",
        "LARGA": "LARGA",
        "CORTA": "CORTA",
        "PISTOLA": "PISTOLA",
        "REVOLVER": "REVOLVER",
        "CARABINA": "CARABINA",
        "ESCOPETA": "ESCOPETA",
    }
    return equivalencias.get(base, base)


def inferir_objetivo_arma_desde_excel(valor: str) -> str:
    """
    Interpreta texto libre de tipo_arma y devuelve una clave usable.
    Ejemplos válidos: "CORTA", "CORTA PISTOLA", "LARGA ESCOPETA".
    """
    base = normalizar_texto_comparable(valor)
    if not base:
        return ""

    # Priorizamos el arma específica si está presente.
    if "ESCOPETA" in base:
        return "ESCOPETA"
    if "CARABINA" in base:
        return "CARABINA"
    if "REVOLVER" in base:
        return "REVOLVER"
    if "PISTOLA" in base:
        return "PISTOLA"

    # Si no hay arma específica, devolvemos tipo general.
    if "LARG" in base:
        return "LARGA"
    if "CORT" in base:
        return "CORTA"

    return normalizar_tipo_arma_excel(base)


def fecha_comparable(valor_fecha: str) -> str:
    """Convierte fecha de Excel a una cadena comparable dd/mm/yyyy."""
    return normalizar_fecha_excel(valor_fecha)


def normalizar_ruc_operativo(valor_ruc: str) -> str:
    """Normaliza texto de RUC/razón social para clasificación operativa."""
    return normalizar_texto_comparable(limpiar_valor_excel(valor_ruc))


def obtener_grupo_ruc(valor_ruc: str) -> str:
    """Clasifica el RUC/razón social en SELVA, JV u OTRO."""
    base = normalizar_ruc_operativo(valor_ruc)
    if "SELVA" in base or "20493762789" in base:
        return "SELVA"
    if "J&V" in base or "J V" in base or "RESGUARDO" in base or "20100901481" in base:
        return "JV"
    return "OTRO"


def prioridad_orden(valor_prioridad: str) -> int:
    """ALTA tiene precedencia sobre NORMAL; cualquier otro valor cae en NORMAL."""
    base = normalizar_texto_comparable(limpiar_valor_excel(valor_prioridad))
    return 0 if base == "ALTA" else 1


def sede_es_lima(valor_sede: str) -> bool:
    """Detecta si la sede corresponde a Lima para priorización operativa."""
    base = normalizar_texto_comparable(limpiar_valor_excel(valor_sede))
    return "LIMA" in base


def orden_prioridad_geografica(valor_prioridad: str, valor_sede: str) -> int:
    """
    Orden requerido dentro de cada razón social:
    0) Lima + Alta
    1) Provincia + Alta
    2) Lima + Normal
    3) Provincia + Normal
    """
    prio = prioridad_orden(valor_prioridad)  # 0=Alta, 1=Normal
    es_lima = sede_es_lima(valor_sede)

    if prio == 0 and es_lima:
        return 0
    if prio == 0 and not es_lima:
        return 1
    if prio == 1 and es_lima:
        return 2
    return 3


def resolver_credenciales_por_grupo_ruc(grupo_ruc: str) -> dict:
    if grupo_ruc == "SELVA":
        return CREDENCIALES_SELVA
    return CREDENCIALES


def obtener_trabajos_pendientes_excel(ruta_excel: str) -> list:
    """
    Devuelve trabajos pendientes únicos y ordenados por prioridad operativa:
    1) SELVA primero
    2) dentro de cada razón social: Lima Alta -> Provincia Alta -> Lima Normal -> Provincia Normal
    3) orden original del Excel como desempate
    """
    if pd is None:
        raise Exception("Falta dependencia 'pandas'. Instala con: pip install pandas openpyxl")
    if not os.path.exists(ruta_excel):
        raise Exception(f"No se encontró el Excel en: {ruta_excel}")

    df = pd.read_excel(ruta_excel, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    if "estado" not in df.columns:
        raise Exception("El Excel no contiene la columna 'estado'")

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).apply(limpiar_valor_excel)

    if "doc_vigilante" not in df.columns:
        df["doc_vigilante"] = ""
    if "dni" not in df.columns:
        df["dni"] = ""
    if "nro_solicitud" not in df.columns:
        df["nro_solicitud"] = ""
    if "sede" not in df.columns:
        df["sede"] = ""
    if "ruc" not in df.columns:
        df["ruc"] = ""
    if "prioridad" not in df.columns:
        df["prioridad"] = "Normal"

    fecha_col_programacion = "fecha_programacion" if "fecha_programacion" in df.columns else "fecha"
    if fecha_col_programacion not in df.columns:
        raise Exception("El Excel no contiene columna de fecha (fecha_programacion/fecha)")

    pendientes = df[df["estado"].str.upper().str.contains("PENDIENTE", na=False)].copy()
    if pendientes.empty:
        return []

    validar_hoy = os.getenv("VALIDAR_FECHA_PROGRAMACION_HOY", "1").strip().lower() in {"1", "true", "si", "sí", "yes"}
    if validar_hoy:
        hoy = date.today().strftime("%d/%m/%Y")
        pendientes = pendientes[
            pendientes[fecha_col_programacion].apply(fecha_comparable) == hoy
        ]
        if pendientes.empty:
            return []

    pendientes["_idx_excel"] = pendientes.index
    pendientes["_doc_norm"] = pendientes.apply(
        lambda r: str(r.get("doc_vigilante", "") or r.get("dni", "")).strip(),
        axis=1,
    )
    pendientes["_nro_norm"] = pendientes["nro_solicitud"].apply(lambda v: str(v or "").strip())
    pendientes["_fecha_prog"] = pendientes[fecha_col_programacion].apply(fecha_comparable)
    pendientes["_ruc_raw"] = pendientes["ruc"].apply(lambda v: str(v or "").strip())
    pendientes["_ruc_grupo"] = pendientes["_ruc_raw"].apply(obtener_grupo_ruc)
    pendientes["_ruc_orden"] = pendientes["_ruc_grupo"].map({"SELVA": 0, "JV": 1, "OTRO": 2})
    pendientes["_prioridad_raw"] = pendientes["prioridad"].apply(lambda v: str(v or "").strip())
    pendientes["_sede_raw"] = pendientes["sede"].apply(lambda v: str(v or "").strip())
    pendientes["_geo_prio_orden"] = pendientes.apply(
        lambda r: orden_prioridad_geografica(r.get("_prioridad_raw", ""), r.get("_sede_raw", "")),
        axis=1,
    )

    pendientes = pendientes.sort_values(
        by=["_ruc_orden", "_geo_prio_orden", "_idx_excel"],
        ascending=[True, True, True],
        kind="stable",
    )

    trabajos = []
    claves_vistas = set()
    for _, fila in pendientes.iterrows():
        clave = (
            fila.get("_doc_norm", ""),
            fila.get("_nro_norm", ""),
            fila.get("_fecha_prog", ""),
            fila.get("_ruc_grupo", "OTRO"),
        )
        if clave in claves_vistas:
            continue
        claves_vistas.add(clave)
        trabajos.append(
            {
                "idx_excel": int(fila.get("_idx_excel")),
                "ruc": fila.get("_ruc_raw", ""),
                "ruc_grupo": fila.get("_ruc_grupo", "OTRO"),
                "prioridad": fila.get("_prioridad_raw", "Normal"),
                "fecha_programacion": fila.get("_fecha_prog", ""),
            }
        )

    return trabajos


def obtener_indices_pendientes_excel(ruta_excel: str) -> list:
    """
    Devuelve índices de trabajo (únicos) en estado Pendiente.
    Deduplica por doc_vigilante+dni, nro_solicitud y fecha_programacion/fecha,
    para evitar reprocesar registros que pertenecen a una misma cita/iteración.
    """
    trabajos = obtener_trabajos_pendientes_excel(ruta_excel)
    return [t["idx_excel"] for t in trabajos]


def cargar_primer_registro_pendiente_desde_excel(ruta_excel: str, indice_excel_objetivo: int = None) -> dict:
    """
    Lee el Excel y devuelve el primer registro con estado 'Pendiente'.
    Campos mínimos requeridos para este paso: sede y fecha.
    """
    if pd is None:
        raise Exception("Falta dependencia 'pandas'. Instala con: pip install pandas openpyxl")

    if not os.path.exists(ruta_excel):
        raise Exception(f"No se encontró el Excel en: {ruta_excel}")

    df = pd.read_excel(ruta_excel, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    columnas_requeridas = {"sede", "fecha", "hora_rango", "tipo_operacion", "nro_solicitud", "tipo_arma", "arma", "estado"}
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        raise Exception(f"Faltan columnas requeridas en Excel: {faltantes}")

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).apply(limpiar_valor_excel)

    pendientes = df[df["estado"].str.upper().str.contains("PENDIENTE", na=False)]
    if pendientes.empty:
        raise Exception("No hay registros con estado 'Pendiente' en el Excel")

    indice_primer_pendiente = pendientes.index[0] if indice_excel_objetivo is None else indice_excel_objetivo
    if indice_primer_pendiente not in pendientes.index:
        raise Exception(f"El índice objetivo {indice_primer_pendiente} no está en estado Pendiente")

    registro = pendientes.loc[indice_primer_pendiente].to_dict()
    registro["_excel_index"] = int(indice_primer_pendiente)
    registro["_excel_path"] = ruta_excel

    fecha_col_programacion = "fecha_programacion" if "fecha_programacion" in df.columns else "fecha"
    fecha_programacion_valor = fecha_comparable(registro.get(fecha_col_programacion, registro.get("fecha", "")))

    sede = registro.get("sede", "").strip()
    fecha = normalizar_fecha_excel(registro.get("fecha", ""))
    hora_rango = normalizar_hora_rango(registro.get("hora_rango", ""))
    tipo_operacion = registro.get("tipo_operacion", "").strip()
    nro_solicitud = registro.get("nro_solicitud", "").strip()
    doc_vigilante = registro.get("doc_vigilante", registro.get("dni", "")).strip()
    tipo_arma_base = inferir_objetivo_arma_desde_excel(registro.get("tipo_arma", ""))
    arma_base = inferir_objetivo_arma_desde_excel(registro.get("arma", ""))

    if not sede or not fecha or not hora_rango:
        raise Exception("El registro pendiente no tiene 'sede', 'fecha' o 'hora_rango' con valor")
    if not tipo_operacion or not nro_solicitud or not doc_vigilante:
        raise Exception("El registro pendiente no tiene 'tipo_operacion', 'doc_vigilante/dni' o 'nro_solicitud'")
    if not tipo_arma_base:
        raise Exception("El registro pendiente no tiene 'tipo_arma'")
    if not arma_base:
        raise Exception("El registro pendiente no tiene 'arma'")

    # Agrupa registros de la misma programación/cita:
    # mismo usuario + misma solicitud + misma fecha_programacion/fecha.
    fecha_base = fecha_comparable(registro.get(fecha_col_programacion, registro.get("fecha", "")))
    doc_base = doc_vigilante
    nro_base = nro_solicitud
    pendientes_aux = pendientes.copy()
    pendientes_aux["fecha_norm"] = pendientes_aux[fecha_col_programacion].apply(fecha_comparable)
    pendientes_aux["doc_norm"] = pendientes_aux.apply(
        lambda r: str(r.get("doc_vigilante", "") or r.get("dni", "")).strip(), axis=1
    )
    pendientes_aux["nro_norm"] = pendientes_aux["nro_solicitud"].apply(lambda v: str(v or "").strip())
    relacionados = pendientes_aux[
        (pendientes_aux["fecha_norm"] == fecha_base) &
        (pendientes_aux["doc_norm"] == doc_base) &
        (pendientes_aux["nro_norm"] == nro_base)
    ]

    # Validación adicional: revisar explícitamente el siguiente registro.
    siguiente_mismo_doc_y_fecha = False
    siguiente_idx = indice_primer_pendiente + 1
    if siguiente_idx in df.index:
        fila_sig = df.loc[siguiente_idx]
        estado_sig = str(fila_sig.get("estado", "")).strip().upper()
        doc_sig = str(fila_sig.get("doc_vigilante", "") or fila_sig.get("dni", "")).strip()
        nro_sig = str(fila_sig.get("nro_solicitud", "")).strip()
        fecha_sig = fecha_comparable(fila_sig.get(fecha_col_programacion, fila_sig.get("fecha", "")))
        if estado_sig == "PENDIENTE" and doc_sig == doc_base and nro_sig == nro_base and fecha_sig == fecha_base:
            siguiente_mismo_doc_y_fecha = True

    tipos_arma_excel = []
    armas_excel = []
    objetivos_arma = []
    armas_especificas = {"PISTOLA", "REVOLVER", "CARABINA", "ESCOPETA"}

    for _, fila in relacionados.iterrows():
        tipo_raw = str(fila.get("tipo_arma", "")).strip()
        arma_raw = str(fila.get("arma", "")).strip()
        tipo_inferido = inferir_objetivo_arma_desde_excel(tipo_raw)
        arma_inferida = inferir_objetivo_arma_desde_excel(arma_raw)

        if not arma_inferida:
            arma_inferida = inferir_objetivo_arma_desde_excel(tipo_raw)

        tipo_norm_texto = normalizar_texto_comparable(tipo_raw)
        if arma_inferida in {"PISTOLA", "REVOLVER"}:
            tipo_fila = "CORTA"
        elif arma_inferida in {"CARABINA", "ESCOPETA"}:
            tipo_fila = "LARGA"
        elif "CORT" in tipo_norm_texto or tipo_inferido == "CORTA":
            tipo_fila = "CORTA"
        elif "LARG" in tipo_norm_texto or tipo_inferido == "LARGA":
            tipo_fila = "LARGA"
        else:
            continue

        if arma_inferida in armas_especificas:
            arma_objetivo = arma_inferida
        else:
            arma_objetivo = "PISTOLA" if tipo_fila == "CORTA" else "CARABINA"

        if tipo_fila not in tipos_arma_excel:
            tipos_arma_excel.append(tipo_fila)
        if arma_objetivo not in armas_excel:
            armas_excel.append(arma_objetivo)

        par_objetivo = (tipo_fila, arma_objetivo)
        if par_objetivo not in objetivos_arma:
            objetivos_arma.append(par_objetivo)

    if not objetivos_arma:
        # Fallback mínimo usando el primer registro, manteniendo origen en Excel.
        if arma_base in {"PISTOLA", "REVOLVER"}:
            tipo_base = "CORTA"
            arma_objetivo = arma_base
        elif arma_base in {"CARABINA", "ESCOPETA"}:
            tipo_base = "LARGA"
            arma_objetivo = arma_base
        elif tipo_arma_base == "LARGA":
            tipo_base = "LARGA"
            arma_objetivo = "CARABINA"
        else:
            tipo_base = "CORTA"
            arma_objetivo = "PISTOLA"

        objetivos_arma = [(tipo_base, arma_objetivo)]
        tipos_arma_excel = [tipo_base]
        armas_excel = [arma_objetivo]

    tipos_arma_objetivo = [t for t, _ in objetivos_arma]

    registro["fecha"] = fecha
    registro["hora_rango"] = hora_rango
    registro["doc_vigilante"] = doc_vigilante
    registro["fecha_programacion"] = fecha_programacion_valor
    registro["ruc"] = registro.get("ruc", "")
    registro["prioridad"] = registro.get("prioridad", "")
    registro["objetivos_arma"] = objetivos_arma
    registro["tipos_arma_objetivo"] = tipos_arma_objetivo
    registro["armas_objetivo"] = armas_excel

    print("📄 Registro tomado desde Excel:")
    print(f"   • id_registro: {registro.get('id_registro', '')}")
    print(f"   • sede: {sede}")
    print(f"   • fecha: {fecha}")
    print(f"   • hora_rango: {hora_rango}")
    print(f"   • tipo_operacion: {tipo_operacion}")
    print(f"   • doc_vigilante: {doc_vigilante}")
    print(f"   • nro_solicitud: {nro_solicitud}")
    print(f"   • fecha_programacion: {fecha_programacion_valor}")
    print(f"   • ruc: {registro.get('ruc', '')}")
    print(f"   • prioridad: {registro.get('prioridad', '')}")
    print(f"   • siguiente_mismo_doc_y_fecha: {siguiente_mismo_doc_y_fecha}")
    print(f"   • tipo_arma (excel): {tipos_arma_excel}")
    print(f"   • arma (excel): {armas_excel}")
    print(f"   • objetivos_arma: {objetivos_arma}")
    print(f"   • tipos_arma_objetivo: {tipos_arma_objetivo}")
    return registro


def registrar_sin_cupo_en_excel(ruta_excel: str, registro: dict, observacion: str):
    """Registra observación de sin cupo en Excel sin modificar el estado actual."""
    if pd is None:
        return
    if not ruta_excel or not os.path.exists(ruta_excel):
        return

    try:
        df = pd.read_excel(ruta_excel, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]

        col_obs = "observaciones" if "observaciones" in df.columns else (
            "observacion" if "observacion" in df.columns else "observaciones"
        )
        if col_obs not in df.columns:
            df[col_obs] = ""

        idx = registro.get("_excel_index", None)
        actualizado = False

        if idx is not None and idx in df.index:
            df.loc[idx, col_obs] = observacion
            actualizado = True
        else:
            # Fallback por coincidencia de campos claves.
            sede = str(registro.get("sede", "")).strip()
            fecha = str(registro.get("fecha", "")).strip()
            hora = str(registro.get("hora_rango", "")).strip()
            nro = str(registro.get("nro_solicitud", "")).strip()

            def col_norm(nombre_col: str):
                if nombre_col in df.columns:
                    return df[nombre_col].fillna("").astype(str).str.strip()
                return pd.Series([""] * len(df), index=df.index)

            mask = (
                (col_norm("sede") == sede) &
                (col_norm("fecha") == fecha) &
                (col_norm("hora_rango") == hora) &
                (col_norm("nro_solicitud") == nro)
            )
            candidatos = df[mask]
            if not candidatos.empty:
                idx2 = candidatos.index[0]
                df.loc[idx2, col_obs] = observacion
                actualizado = True

        if actualizado:
            df.to_excel(ruta_excel, index=False)
            print(f"   📝 Excel actualizado: {col_obs}='{observacion}'")
        else:
            print("   ⚠️ No se pudo ubicar el registro en Excel para actualizar observación de sin cupo")
    except Exception as e:
        print(f"   ⚠️ No se pudo actualizar Excel con observación de sin cupo: {e}")


def seleccionar_en_selectonemenu(page, trigger_selector: str, panel_selector: str, label_selector: str, valor: str, nombre_campo: str):
    """Selecciona una opción PrimeFaces SelectOneMenu por data-label o texto visible."""
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


# ============================================================
# NAVEGACIÓN: CITAS → RESERVAS DE CITAS
# ============================================================

def navegar_reservas_citas(page):
    """
    El menú de SUCAMEC es un PrimeFaces PanelMenu (acordeón).
    NO usa hover — se expande haciendo CLIC en el <h3> header.

    Flujo:
      1. Clic en el <h3> de "CITAS" para expandir el panel.
      2. Esperar a que el panel interno sea visible (display:block).
      3. Clic en el <a> de "RESERVAS DE CITAS" (dispara el submit JSF).
      4. Esperar a que la nueva vista cargue.
    """
    print("\n📋 Navegando a CITAS → RESERVAS DE CITAS...")

    # 1. Esperar carga base
    try:
        page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        pass

    def vista_reservas_lista(timeout_ms: int = 3500) -> bool:
        """Confirma que ya estamos en la vista donde aparece el combo 'Cita para'."""
        try:
            page.locator("form#gestionCitasForm").wait_for(state="visible", timeout=timeout_ms)
            page.locator(SEL["tipo_cita_trigger"]).wait_for(state="visible", timeout=timeout_ms)
            return True
        except Exception:
            return False

    # FAST PATH: clic directo al item menuid=7_1 dentro del panel lateral j_idt10.
    # Es más rápido porque evita expandir manualmente el acordeón CITAS.
    url_antes = page.url
    try:
        page.locator("#j_idt10").wait_for(state="visible", timeout=4000)
        click_directo = page.evaluate(
            '''() => {
                const link = document.querySelector('#j_idt10 a[onclick*="7_1"][onclick*="menuPrincipal"]');
                if (!link) return false;
                link.click();
                return true;
            }'''
        )
        if click_directo:
            print("   ⚡ Fast-path: click directo en 'RESERVAS DE CITAS' (menuid 7_1)")
            try:
                page.wait_for_load_state("networkidle", timeout=7000)
            except Exception:
                pass
            if ("GestionCitas.xhtml" in page.url) or (page.url != url_antes) or vista_reservas_lista(5000):
                print(f"✅ Navegación completada (fast-path) → URL: {page.url}")
                return
            print("   ⚠️ Fast-path no confirmó navegación → usando flujo estándar")
    except Exception:
        pass

    # ── PASO 1: Clic en el header "CITAS" del PanelMenu ──────────────────────
    # El header es el <h3> que contiene <a href="#" tabindex="-1">CITAS</a>
    # Usamos el <a> interno como punto de clic (más preciso).
    header_citas = page.locator(
        '#j_idt11\\:menuPrincipal .ui-panelmenu-header a[tabindex="-1"]'
    ).filter(has_text="CITAS")

    try:
        header_citas.wait_for(state="visible", timeout=5000)
    except PlaywrightTimeoutError:
        raise Exception("No se encontró el header 'CITAS' en el PanelMenu")

    header_citas.click()
    print("   ✓ Clic en header 'CITAS' → expandiendo panel...")

    # ── PASO 2: Esperar a que el panel de CITAS sea visible ──────────────────
    # El panel tiene id fijo: j_idt11:menuPrincipal_7
    # PrimeFaces lo muestra quitando la clase ui-helper-hidden y poniendo display:block
    panel_citas = page.locator('#j_idt11\\:menuPrincipal_7')
    try:
        # Esperar a que el panel sea visible (PrimeFaces hace toggle de display)
        panel_citas.wait_for(state="visible", timeout=2500)
        print("   ✓ Panel CITAS desplegado")
    except PlaywrightTimeoutError:
        # En algunas versiones de PF el panel ya está en el DOM pero con display:none
        # Forzamos visibilidad vía JS como fallback
        print("   ⚠️ Panel no visible por Playwright → forzando visibilidad vía JS")
        page.evaluate("""
            const panel = document.getElementById('j_idt11:menuPrincipal_7');
            if (panel) {
                panel.classList.remove('ui-helper-hidden');
                panel.style.display = 'block';
            }
        """)
        page.wait_for_timeout(180)

    # ── PASO 3: Clic en "RESERVAS DE CITAS" ──────────────────────────────────
    # Buscamos el <a> que contiene el span con texto "RESERVAS DE CITAS"
    # dentro del panel de CITAS ya desplegado.
    reservas_link = panel_citas.locator(
        'a.ui-menuitem-link:has(span.ui-menuitem-text:text-is("RESERVAS DE CITAS"))'
    )
    try:
        reservas_link.wait_for(state="visible", timeout=2500)
    except PlaywrightTimeoutError:
        # Fallback: buscar directamente por el onclick con menuid 7_1
        print("   ⚠️ Link no visible → usando fallback por menuid 7_1")
        reservas_link = page.locator(
            'a[onclick*="7_1"][onclick*="menuPrincipal"]'
        )
        reservas_link.wait_for(state="visible", timeout=3000)

    reservas_link.click()
    print("   ✓ Clic en 'RESERVAS DE CITAS'")

    # ── PASO 4: Esperar a que la nueva vista cargue ───────────────────────────
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    if not vista_reservas_lista(6000):
        raise Exception("No se confirmó la vista de 'Reservas de Citas' tras la navegación")

    print(f"✅ Navegación completada → URL: {page.url}")


def seleccionar_tipo_cita_poligono(page):
    """
    En la vista de Gestión de Citas, abre el SelectOneMenu de tipo de cita
    y selecciona la opción "EXAMEN PARA POLÍGONO DE TIRO".
    """
    print("\n🎯 Seleccionando tipo de cita: EXAMEN PARA POLÍGONO DE TIRO...")

    # Esperar que la vista de gestión esté lista
    page.locator("form#gestionCitasForm").wait_for(state="visible", timeout=12000)

    # 1) Abrir el combo (trigger)
    trigger = page.locator(SEL["tipo_cita_trigger"])
    try:
        trigger.wait_for(state="visible", timeout=6000)
        trigger.click()
    except PlaywrightTimeoutError:
        # Fallback: clic en el label del select para abrir panel
        print("   ⚠️ Trigger no visible → usando fallback sobre label")
        label = page.locator(SEL["tipo_cita_label"])
        label.wait_for(state="visible", timeout=6000)
        label.click()

    # 2) Esperar panel de opciones
    panel = page.locator(SEL["tipo_cita_panel"])
    panel.wait_for(state="visible", timeout=6000)

    # 3) Seleccionar opción de polígono
    opcion = page.locator(SEL["tipo_cita_opcion_poligono"])
    try:
        opcion.wait_for(state="visible", timeout=4000)
    except PlaywrightTimeoutError:
        print("   ⚠️ Opción por data-label no visible → buscando por texto")
        opcion = panel.locator("li.ui-selectonemenu-item").filter(has_text="EXAMEN PARA POLÍGONO DE TIRO")
        opcion.wait_for(state="visible", timeout=4000)

    opcion.click()

    # 4) Validar que el label del combo refleje la selección
    label = page.locator(SEL["tipo_cita_label"])
    page.wait_for_timeout(250)
    texto_label = label.inner_text().strip().upper()
    if "POLÍGONO DE TIRO" not in texto_label and "POLIGONO DE TIRO" not in texto_label:
        raise Exception(f"No se confirmó la selección en el combo. Label actual: '{texto_label}'")

    print(f"   ✓ Tipo de cita seleccionado: {texto_label}")


def seleccionar_sede_y_fecha_desde_registro(page, registro: dict):
    """
    En Reserva de Cupos, selecciona Sede y Fecha según el registro del Excel.
    """
    sede = registro["sede"].strip()
    fecha = registro["fecha"].strip()

    print("\n🧭 Completando Reserva de Cupos con datos del Excel...")
    page.locator(SEL["reserva_form"]).wait_for(state="visible", timeout=15000)

    seleccionar_en_selectonemenu(
        page,
        trigger_selector=SEL["sede_trigger"],
        panel_selector=SEL["sede_panel"],
        label_selector=SEL["sede_label"],
        valor=sede,
        nombre_campo="Sede"
    )

    # Al cambiar sede, PrimeFaces suele refrescar opciones de fecha por AJAX.
    page.wait_for_timeout(700)

    seleccionar_en_selectonemenu(
        page,
        trigger_selector=SEL["fecha_trigger"],
        panel_selector=SEL["fecha_panel"],
        label_selector=SEL["fecha_label"],
        valor=fecha,
        nombre_campo="Fecha"
    )


def seleccionar_hora_con_cupo_y_avanzar(page, registro: dict):
    """
    Busca la hora del Excel en la tabla de cupos, valida cupos > 0,
    selecciona el radiobutton de la fila y presiona 'Siguiente'.
    """
    hora_objetivo = normalizar_hora_rango(registro.get("hora_rango", ""))
    if not hora_objetivo:
        raise Exception("El registro no tiene 'hora_rango' válido")

    print(f"\n🕒 Buscando hora en tabla: {hora_objetivo}")

    tabla = page.locator(SEL["tabla_programacion"])
    tabla.wait_for(state="visible", timeout=15000)

    filas = page.locator(SEL["tabla_programacion_rows"])
    total_filas = filas.count()
    if total_filas == 0:
        # Fallback para tablas PrimeFaces donde el sufijo _data no aparece en todos los entornos.
        filas = page.locator(f"{SEL['tabla_programacion']} tbody tr")
        total_filas = filas.count()
    if total_filas == 0:
        raise Exception("La tabla de programación no tiene filas para la fecha/sede seleccionadas")

    fila_objetivo = None
    cupos_objetivo = 0
    resumen = []

    def extraer_hora_rango_desde_texto(texto: str) -> str:
        t = str(texto or "").replace(".", ":")
        m = re.search(r"(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})", t)
        if m:
            ini = normalizar_hora_fragmento(m.group(1))
            fin = normalizar_hora_fragmento(m.group(2))
            return f"{ini}-{fin}"
        return normalizar_hora_rango(t)

    def extraer_cupos_desde_celdas(textos_celdas: list) -> int:
        # Busca el último texto numérico que no sea rango horario.
        for txt in reversed(textos_celdas):
            t = str(txt or "").strip()
            if not t or ":" in t:
                continue
            if re.search(r"\d+", t):
                return convertir_a_entero(t)
        return 0

    def click_boton_limpiar_obligatorio():
        try:
            boton_limpiar = page.locator(SEL["boton_limpiar"])
            boton_limpiar.wait_for(state="visible", timeout=7000)
            boton_limpiar.first.click(timeout=7000)
            page.wait_for_timeout(350)
            print("   ✓ Click en botón 'Limpiar' por falta de cupos")
        except Exception as e:
            raise SinCupoError(f"No se pudo accionar el botón 'Limpiar' tras detectar cupo 0: {e}")

    for i in range(total_filas):
        fila = filas.nth(i)
        celdas = fila.locator("td")
        total_celdas = celdas.count()
        if total_celdas == 0:
            continue

        textos_celdas = []
        for j in range(total_celdas):
            try:
                textos_celdas.append((celdas.nth(j).inner_text() or "").strip())
            except Exception:
                textos_celdas.append("")

        hora_tabla = ""
        for txt in textos_celdas:
            cand = extraer_hora_rango_desde_texto(txt)
            if cand and "-" in cand and re.search(r"\d{2}:\d{2}-\d{2}:\d{2}", cand):
                hora_tabla = cand
                break

        cupos = extraer_cupos_desde_celdas(textos_celdas)
        if hora_tabla:
            resumen.append(f"{hora_tabla} ({cupos})")

        if hora_tabla == hora_objetivo:
            fila_objetivo = fila
            cupos_objetivo = cupos
            break

    if fila_objetivo is None:
        raise Exception(
            "No se encontró la hora objetivo en la tabla. "
            f"Objetivo: '{hora_objetivo}' | Disponibles: {', '.join(resumen)}"
        )

    if cupos_objetivo <= 0:
        click_boton_limpiar_obligatorio()
        raise SinCupoError(f"La hora '{hora_objetivo}' no tiene cupos disponibles (Cupos Libres={cupos_objetivo})")

    radio_box = fila_objetivo.locator("td.ui-selection-column div.ui-radiobutton-box")
    if radio_box.count() == 0:
        raise Exception("No se encontró radiobutton en la fila de la hora objetivo")

    radio_box.first.click()
    page.wait_for_timeout(250)

    clase_radio = (radio_box.first.get_attribute("class") or "")
    aria_fila = (fila_objetivo.get_attribute("aria-selected") or "").lower()
    if "ui-state-active" not in clase_radio and aria_fila != "true":
        raise Exception("No se confirmó la selección del radiobutton de la hora")

    print(f"   ✓ Hora seleccionada: {hora_objetivo} (Cupos Libres={cupos_objetivo})")

    boton_siguiente = page.locator(SEL["boton_siguiente"])
    boton_siguiente.wait_for(state="visible", timeout=7000)
    boton_siguiente.click()
    print("   ✓ Click en botón 'Siguiente'")


def seleccionar_opcion_flexible_en_panel(page, panel_selector: str, texto_objetivo: str, nombre_campo: str):
    """Selecciona un li dentro de un panel PrimeFaces por coincidencia flexible de texto."""
    panel = page.locator(panel_selector)
    panel.wait_for(state="visible", timeout=7000)

    items = panel.locator("li.ui-selectonemenu-item")
    total = items.count()
    if total == 0:
        raise Exception(f"No hay opciones disponibles en {nombre_campo}")

    objetivo_norm = normalizar_texto_comparable(texto_objetivo)
    for i in range(total):
        item = items.nth(i)
        label = (item.get_attribute("data-label") or item.inner_text() or "").strip()
        label_norm = normalizar_texto_comparable(label)
        if objetivo_norm == label_norm or objetivo_norm in label_norm or label_norm in objetivo_norm:
            item.click()
            return label

    opciones = []
    for i in range(total):
        item = items.nth(i)
        opciones.append((item.get_attribute("data-label") or item.inner_text() or "").strip())
    raise Exception(
        f"No se encontró coincidencia para {nombre_campo}. "
        f"Objetivo: '{texto_objetivo}' | Opciones: {opciones}"
    )


def completar_paso_2_desde_registro(page, registro: dict):
    """
    Paso 2: tipo operación, doc. vigilante (autocomplete), seleccionar SI,
    y elegir número de solicitud por coincidencia con nro_solicitud del Excel.
    """
    tipo_operacion = registro.get("tipo_operacion", "").strip()
    doc_vigilante = registro.get("doc_vigilante", "").strip()
    nro_solicitud_excel = registro.get("nro_solicitud", "").strip()
    token_solicitud = extraer_token_solicitud(nro_solicitud_excel)

    print("\n🧩 Completando Paso 2 con datos del Excel...")

    # 2.1 Tipo de operación
    page.locator(SEL["tipo_operacion_trigger"]).wait_for(state="visible", timeout=12000)
    page.locator(SEL["tipo_operacion_trigger"]).click()
    page.locator(SEL["tipo_operacion_panel"]).wait_for(state="visible", timeout=7000)

    opcion_tipo = None
    items_tipo = page.locator(SEL["tipo_operacion_items"])
    total_tipo = items_tipo.count()
    objetivo_tipo = normalizar_texto_comparable(tipo_operacion)
    for i in range(total_tipo):
        item = items_tipo.nth(i)
        label = (item.get_attribute("data-label") or item.inner_text() or "").strip()
        label_norm = normalizar_texto_comparable(label)
        if objetivo_tipo == label_norm or objetivo_tipo in label_norm or label_norm in objetivo_tipo:
            item.click()
            opcion_tipo = label
            break
    if not opcion_tipo:
        raise Exception(f"No se encontró Tipo Operación '{tipo_operacion}' en el combo")

    page.wait_for_timeout(250)
    label_tipo = page.locator(SEL["tipo_operacion_label"]).inner_text().strip()
    if not label_tipo or label_tipo == "---":
        raise Exception("No se confirmó la selección de Tipo Operación")
    print(f"   ✓ Tipo Operación seleccionado: {opcion_tipo}")

    es_inicial = (
        "INICIAL" in normalizar_texto_comparable(label_tipo)
        or "INICIAL" in normalizar_texto_comparable(tipo_operacion)
    )

    def seleccionar_doc_vigilante_autocomplete():
        doc_input = page.locator(SEL["doc_vig_input"])
        doc_input.wait_for(state="visible", timeout=12000)
        doc_input.click()
        doc_input.fill("")
        doc_input.type(doc_vigilante, delay=20)

        panel_doc = page.locator(SEL["doc_vig_panel"])
        items_doc = page.locator(SEL["doc_vig_items"])

        elegido = False
        try:
            panel_doc.wait_for(state="visible", timeout=2500)
        except PlaywrightTimeoutError:
            # Fallback: algunos autocompletes solo abren panel si se navega por teclado.
            doc_input.press("ArrowDown")
            page.wait_for_timeout(350)

        if panel_doc.is_visible():
            try:
                items_doc.first.wait_for(state="visible", timeout=2500)
            except PlaywrightTimeoutError:
                page.wait_for_timeout(700)

            total_doc = items_doc.count()
            for i in range(total_doc):
                item = items_doc.nth(i)
                data_label = (item.get_attribute("data-item-label") or "").strip()
                data_value = (item.get_attribute("data-item-value") or "").strip()
                texto_item = item.inner_text().strip()
                if doc_vigilante in data_label or doc_vigilante in data_value or doc_vigilante in texto_item:
                    item.click()
                    elegido = True
                    break

            if not elegido and total_doc > 0:
                items_doc.first.click()
                elegido = True

        if not elegido:
            # Fallback final: forzar blur/change por si el valor exacto ya es aceptado por JSF.
            doc_input.evaluate(
                'el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); el.blur(); }'
            )

        page.wait_for_timeout(300)
        valor_doc = doc_input.input_value().strip()
        if doc_vigilante not in valor_doc:
            raise Exception(f"No se confirmó el documento vigilante. Esperado contiene '{doc_vigilante}' | Actual '{valor_doc}'")
        print(f"   ✓ Documento vigilante seleccionado: {valor_doc}")

    # 2.2 Flujo especial solo para INICIAL DE LICENCIA DE USO:
    # Tipo de Licencia -> Documento Vigilante.
    if es_inicial:
        print("   ℹ️ Flujo INICIAL detectado: primero Tipo de Licencia, luego Documento Vigilante")

        trigger_tramite = page.locator(SEL["tipo_tramite_trigger"])
        label_tramite = page.locator(SEL["tipo_tramite_label"])

        # Tras elegir Tipo Operación, JSF puede tardar en habilitar tipoTramite.
        habilitado = False
        for _ in range(8):
            try:
                trigger_tramite.wait_for(state="visible", timeout=2000)
                label_tramite.wait_for(state="visible", timeout=2000)
                habilitado = True
                break
            except Exception:
                page.wait_for_timeout(400)

        if not habilitado:
            raise Exception("No apareció el desplegable 'Tipo de Licencia' para flujo INICIAL")

        trigger_tramite.click()
        page.locator(SEL["tipo_tramite_panel"]).wait_for(state="visible", timeout=7000)

        opcion_tramite = page.locator(SEL["tipo_tramite_seg_priv"])
        try:
            opcion_tramite.wait_for(state="visible", timeout=2500)
            opcion_tramite.first.click()
        except PlaywrightTimeoutError:
            seleccionar_opcion_flexible_en_panel(
                page,
                panel_selector=SEL["tipo_tramite_panel"],
                texto_objetivo="SEGURIDAD PRIVADA",
                nombre_campo="Tipo de Licencia"
            )

        page.wait_for_timeout(350)
        texto_tramite = page.locator(SEL["tipo_tramite_label"]).inner_text().strip()
        if normalizar_texto_comparable(texto_tramite) != "SEGURIDAD PRIVADA":
            raise Exception(f"No se confirmó Tipo de Licencia = SEGURIDAD PRIVADA. Actual: '{texto_tramite}'")
        print("   ✓ Tipo de Licencia: SEGURIDAD PRIVADA")

        # Con Tipo de Licencia ya seteado, recien se habilita/visibiliza el DNI.
        seleccionar_doc_vigilante_autocomplete()
    else:
        # Flujo RENOVACION (u otros): Documento Vigilante directo.
        seleccionar_doc_vigilante_autocomplete()

    # 2.3 Seleccione Solicitud -> SI (siempre)
    page.locator(SEL["seleccione_solicitud_trigger"]).wait_for(state="visible", timeout=12000)
    page.locator(SEL["seleccione_solicitud_trigger"]).click()
    page.locator(SEL["seleccione_solicitud_panel"]).wait_for(state="visible", timeout=7000)
    page.locator(SEL["seleccione_solicitud_si"]).first.click()
    page.wait_for_timeout(350)
    label_si = page.locator(SEL["seleccione_solicitud_label"]).inner_text().strip().upper()
    if label_si.replace(" ", "") != "SI":
        raise Exception(f"No se confirmó Seleccione Solicitud = SI. Actual: '{label_si}'")
    print("   ✓ Seleccione Solicitud: SI")

    if es_inicial:
        print("   ℹ️ Flujo INICIAL: también se seleccionará Nro Solicitud")

    # 2.4 Nro Solicitud por coincidencia parcial (ej. 90086)
    if not token_solicitud:
        raise Exception(f"No se pudo extraer token numérico de nro_solicitud: '{nro_solicitud_excel}'")

    page.locator(SEL["nro_solicitud_trigger"]).wait_for(state="visible", timeout=12000)
    page.locator(SEL["nro_solicitud_trigger"]).click()

    panel_nro = page.locator(SEL["nro_solicitud_panel"])
    panel_nro.wait_for(state="visible", timeout=7000)
    items_nro = page.locator(SEL["nro_solicitud_items"])
    total_nro = items_nro.count()
    if total_nro == 0:
        raise Exception("No hay opciones en el combo de Nro Solicitud")

    seleccionado_label = None
    for i in range(total_nro):
        item = items_nro.nth(i)
        label = (item.get_attribute("data-label") or item.inner_text() or "").strip()
        # Comparamos contra todos los bloques numéricos del label para encontrar el Nro Empoce
        bloques = re.findall(r"\d+", label)
        bloques_norm = [b.lstrip("0") or "0" for b in bloques]
        if token_solicitud in bloques_norm:
            item.click()
            seleccionado_label = label
            break

    if not seleccionado_label:
        disponibles = []
        for i in range(total_nro):
            item = items_nro.nth(i)
            disponibles.append((item.get_attribute("data-label") or item.inner_text() or "").strip())
        raise Exception(
            f"No se encontró Nro Solicitud con token '{token_solicitud}'. Opciones: {disponibles}"
        )

    page.wait_for_timeout(300)
    label_nro = page.locator(SEL["nro_solicitud_label"]).inner_text().strip()
    bloques_final = [b.lstrip("0") or "0" for b in re.findall(r"\d+", label_nro)]
    if token_solicitud not in bloques_final:
        raise Exception(
            f"No se confirmó Nro Solicitud. Esperado token '{token_solicitud}' | Actual '{label_nro}'"
        )
    print(f"   ✓ Nro Solicitud seleccionado: {label_nro}")


def completar_tabla_tipos_arma_y_avanzar(page, registro: dict):
    """
    En Fase 2 completa la tabla dtTipoLic según tipo_arma del Excel y
    pulsa 'Siguiente' (botonSiguiente3).

    Reglas:
      - Si hay más de un registro del mismo usuario+fecha, se infiere misma programación
        y se aplican todos los tipos/armas encontrados.
      - Si hay solo un registro, se aplica solo ese.
    """
    print("\n🔫 Completando tabla de tipos de arma (Fase 2)...")

    objetivos_excel = registro.get("objetivos_arma", []) or []
    objetivos = []
    for item in objetivos_excel:
        if isinstance(item, (list, tuple)) and len(item) == 2:
            tipo_fila = normalizar_tipo_arma_excel(item[0])
            arma_objetivo = normalizar_tipo_arma_excel(item[1])
            if tipo_fila and arma_objetivo and (tipo_fila, arma_objetivo) not in objetivos:
                objetivos.append((tipo_fila, arma_objetivo))

    if not objetivos:
        raise Exception("No se recibieron objetivos de arma válidos desde Excel (tipo_arma + arma)")

    # PrimeFaces puede renderizar filas sin el sufijo _data y en modo editable por celda.
    filas = page.locator('#tabGestion\\:creaCitaPolJurForm\\:dtTipoLic tbody tr')
    try:
        filas.first.wait_for(state="visible", timeout=9000)
    except PlaywrightTimeoutError:
        filas = page.locator('table[id^="tabGestion:creaCitaPolJurForm:dtTipoLic"] tbody tr')
        try:
            filas.first.wait_for(state="visible", timeout=4000)
        except PlaywrightTimeoutError:
            raise Exception("No se encontró la tabla de tipos de arma (dtTipoLic)")

    total_filas = filas.count()
    if total_filas == 0:
        raise Exception("La tabla dtTipoLic no tiene filas")

    aplicados = []
    for tipo_fila, arma_objetivo in objetivos:
        fila_match = None
        for i in range(total_filas):
            fila = filas.nth(i)
            celdas = fila.locator('td[role="gridcell"]')
            if celdas.count() == 0:
                celdas = fila.locator("td")

            textos = []
            for j in range(celdas.count()):
                texto_celda = normalizar_texto_comparable(celdas.nth(j).inner_text().strip())
                if texto_celda:
                    textos.append(texto_celda)

            tipo_texto = " ".join(textos)
            if tipo_fila in tipo_texto:
                fila_match = fila
                break

        if fila_match is None:
            raise Exception(f"No se encontró fila para tipo de arma '{tipo_fila}' en dtTipoLic")

        # La columna "Arma" es editable; activamos la celda para mostrar el select.
        celdas_editables = fila_match.locator("td.ui-editable-column")
        if celdas_editables.count() > 0:
            celdas_editables.last.click()
            page.wait_for_timeout(180)

        combo = fila_match.locator("select")
        if combo.count() == 0:
            raise Exception(f"No se encontró combo de Arma para tipo '{tipo_fila}'")

        combo.first.wait_for(state="visible", timeout=3500)
        combo.first.select_option(label=arma_objetivo)
        page.wait_for_timeout(350)

        try:
            page.wait_for_load_state("networkidle", timeout=3500)
        except Exception:
            pass

        seleccionado = combo.first.evaluate(
            "el => el.options[el.selectedIndex] ? el.options[el.selectedIndex].text.trim() : ''"
        )
        if normalizar_texto_comparable(seleccionado) != normalizar_texto_comparable(arma_objetivo):
            raise Exception(
                f"No se confirmó Arma para '{tipo_fila}'. Esperado '{arma_objetivo}' | Actual '{seleccionado}'"
            )

        aplicados.append(f"{tipo_fila} -> {seleccionado}")
        print(f"   ✓ {tipo_fila}: {seleccionado}")

    if not aplicados:
        raise Exception("No se aplicó ninguna selección de arma en dtTipoLic")

    boton_siguiente_3 = page.locator('#tabGestion\\:creaCitaPolJurForm\\:botonSiguiente3')
    boton_siguiente_3.wait_for(state="visible", timeout=8000)
    boton_siguiente_3.click()
    print("   ✓ Click en botón 'Siguiente' de Fase 2 (botonSiguiente3)")

    try:
        page.wait_for_load_state("networkidle", timeout=7000)
    except Exception:
        pass


# ============================================================
# FLUJO PRINCIPAL
# ============================================================

def llenar_login_sel():
    print("🚀 INICIANDO SCRIPT SEL - Login Automático")

    inicio_total_flujo = time.time()
    duracion_total_flujo = None

    playwright = sync_playwright().start()
    browser = None
    login_exitoso = False
    total_ok = 0
    total_sin_cupo = 0
    total_error = 0

    def validar_credenciales_configuradas(credenciales: dict, etiqueta: str):
        faltantes = []
        if not str(credenciales.get("numero_documento", "")).strip():
            faltantes.append("numero_documento")
        if not str(credenciales.get("usuario", "")).strip():
            faltantes.append("usuario")
        if not str(credenciales.get("contrasena", "")).strip():
            faltantes.append("contrasena")
        if faltantes:
            raise Exception(
                f"Faltan credenciales para grupo {etiqueta}: {faltantes}. "
                "Configúralas en .env"
            )

    try:
        trabajos_pendientes = obtener_trabajos_pendientes_excel(EXCEL_PATH)
        if not trabajos_pendientes:
            raise Exception("No hay registros Pendiente para procesar")

        print(f"\n📚 Registros pendientes a procesar: {len(trabajos_pendientes)}")

        grupos_ordenados = ["SELVA", "JV", "OTRO"]
        trabajos_por_grupo = {g: [] for g in grupos_ordenados}
        for trabajo in trabajos_pendientes:
            grupo = trabajo.get("ruc_grupo", "OTRO")
            if grupo not in trabajos_por_grupo:
                grupo = "OTRO"
            trabajos_por_grupo[grupo].append(trabajo)

        for grupo_ruc in grupos_ordenados:
            trabajos_grupo = trabajos_por_grupo.get(grupo_ruc, [])
            if not trabajos_grupo:
                continue

            credenciales_grupo = resolver_credenciales_por_grupo_ruc(grupo_ruc)
            validar_credenciales_configuradas(credenciales_grupo, grupo_ruc)

            print(f"\n🏢 Procesando grupo RUC {grupo_ruc} - Registros: {len(trabajos_grupo)}")
            grupo_procesado = False

            for intento_global in range(3):
                start_time = time.time()
                print(f"\n🔄 Intento login {intento_global+1}/3 para grupo {grupo_ruc}")

                if browser is not None:
                    try:
                        browser.close()
                    except Exception:
                        pass

                browser = playwright.chromium.launch(
                    headless=False,
                    slow_mo=0,
                    args=[
                        "--start-maximized",
                        "--disable-infobars",
                        "--window-size=1920,1080",
                        "--window-position=0,0"
                    ]
                )
                context = browser.new_context(viewport=None, ignore_https_errors=True)
                page = context.new_page()
                page.evaluate("() => { window.moveTo(0, 0); window.resizeTo(screen.width, screen.height); }")

                try:
                    page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=45000)
                    esperar_hasta_servicio_disponible(page, URL_LOGIN, espera_segundos=8)
                    print("1. Página de login cargada")

                    tab = page.locator(SEL["tab_tradicional"])
                    tab.wait_for(state="visible", timeout=8000)
                    tab.click()
                    print("2. Pestaña 'Autenticación Tradicional' seleccionada")

                    page.locator(SEL["numero_documento"]).wait_for(state="visible", timeout=8000)

                    page.select_option(SEL["tipo_doc_select"], value=credenciales_grupo["tipo_documento_valor"])
                    page.wait_for_timeout(450)
                    page.locator(SEL["numero_documento"]).wait_for(state="visible", timeout=8000)
                    escribir_input_jsf(page, SEL["numero_documento"], credenciales_grupo["numero_documento"])
                    escribir_input_rapido(page, SEL["usuario"], credenciales_grupo["usuario"])
                    escribir_input_rapido(page, SEL["clave"], credenciales_grupo["contrasena"])
                    print(f"✅ Credenciales llenadas para grupo {grupo_ruc}")

                    captcha_text = solve_captcha_ocr(page)
                    if captcha_text and len(captcha_text) == 5:
                        escribir_input_rapido(page, SEL["captcha_input"], captcha_text)
                        print(f"✅ CAPTCHA automático: {captcha_text}")
                    else:
                        solve_captcha_manual(page)

                    print("🔘 Enviando login...")
                    page.locator(SEL["ingresar"]).click(timeout=10000)

                    print("⏳ Validando acceso...")
                    url_ok, mensaje_error, tiempo_espera = validar_resultado_login_por_ui(page, timeout_ms=3000)

                    if not url_ok:
                        print("❌ Login falló - no se detectó sesión autenticada")
                        print(f"   → URL actual: {page.url}")
                        if mensaje_error:
                            print(f"   → Error detectado: {mensaje_error}")
                        print(f"   ⏱️ Tiempo validación: {tiempo_espera:.2f} segundos")
                        raise Exception("CAPTCHA incorrecto o credenciales inválidas")

                    total_time = time.time() - start_time
                    print("🎉 ¡ACCESO EXITOSO!")
                    print(f"   → URL: {page.url}")
                    print(f"⏱️ Tiempo total login: {total_time:.2f} segundos")
                    login_exitoso = True

                    navegar_reservas_citas(page)
                    seleccionar_tipo_cita_poligono(page)

                    for n, trabajo in enumerate(trabajos_grupo, start=1):
                        idx_excel = trabajo["idx_excel"]
                        print(
                            f"\n━━━━━━━━ {grupo_ruc} Registro {n}/{len(trabajos_grupo)} "
                            f"(idx={idx_excel}, prioridad={trabajo.get('prioridad', 'Normal')}) ━━━━━━━━"
                        )

                        esperar_hasta_servicio_disponible(page, page.url, espera_segundos=8)

                        registro_excel = cargar_primer_registro_pendiente_desde_excel(
                            EXCEL_PATH,
                            indice_excel_objetivo=idx_excel,
                        )

                        try:
                            try:
                                page.locator(SEL["reserva_form"]).wait_for(state="visible", timeout=2500)
                            except Exception:
                                seleccionar_tipo_cita_poligono(page)

                            seleccionar_sede_y_fecha_desde_registro(page, registro_excel)
                            seleccionar_hora_con_cupo_y_avanzar(page, registro_excel)
                            completar_paso_2_desde_registro(page, registro_excel)
                            completar_tabla_tipos_arma_y_avanzar(page, registro_excel)
                            completar_fase_3_resumen(page)

                            limpiar_para_siguiente_registro(page, motivo="fin de flujo")
                            total_ok += 1

                        except SinCupoError as e:
                            total_sin_cupo += 1
                            print(f"⛔ Sin cupo en este registro: {e}")
                            registrar_sin_cupo_en_excel(
                                EXCEL_PATH,
                                registro_excel,
                                f"No alcanzo cupo para horario {registro_excel.get('hora_rango', '')}"
                            )
                            continue

                        except Exception as e:
                            total_error += 1
                            print(f"❌ Error en registro idx={idx_excel}: {e}")

                            error_txt = str(e or "")
                            if "No se encontró la hora objetivo en la tabla" in error_txt:
                                registrar_sin_cupo_en_excel(
                                    EXCEL_PATH,
                                    registro_excel,
                                    (
                                        "Horario no figura en la tabla de cupos: "
                                        f"{registro_excel.get('hora_rango', '')}"
                                    ),
                                )

                            if "documento vigilante" in error_txt.lower():
                                registrar_sin_cupo_en_excel(
                                    EXCEL_PATH,
                                    registro_excel,
                                    (
                                        "Documento vigilante no disponible para esta razón social/RUC. "
                                        f"DNI={registro_excel.get('doc_vigilante', '')} | "
                                        f"RUC={registro_excel.get('ruc', '')}"
                                    ),
                                )

                            try:
                                limpiar_para_siguiente_registro(page, motivo="recuperación por error")
                            except Exception:
                                pass
                            continue

                    grupo_procesado = True
                    break

                except Exception as e:
                    print(f"❌ Intento {intento_global+1} para grupo {grupo_ruc} falló: {e}")
                    if intento_global < 2:
                        print("   Reintentando...")
                        time.sleep(1)
                    else:
                        print("   Se agotaron los 3 intentos para este grupo")

            if not grupo_procesado:
                total_error += len(trabajos_grupo)
                print(
                    f"⚠️ No se pudo procesar el grupo {grupo_ruc}. "
                    f"Se contabilizan {len(trabajos_grupo)} registros con error."
                )

        duracion_total_flujo = time.time() - inicio_total_flujo
        print(f"\n⏱️ Tiempo total del flujo: {duracion_total_flujo:.2f} segundos")
        print(f"📊 Resumen: OK={total_ok} | SIN_CUPO={total_sin_cupo} | ERROR={total_error}")

        if login_exitoso:
            print("\n✅ Flujo completado. Navegador abierto para uso manual.")
            if duracion_total_flujo is not None:
                print(f"   ⏱️ Duración final del flujo: {duracion_total_flujo:.2f} segundos")
            print("   Presiona Ctrl+C o cierra la ventana cuando termines.")
            try:
                while True:
                    time.sleep(60)
            except KeyboardInterrupt:
                print("\n🛑 Interrupción manual. Cerrando navegador...")
        else:
            print("\n❌ No se pudo completar el login después de todos los intentos.")
            input("   Presiona ENTER para cerrar el navegador...")

    finally:
        try:
            if browser is not None:
                browser.close()
        except Exception:
            pass
        try:
            playwright.stop()
        except Exception:
            pass
        print("Navegador cerrado.")


if __name__ == "__main__":
    llenar_login_sel()