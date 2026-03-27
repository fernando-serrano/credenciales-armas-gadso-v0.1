from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
import os
import threading
import time
import re
import unicodedata

load_dotenv()

# ============================================================
# CONFIGURACION
# ============================================================

URL_INSCRIPCION = "https://www.sucamec.gob.pe/sel/faces/pub/inscripcionAcceso.xhtml"
EXCEL_NORMALIZADO = os.path.join("data", "credenciales-normalizado.xlsx")
DASHBOARD_VALIDACION_ACCESO = os.path.join("data", "dashboard_validacion_acceso.png")
HEADLESS_BROWSER = False
ESCRIBIR_EXCEL = True  # Ahora escribimos los cambios en el Excel

# Mapeo de resultados a estados/detalles en Excel
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
    "PENDIENTE_ACTIVACION": {
        "estado": "No Activo",
        "detalle": "Cuenta pendiente de activación (revisar correo)",
    },
}

CANCEL_EVENT = threading.Event()


def iniciar_listener_cancelacion():
    """Inicia listener para cancelar con Enter."""
    if CANCEL_EVENT.is_set():
        return

    def _esperar_enter_cancelacion():
        try:
            input("\nPresiona Enter para cancelar en cualquier momento...\n")
            CANCEL_EVENT.set()
            print("\nCancelacion solicitada.")
        except Exception:
            return

    hilo = threading.Thread(target=_esperar_enter_cancelacion, daemon=True)
    hilo.start()


def cancelacion_solicitada() -> bool:
    return CANCEL_EVENT.is_set()


def verificar_cancelacion():
    if cancelacion_solicitada():
        raise KeyboardInterrupt("Cancelado por usuario")


def es_error_pestana_cerrada(error: Exception) -> bool:
    """Detecta errores cuando usuario cierra pestaña/contexto/browser."""
    msg = str(error or "").lower()
    patrones = [
        "target page, context or browser has been closed",
        "page closed",
        "browser has been closed",
        "context has been closed",
    ]
    return any(p in msg for p in patrones)


# ============================================================
# SELECTORES DE INSCRIPCION
# ============================================================

SEL_INSCRIPCION = {
    # Selectores de tipo de documento
    "tipo_doc_label": "#formInscAcceso\\:cbTipoDoc_label",
    "tipo_doc_trigger": "#formInscAcceso\\:cbTipoDoc .ui-selectonemenu-trigger",
    "tipo_doc_panel": "#formInscAcceso\\:cbTipoDoc_panel",
    "tipo_doc_dni": "#formInscAcceso\\:cbTipoDoc_1",

    # Campos de texto
    "numero_doc": "#formInscAcceso\\:numDoc",
    "nombres": "#formInscAcceso\\:nomb",
    "apellido_paterno": "#formInscAcceso\\:appat",
    "apellido_materno": "#formInscAcceso\\:apmat",

    # Botones
    "btn_validar": "#formInscAcceso\\:btnValidar",
    "btn_salir": "button.btn-dark",

    # Mensajes de respuesta
    "msg_existe": ".ui-growl-message",
    "msg_error": ".ui-messages-error, .ui-message-error, .ui-growl-message-error",

    # Campos que se habilitan cuando la persona no tiene cuenta activa
    "genero_label": "#formInscAcceso\\:cbGenero_label",
    "genero_trigger": "#formInscAcceso\\:cbGenero .ui-selectonemenu-trigger",
}


# ============================================================
# HELPERS
# ============================================================

def normalizar_texto(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "").strip()).lower()


def limpiar_dni(valor) -> str:
    s = re.sub(r"\D", "", str(valor or ""))
    if len(s) <= 8:
        return s.zfill(8)
    return s[:9]


def limpiar_texto(valor) -> str:
    return re.sub(r"\s+", " ", str(valor or "").strip())


def normalizar_id(valor, fallback: str = "") -> str:
    """Normaliza id para evitar nulos y formatos numericos como 12.0."""
    s = str(valor or "").strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s or str(fallback or "")


def quitar_tildes(texto: str) -> str:
    t = str(texto or "")
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def limpiar_tildes_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Quita tildes de todas las columnas de texto antes de guardar en Excel."""
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(lambda v: quitar_tildes(v) if pd.notna(v) else v)
    return df


def formatear_duracion(segundos: float) -> str:
    total = int(max(0, round(segundos)))
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def guardar_progreso_excel(df, idx_registro: int):
    """Guarda progreso incremental para persistir cada registro procesado."""
    try:
        df = limpiar_tildes_dataframe(df)
        df.to_excel(EXCEL_NORMALIZADO, index=False)
        print(f"   💾 Progreso guardado en Excel tras registro {idx_registro + 1}")
    except Exception as e:
        print(f"   ⚠️ No se pudo guardar progreso incremental: {e}")


def generar_dashboard_validacion_acceso_desde_excel(ruta_excel: str = EXCEL_NORMALIZADO):
    """Genera dashboard leyendo directamente el Excel normalizado guardado en disco."""
    try:
        import matplotlib.pyplot as plt

        if not os.path.exists(ruta_excel):
            print(f"⚠️ No existe el Excel para dashboard: {ruta_excel}")
            return None

        df_excel = pd.read_excel(ruta_excel)
        if len(df_excel) == 0:
            print("⚠️ El Excel está vacío; no se puede generar dashboard")
            return None

        if "estado" not in df_excel.columns:
            df_excel["estado"] = ""
        if "detalle_validacion" not in df_excel.columns:
            df_excel["detalle_validacion"] = ""

        # Vista global del Excel
        conteo_estado = (
            df_excel["estado"]
            .fillna("SIN_ESTADO")
            .astype(str)
            .str.strip()
            .replace("", "SIN_ESTADO")
            .value_counts()
        )
        conteo_detalle = (
            df_excel["detalle_validacion"]
            .fillna("SIN_DETALLE")
            .astype(str)
            .str.strip()
            .replace("", "SIN_DETALLE")
            .value_counts()
            .head(8)
        )

        # Vista del subconjunto objetivo del pipeline 3
        filtro_objetivo = (
            (df_excel["estado"].astype(str).str.strip() == "No Activo")
            & (
                df_excel["detalle_validacion"]
                .astype(str)
                .str.contains("Error de login: usuario o clave incorrectos", case=False, na=False)
            )
        )
        df_objetivo = df_excel[filtro_objetivo].copy()
        conteo_objetivo = (
            df_objetivo["detalle_validacion"]
            .fillna("SIN_DETALLE")
            .astype(str)
            .str.strip()
            .replace("", "SIN_DETALLE")
            .value_counts()
            .head(8)
        )

        fig, axes = plt.subplots(2, 2, figsize=(16, 10))
        fig.suptitle("Dashboard - Validación de Acceso (Inscripción)", fontsize=15, fontweight="bold")

        # Estado global del Excel
        colores_estado = ["#2ca02c" if k == "Activo" else "#d62728" if k == "No Activo" else "#7f7f7f" for k in conteo_estado.index]
        axes[0, 0].bar(conteo_estado.index, conteo_estado.values, color=colores_estado)
        axes[0, 0].set_title("Estado global en Excel")
        axes[0, 0].set_ylabel("Cantidad")
        axes[0, 0].tick_params(axis="x", rotation=25)
        for i, v in enumerate(conteo_estado.values):
            axes[0, 0].text(i, v, str(v), ha="center", va="bottom", fontsize=9)

        # Pie de estado global
        axes[0, 1].pie(
            conteo_estado.values,
            labels=conteo_estado.index,
            autopct="%1.1f%%",
            startangle=90,
            textprops={"fontsize": 10},
        )
        axes[0, 1].set_title("Estado final escrito en Excel")

        # Top detalle_validacion global
        etiquetas_detalle = [d if len(d) <= 45 else d[:42] + "..." for d in conteo_detalle.index]
        axes[1, 0].barh(range(len(conteo_detalle)), conteo_detalle.values, color="#4c78a8")
        axes[1, 0].set_yticks(range(len(conteo_detalle)))
        axes[1, 0].set_yticklabels(etiquetas_detalle, fontsize=9)
        axes[1, 0].invert_yaxis()
        axes[1, 0].set_title("Top detalle_validacion (global Excel)")
        axes[1, 0].set_xlabel("Cantidad")
        for i, v in enumerate(conteo_detalle.values):
            axes[1, 0].text(v, i, f" {v}", va="center", fontsize=9)

        # Resumen numérico desde Excel
        axes[1, 1].axis("off")
        activos = int((df_excel["estado"].astype(str).str.strip() == "Activo").sum())
        no_activos = int((df_excel["estado"].astype(str).str.strip() == "No Activo").sum())
        sin_estado = int((df_excel["estado"].fillna("").astype(str).str.strip() == "").sum())
        objetivo_restante = len(df_objetivo)
        top_objetivo = "-"
        if len(conteo_objetivo) > 0:
            top_objetivo = f"{conteo_objetivo.index[0]} ({int(conteo_objetivo.iloc[0])})"
        resumen = (
            f"Total filas en Excel: {len(df_excel)}\n"
            f"Activos: {activos}\n"
            f"No Activos: {no_activos}\n"
            f"Sin estado: {sin_estado}\n"
            f"Pendientes objetivo P3: {objetivo_restante}\n"
            f"Top detalle objetivo: {top_objetivo}\n"
            f"Imagen: {DASHBOARD_VALIDACION_ACCESO}"
        )
        axes[1, 1].text(
            0.02,
            0.98,
            resumen,
            va="top",
            fontsize=10,
            family="monospace",
            bbox={"boxstyle": "round", "facecolor": "#f7f7f7", "alpha": 0.8},
        )

        plt.tight_layout(rect=[0, 0.02, 1, 0.95])
        fig.savefig(DASHBOARD_VALIDACION_ACCESO, dpi=150)
        plt.close(fig)
        print(f"✅ Dashboard guardado en: {DASHBOARD_VALIDACION_ACCESO}")
        return DASHBOARD_VALIDACION_ACCESO
    except Exception as e:
        print(f"⚠️ Error generando dashboard de validación de acceso: {e}")
        return None


def esperar_fin_ajax(page, timeout_ms: int = 3000):
    """Espera a que no haya AJAX activo de jQuery/PrimeFaces si existe."""
    inicio = time.time()
    while (time.time() - inicio) * 1000 < timeout_ms:
        if cancelacion_solicitada():
            raise KeyboardInterrupt("Cancelado por usuario")
        try:
            ajax_activo = page.evaluate("""
                () => {
                    try {
                        if (window.jQuery && typeof window.jQuery.active !== 'undefined') {
                            return window.jQuery.active > 0;
                        }
                    } catch (e) {}
                    return false;
                }
            """)
            if not ajax_activo:
                return
        except Exception:
            return
        page.wait_for_timeout(120)


def seleccionar_dni_tipo_doc(page) -> bool:
    """Compatibilidad: selecciona DNI."""
    return seleccionar_tipo_doc(page, "DNI")


def seleccionar_tipo_doc(page, tipo_doc: str = "DNI") -> bool:
    """Selecciona tipo de documento (DNI o CARNET EXTRANJERIA) en el combo."""
    try:
        tipo_doc_norm = str(tipo_doc or "DNI").strip().upper()
        print(f"   Seleccionando tipo de documento: {tipo_doc_norm}")
        trigger = page.locator(SEL_INSCRIPCION["tipo_doc_trigger"]).first
        trigger.wait_for(state="visible", timeout=5000)
        trigger.click()
        panel = page.locator(SEL_INSCRIPCION["tipo_doc_panel"]).first
        panel.wait_for(state="visible", timeout=5000)

        if "CARNET" in tipo_doc_norm:
            opcion_ce = panel.locator("li.ui-selectonemenu-item").filter(has_text="CARNET")
            opcion_ce.first.wait_for(state="visible", timeout=5000)
            opcion_ce.first.click()
        else:
            opcion_dni = page.locator(SEL_INSCRIPCION["tipo_doc_dni"]).first
            opcion_dni.wait_for(state="visible", timeout=5000)
            opcion_dni.click()

        esperar_fin_ajax(page, timeout_ms=2500)

        label = page.locator(SEL_INSCRIPCION["tipo_doc_label"]).first
        texto = (label.inner_text() or "").strip()
        texto_lower = texto.lower()
        if "CARNET" in tipo_doc_norm:
            ok = ("carnet" in texto_lower) or ("extranjer" in texto_lower)
        else:
            ok = "dni" in texto_lower

        if ok:
            print(f"   OK tipo documento: {texto}")
            return True

        print(f"   Tipo documento no confirmado. Label actual: {texto}")
        return False
    except Exception as e:
        print(f"   Error seleccionando tipo de documento: {e}")
        return False


def rellenar_campo(page, selector: str, valor: str, nombre: str) -> bool:
    """Rellena un campo y valida que quede el valor correcto."""
    try:
        campo = page.locator(selector).first
        campo.wait_for(state="visible", timeout=5000)

        valor_esperado = limpiar_texto(valor)
        for intento in range(1, 4):
            campo.click()
            page.wait_for_timeout(90)

            # Ruta rapida: fill directo suele evitar perder el primer caracter.
            try:
                campo.fill(valor_esperado)
                campo.dispatch_event("input")
                campo.dispatch_event("change")
                campo.dispatch_event("blur")
                page.wait_for_timeout(180)
                valor_real = (campo.input_value() or "").strip()
                if valor_real == valor_esperado:
                    print(f"   OK {nombre}: {valor_real}")
                    return True
            except Exception:
                pass

            # Fallback por tipeo para casos JSF que ignoran fill en primer intento.
            campo.click()
            page.wait_for_timeout(90)
            try:
                campo.press("Control+A")
            except Exception:
                pass
            try:
                campo.press("Backspace")
            except Exception:
                pass
            try:
                campo.fill("")
            except Exception:
                pass

            page.wait_for_timeout(90)
            delay = 30 if intento == 1 else 55
            campo.type(valor_esperado, delay=delay)

            try:
                campo.dispatch_event("input")
                campo.dispatch_event("change")
                campo.dispatch_event("blur")
            except Exception:
                pass

            page.wait_for_timeout(500)

            valor_real = (campo.input_value() or "").strip()
            if valor_real == valor_esperado:
                print(f"   OK {nombre}: {valor_real}")
                return True

            print(
                f"   Reintento {intento}/3 en {nombre}: esperado='{valor_esperado}' real='{valor_real}'"
            )

        print(f"   {nombre} no quedo correcto tras 3 intentos")
        return False
    except Exception as e:
        print(f"   Error al rellenar {nombre}: {e}")
        return False


def asegurar_tipo_doc_dni(page, tipo_doc: str = "DNI") -> bool:
    """Confirma que Tipo de documento se mantenga; si no, lo re-selecciona."""
    try:
        label = page.locator(SEL_INSCRIPCION["tipo_doc_label"]).first
        texto = (label.inner_text() or "").strip().lower()
        tipo_doc_norm = str(tipo_doc or "DNI").strip().upper()
        if ("CARNET" in tipo_doc_norm and ("carnet" in texto or "extranjer" in texto)) or ("CARNET" not in tipo_doc_norm and "dni" in texto):
            return True
    except Exception:
        pass
    return seleccionar_tipo_doc(page, tipo_doc=tipo_doc)


def detectar_formulario_habilitado(page) -> bool:
    """Detecta si el formulario inferior se habilito tras validar."""
    try:
        # En PrimeFaces el estado real suele reflejarse en contenedor e input oculto.
        contenedor = page.locator("#formInscAcceso\\:cbGenero").first
        trigger = page.locator(SEL_INSCRIPCION["genero_trigger"]).first
        input_hidden = page.locator("#formInscAcceso\\:cbGenero_input").first

        if contenedor.count() == 0 and trigger.count() == 0:
            return False

        cls_cont = (contenedor.get_attribute("class") or "").lower() if contenedor.count() > 0 else ""
        cls_trg = (trigger.get_attribute("class") or "").lower() if trigger.count() > 0 else ""
        aria_trg = (trigger.get_attribute("aria-disabled") or "").lower() if trigger.count() > 0 else ""
        dis_inp = (input_hidden.get_attribute("disabled") or "").lower() if input_hidden.count() > 0 else ""
        aria_inp = (input_hidden.get_attribute("aria-disabled") or "").lower() if input_hidden.count() > 0 else ""

        deshabilitado = (
            ("ui-state-disabled" in cls_cont)
            or ("ui-state-disabled" in cls_trg)
            or (aria_trg == "true")
            or (dis_inp == "disabled")
            or (aria_inp == "true")
        )
        return not deshabilitado
    except Exception:
        return False


def obtener_texto_respuesta(page) -> str:
    """Obtiene texto visible de mensajes growl o de error."""
    try:
        growls = page.locator(SEL_INSCRIPCION["msg_existe"])
        for i in range(growls.count()):
            item = growls.nth(i)
            if item.is_visible():
                texto = (item.inner_text() or "").strip()
                if texto:
                    return texto
    except Exception:
        pass

    try:
        errs = page.locator(SEL_INSCRIPCION["msg_error"])
        for i in range(errs.count()):
            item = errs.nth(i)
            if item.is_visible():
                texto = (item.inner_text() or "").strip()
                if texto:
                    return texto
    except Exception:
        pass

    return ""


def clasificar_texto_resultado(texto: str):
    """Clasifica textos de respuesta (DOM o payload AJAX).
    Retorna: (clasificacion_key, mensaje)
    """
    t = str(texto or "").strip()
    if not t:
        return None, None

    tl = t.lower()
    if "ya existe una cuenta activa" in tl:
        return "CUENTA_ACTIVA", t
    if "cuenta pendiente de activación" in tl or "cuenta pendiente de activacion" in tl:
        return "PENDIENTE_ACTIVACION", t
    if "no coincide" in tl:
        return "NO_COINCIDE", t
    if "se ha validado los datos ingresados" in tl:
        return "PUEDE_REGISTRARSE", t
    if "error" in tl or "obligatorio" in tl or "requerido" in tl:
        return None, f"Error/resultado: {t}"

    return None, f"Sin clasificación: {t}"


def es_payload_ajax_silencioso(payload_ajax: str) -> bool:
    """Detecta respuesta AJAX sin mensajes visibles, típica de datos no coincidentes."""
    payload = str(payload_ajax or "").strip()
    if not payload:
        return False

    payload_lower = payload.lower()
    if "ui-growl" in payload_lower:
        return False
    if "ya existe una cuenta activa" in payload_lower:
        return False
    if "no coincide" in payload_lower:
        return False
    if "cbgenero" in payload_lower:
        return False

    return True


def clasificar_payload_ajax(payload_ajax: str):
    """Clasifica respuesta AJAX solo si contiene senales de negocio reales."""
    p = str(payload_ajax or "").strip()
    if not p:
        return None, None

    pl = p.lower()

    if "<partial-response" in pl and "msgs:[]" in pl:
        return None, None

    for frase in [
        "ya existe una cuenta activa",
        "cuenta pendiente de activación",
        "cuenta pendiente de activacion",
        "no coincide",
        "se ha validado los datos ingresados",
        "captcha",
        "error",
        "obligatorio",
        "requerido",
    ]:
        if frase in pl:
            return clasificar_texto_resultado(p)

    return None, None


def log_diagnostico_post_validar(page):
    """Imprime un mini diagnóstico visual del estado post-validación."""
    try:
        print(f"   URL actual: {page.url}")
    except Exception:
        pass

    try:
        texto_msg = obtener_texto_respuesta(page)
        if texto_msg:
            print(f"   Mensaje visible: {texto_msg}")
    except Exception:
        pass

    try:
        print(f"   Formulario habilitado: {detectar_formulario_habilitado(page)}")
    except Exception:
        pass


def click_validar_robusto(page) -> tuple:
    """Acciona Validar de forma robusta.
    Retorna: (accion_ok: bool, payload_ajax: str)
    """
    btn = page.locator(SEL_INSCRIPCION["btn_validar"]).first
    btn.wait_for(state="visible", timeout=5000)
    btn.scroll_into_view_if_needed()
    form_habilitado_antes_click = detectar_formulario_habilitado(page)

    habilitado = False
    for _ in range(20):
        aria = (btn.get_attribute("aria-disabled") or "").strip().lower()
        dis = (btn.get_attribute("disabled") or "").strip().lower()
        cls = (btn.get_attribute("class") or "").lower()

        print(f"   Estado boton Validar -> aria-disabled={aria} disabled={dis} class={cls}")

        if aria != "true" and dis != "disabled" and "ui-state-disabled" not in cls:
            habilitado = True
            break
        page.wait_for_timeout(100)

    if not habilitado:
        print("   Boton Validar sigue deshabilitado (aria/disabled/class)")
        return False, ""

    try:
        btn.hover()
    except Exception:
        pass

    page.wait_for_timeout(150)

    try:
        btn.focus()
    except Exception:
        pass

    page.wait_for_timeout(100)

    def _hay_evidencia_accion(timeout_ms: int = 1400) -> bool:
        inicio = time.time()
        while (time.time() - inicio) * 1000 < timeout_ms:
            if cancelacion_solicitada():
                raise KeyboardInterrupt("Cancelado por usuario")
            if obtener_texto_respuesta(page):
                return True

            if (not form_habilitado_antes_click) and detectar_formulario_habilitado(page):
                return True

            try:
                contenido = page.content().lower()
                if (
                    "ya existe una cuenta activa" in contenido
                    or "no coincide" in contenido
                    or "captcha" in contenido
                    or "turno registrado" in contenido
                ):
                    return True
            except Exception:
                pass

            page.wait_for_timeout(120)
        return False

    def _extraer_post_data(request) -> str:
        """Compatibilidad Playwright: post_data puede ser atributo o método."""
        try:
            post_data_attr = getattr(request, "post_data", None)
            if callable(post_data_attr):
                return post_data_attr() or ""
            return post_data_attr or ""
        except Exception:
            return ""

    def _es_post_btn_validar(resp) -> bool:
        try:
            if resp.request.method.upper() != "POST":
                return False
            if "inscripcionAcceso.xhtml" not in resp.url:
                return False

            post_data = _extraer_post_data(resp.request)
            if not post_data:
                # Fallback: aceptar POST de la misma pantalla para no perder payload.
                return True

            return "formInscAcceso:btnValidar" in post_data
        except Exception:
            return False

    def _ejecutar_con_confirmacion(accion, nombre: str, usar_expect_response: bool = True):
        post_detectado = False
        payload_ajax = ""
        print(f"   Intentando accion: {nombre}")
        if not usar_expect_response:
            try:
                accion()
            except Exception:
                return False, ""
            if _hay_evidencia_accion(timeout_ms=900):
                print(f"   OK accion UI detectada tras {nombre}")
                return True, ""
            print(f"   Sin evidencia tras {nombre}")
            return False, ""

        try:
            with page.expect_response(
                _es_post_btn_validar,
                timeout=1100,
            ) as resp_info:
                accion()
            post_detectado = True
            try:
                payload_ajax = resp_info.value.text() or ""
            except Exception:
                payload_ajax = ""
        except Exception:
            # Evita re-click duplicado cuando la accion ya disparo evento sin respuesta capturada.
            if _hay_evidencia_accion(timeout_ms=900):
                print(f"   OK accion UI detectada tras {nombre}")
                return True, payload_ajax
            return False, payload_ajax

        if post_detectado:
            print(f"   OK POST real de btnValidar detectado tras {nombre}")
            return True, payload_ajax

        if _hay_evidencia_accion(timeout_ms=2200):
            print(f"   OK accion UI detectada tras {nombre}")
            return True, payload_ajax

        print(f"   Sin evidencia tras {nombre}")
        return False, payload_ajax

    ok, payload = _ejecutar_con_confirmacion(
        lambda: btn.click(timeout=2000),
        "click Playwright",
        usar_expect_response=False,
    )
    if ok:
        return True, payload

    ok, payload = _ejecutar_con_confirmacion(lambda: btn.click(force=True, timeout=2500), "click Playwright force")
    if ok:
        return True, payload

    span_validar = page.locator("#formInscAcceso\\:btnValidar .ui-button-text").first
    ok, payload = _ejecutar_con_confirmacion(lambda: span_validar.click(timeout=2200), "click span texto Validar")
    if ok:
        return True, payload

    try:
        ok, payload = _ejecutar_con_confirmacion(
            lambda: page.evaluate("""
                () => {
                    const btn = document.getElementById('formInscAcceso:btnValidar');
                    if (!btn) return false;
                    btn.click();
                    return true;
                }
            """),
            "click JS"
        )
        if ok:
            return True, payload
    except Exception:
        pass

    return False, ""


def validar_resultado_inscripcion_por_ui(page, formulario_habilitado_antes: bool, payload_ajax: str = "", timeout_ms: int = 3200):
    """Evalua resultado por payload, mensajes, formulario y fallback HTML."""
    try:
        if payload_ajax:
            clasif_payload = clasificar_payload_ajax(payload_ajax)
            if clasif_payload != (None, None):
                return clasif_payload

        print("   Esperando senales de validacion (mensaje/formulario)...")
        inicio_espera = time.perf_counter()
        selectores_mensaje = [
            SEL_INSCRIPCION["msg_existe"],
            SEL_INSCRIPCION["msg_error"],
            ".ui-messages-info",
            ".ui-message-info",
            ".ui-growl-message-info",
        ]

        # Fast-path: si el growl/error aparece rapido, clasificar sin esperar el loop completo.
        try:
            page.wait_for_selector(
                f"{SEL_INSCRIPCION['msg_existe']}, {SEL_INSCRIPCION['msg_error']}",
                state="visible",
                timeout=900,
            )
            txt_rapido = obtener_texto_respuesta(page)
            if txt_rapido:
                return clasificar_texto_resultado(txt_rapido)
        except Exception:
            pass

        while (time.perf_counter() - inicio_espera) * 1000 < timeout_ms:
            if cancelacion_solicitada():
                raise KeyboardInterrupt("Cancelado por usuario")
            texto_msg = obtener_texto_respuesta(page)
            if texto_msg:
                return clasificar_texto_resultado(texto_msg)

            for sel in selectores_mensaje:
                try:
                    loc = page.locator(sel)
                    total = min(loc.count(), 3)
                    for i in range(total):
                        t = (loc.nth(i).inner_text() or "").strip()
                        if t:
                            return clasificar_texto_resultado(t)
                except Exception:
                    pass

            formulario_habilitado_despues = detectar_formulario_habilitado(page)
            if formulario_habilitado_despues:
                return "NO_REGISTRADO", "No registrado en SUCAMEC (formulario habilitado)"

            try:
                contenido = page.content().lower()
                if "ya existe una cuenta activa" in contenido:
                    return "CUENTA_ACTIVA", "Ya existe una cuenta activa"
                if "cuenta pendiente de activación" in contenido or "cuenta pendiente de activacion" in contenido:
                    return "PENDIENTE_ACTIVACION", "Cuenta pendiente de activación"
                if "no coincide" in contenido:
                    return "NO_COINCIDE", "Los datos no coinciden"
            except Exception:
                pass

            page.wait_for_timeout(120)

        page.wait_for_timeout(500)

        if detectar_formulario_habilitado(page):
            return "NO_REGISTRADO", "No registrado en SUCAMEC (formulario habilitado)"

        txt_final = obtener_texto_respuesta(page)
        if txt_final:
            return clasificar_texto_resultado(txt_final)

        # Antes de clasificar como payload silencioso, priorizar señal de formulario habilitado.
        if detectar_formulario_habilitado(page):
            return "NO_REGISTRADO", "No registrado en SUCAMEC"

        if es_payload_ajax_silencioso(payload_ajax):
            return "NO_COINCIDE", "Datos no coinciden con registro en SUCAMEC"

        return None, "Sin señal concluyente"

    except Exception as e:
        return None, f"Error técnico: {str(e)[:100]}"


def validar_acceso_inscripcion(page, dni: str, nombres: str, apellido_paterno: str, apellido_materno: str, tipo_doc: str = "DNI") -> tuple:
    """Valida acceso por formulario de inscripcion. Retorna (existe_cuenta, mensaje)."""
    try:
        print(f"\nValidando acceso para: {dni}")

        formulario_habilitado_antes = detectar_formulario_habilitado(page)

        try:
            page.evaluate(
                """() => {
                    document.querySelectorAll('.ui-growl-item-container, .ui-growl-message').forEach(el => el.remove());
                }"""
            )
        except Exception:
            pass

        if not seleccionar_tipo_doc(page, tipo_doc=tipo_doc):
            return False, "No se pudo seleccionar tipo de documento"

        if not rellenar_campo(page, SEL_INSCRIPCION["numero_doc"], dni, "Numero de documento"):
            return False, "No se pudo ingresar numero de documento"
        if not rellenar_campo(page, SEL_INSCRIPCION["nombres"], nombres, "Nombres"):
            return False, "No se pudo ingresar nombres"
        if not rellenar_campo(page, SEL_INSCRIPCION["apellido_paterno"], apellido_paterno, "Apellido paterno"):
            return False, "No se pudo ingresar apellido paterno"
        if not rellenar_campo(page, SEL_INSCRIPCION["apellido_materno"], apellido_materno, "Apellido materno"):
            return False, "No se pudo ingresar apellido materno"

        # En algunos casos PrimeFaces pierde seleccion del combo al editar campos.
        if not asegurar_tipo_doc_dni(page, tipo_doc=tipo_doc):
            return False, "No se pudo confirmar tipo de documento DNI"

        esperar_fin_ajax(page, timeout_ms=3000)
        page.wait_for_timeout(300)

        print("   Haciendo click en Validar...")
        click_ok, payload_ajax = click_validar_robusto(page)
        if not click_ok:
            return None, "No se pudo accionar boton Validar"

        esperar_fin_ajax(page)

        clasificacion, mensaje = validar_resultado_inscripcion_por_ui(
            page,
            formulario_habilitado_antes,
            payload_ajax=payload_ajax,
            timeout_ms=3200,
        )

        # Reintento puntual cuando el sistema reporta perdida de tipo de documento.
        if (
            clasificacion is None
            and mensaje
            and "tipo de documento" in mensaje.lower()
            and "obligatorio" in mensaje.lower()
        ):
            print("   Reintentando validacion: se perdio Tipo de documento.")
            if asegurar_tipo_doc_dni(page, tipo_doc=tipo_doc):
                esperar_fin_ajax(page, timeout_ms=1200)
                click_ok_retry, payload_ajax_retry = click_validar_robusto(page)
                if click_ok_retry:
                    esperar_fin_ajax(page, timeout_ms=1800)
                    clasificacion, mensaje = validar_resultado_inscripcion_por_ui(
                        page,
                        formulario_habilitado_antes,
                        payload_ajax=payload_ajax_retry,
                        timeout_ms=3200,
                    )

        if clasificacion:
            print(f"   ✓ Resultado clasificado: {clasificacion}")
            return clasificacion, mensaje

        print("   ⚠️  Sin senal concluyente")
        log_diagnostico_post_validar(page)
        return None, "Sin senal concluyente"

    except Exception as e:
        return False, f"Error tecnico: {str(e)[:120]}"


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def procesar_validacion_acceso():
    """Procesa registros No Activo y valida acceso por pagina publica de inscripcion."""
    print("\n" + "=" * 70)
    print("  VALIDADOR DE ACCESO (INSCRIPCION/REGISTRO)")
    print("=" * 70)
    print(f"Modo navegador: {'VISIBLE' if not HEADLESS_BROWSER else 'HEADLESS'}")
    inicio_flujo = time.perf_counter()

    iniciar_listener_cancelacion()

    print("\nLeyendo Excel normalizado...")
    try:
        df = pd.read_excel(EXCEL_NORMALIZADO, dtype=str)
    except Exception as e:
        print(f"No se pudo leer Excel: {e}")
        return

    if "id" not in df.columns:
        print("⚠️ El Excel no tiene columna 'id'. Se generara de forma secuencial para este flujo.")
        df.insert(0, "id", [str(i) for i in range(1, len(df) + 1)])

    df["id"] = [normalizar_id(v, fallback=i) for i, v in enumerate(df["id"], start=1)]

    print(f"Excel cargado: {len(df)} registros totales")

    # Filtrar solo No Activo + error de credenciales
    df_no_activos = df[
        (df["estado"].astype(str).str.strip() == "No Activo") &
        (df["detalle_validacion"].astype(str).str.contains("Error de login: usuario o clave incorrectos", case=False, na=False))
    ].copy()
    print(f"Registros No Activos encontrados: {len(df_no_activos)}")
    
    # Filtrados específicamente para esta validación
    total_no_activos_especificos = len(df_no_activos)
    if total_no_activos_especificos == 0:
        print("No hay registros No Activo con 'Error de login: usuario o clave incorrectos' para validar.")
        print("(Otros estados No Activo con diferentes detalles no serán procesados)")
        return
    
    print(f"   → Solo se procesarán registros con detalle 'Error de login: usuario o clave incorrectos'")

    disponibles_validar = df_no_activos.copy()
    total_no_activos = len(disponibles_validar)

    if "id" in disponibles_validar.columns:
        disponibles_validar = disponibles_validar.drop_duplicates(subset=["id"], keep="first")
    else:
        subset_cols = [
            c for c in ["dni", "nombres", "apelido paterno", "apellido materno"]
            if c in disponibles_validar.columns
        ]
        if subset_cols:
            disponibles_validar = disponibles_validar.drop_duplicates(subset=subset_cols, keep="first")

    if len(disponibles_validar) != total_no_activos:
        print(f"Duplicados detectados: {total_no_activos - len(disponibles_validar)}")

    print(f"Candidatos para validar acceso: {len(disponibles_validar)}")
    if not ESCRIBIR_EXCEL:
        print("Modo solo validacion activo: no se escribira en el Excel")

    if len(disponibles_validar) == 0:
        print("No hay registros No Activos para procesar.")
        return

    contador_actualizados = 0
    contador_activos_encontrados = 0
    pestana_cerrada = False
    flujo_cancelado = False
    motivo_cancelacion = ""

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS_BROWSER)
        context = browser.new_context()
        page = context.new_page()

        try:
            for contador, (idx, row) in enumerate(disponibles_validar.iterrows(), 1):
                if cancelacion_solicitada():
                    flujo_cancelado = True
                    motivo_cancelacion = "Cancelado por usuario (Enter)"
                    print("Cancelacion detectada.")
                    break

                try:
                    inicio_registro = time.perf_counter()
                    verificar_cancelacion()

                    id_registro = normalizar_id(row.get("id", ""), fallback=idx)
                    nro_documento = limpiar_dni(row.get("nro_documento", row.get("dni", "")))
                    nombres = limpiar_texto(row.get("nombres", ""))
                    apellido_paterno = limpiar_texto(
                        row.get("apellido paterno", row.get("apelido paterno", ""))
                    )
                    apellido_materno = limpiar_texto(row.get("apellido materno", ""))

                    print(f"\n[{contador}/{len(disponibles_validar)}] ID={id_registro} | DOC={nro_documento} | {apellido_paterno}, {nombres}")

                    if len(nro_documento) not in (8, 9):
                        print("   nro_documento invalido (debe tener 8 o 9 digitos), se omite.")
                        continue

                    try:
                        page.goto(URL_INSCRIPCION, wait_until="domcontentloaded", timeout=30000)
                    except PlaywrightTimeoutError:
                        print("Timeout al cargar pagina. Reintentando...")
                        page.reload(wait_until="domcontentloaded")

                    page.wait_for_timeout(180)

                    tipo_doc_registro = str(row.get("tipo_doc", "DNI") or "DNI").strip().upper()

                    clasificacion, detalle_resultado = validar_acceso_inscripcion(
                        page, nro_documento, nombres, apellido_paterno, apellido_materno, tipo_doc=tipo_doc_registro
                    )

                    # Aplicar mapeo de resultados
                    if clasificacion in MAPEO_RESULTADOS:
                        estado_nuevo = MAPEO_RESULTADOS[clasificacion]["estado"]
                        detalle_nuevo = MAPEO_RESULTADOS[clasificacion]["detalle"]
                        
                        if ESCRIBIR_EXCEL:
                            mascara_id = df["id"].astype(str).str.strip() == str(id_registro)
                            if mascara_id.any():
                                df.loc[mascara_id, "estado"] = estado_nuevo
                                df.loc[mascara_id, "detalle_validacion"] = detalle_nuevo
                            else:
                                df.at[idx, "estado"] = estado_nuevo
                                df.at[idx, "detalle_validacion"] = detalle_nuevo
                            guardar_progreso_excel(df, idx)
                        
                        if clasificacion == "CUENTA_ACTIVA":
                            contador_activos_encontrados += 1
                            print(f"   ✅ ACTIVO POTENCIAL: {detalle_nuevo}")
                        else:
                            print(f"   ℹ️  {detalle_nuevo}")

                    else:
                        print(f"   ⚠️  Sin clasificación válida: {detalle_resultado}")
                        if ESCRIBIR_EXCEL:
                            guardar_progreso_excel(df, idx)

                    contador_actualizados += 1
                    duracion_registro = time.perf_counter() - inicio_registro
                    duracion_total = time.perf_counter() - inicio_flujo
                    print(
                        f"   Avance: procesados={contador_actualizados}/{len(disponibles_validar)} | "
                        f"activos={contador_activos_encontrados} | "
                        f"t_registro={formatear_duracion(duracion_registro)} | "
                        f"t_total={formatear_duracion(duracion_total)}"
                    )

                except KeyboardInterrupt:
                    flujo_cancelado = True
                    motivo_cancelacion = "Cancelado por usuario (Enter/Ctrl+C)"
                    print("Cancelacion detectada durante el registro. Deteniendo flujo...")
                    break
                except Exception as e:
                    if es_error_pestana_cerrada(e):
                        pestana_cerrada = True
                        print("Se cerro la pestaña durante la validacion. Flujo detenido.")
                        break
                    print(f"Error en registro ID={id_registro} DOC={nro_documento}: {e}")

        finally:
            try:
                context.close()
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass

    if ESCRIBIR_EXCEL:
        print("\n✅ Guardando cambios en Excel...")
        try:
            df = limpiar_tildes_dataframe(df)
            df.to_excel(EXCEL_NORMALIZADO, index=False)
            print(f"✅ Excel actualizado correctamente: {EXCEL_NORMALIZADO}")
            print(f"   Estados y detalles validación se han inscrito en el archivo")
        except Exception as e:
            print(f"❌ No se pudo guardar Excel: {e}")
            return
    else:
        print("\n⚠️  ESCRIBIR_EXCEL=False - cambios NO se guardaron en Excel")

    print("\n📊 Generando dashboard de validación de acceso...")
    generar_dashboard_validacion_acceso_desde_excel(EXCEL_NORMALIZADO)

    print("\n" + "=" * 70)
    print("  RESUMEN DE VALIDACION DE ACCESO")
    print("=" * 70)
    print(f"Registros No Activos totales: {len(df_no_activos)}")
    print(f"Candidatos para validar: {len(disponibles_validar)}")
    print(f"Procesados exitosamente: {contador_actualizados}")
    print(f"Cuentas activas encontradas: {contador_activos_encontrados}")
    if flujo_cancelado:
        print(f"Estado final: {motivo_cancelacion}")
    if pestana_cerrada:
        print("Estado final: Se cerro la pestaña (manejo controlado).")
    if not flujo_cancelado and not pestana_cerrada:
        print("Estado final: Flujo completado.")
    print("=" * 70)


if __name__ == "__main__":
    try:
        procesar_validacion_acceso()
    except KeyboardInterrupt:
        print("\nFlujo cancelado por usuario (Ctrl+C).")