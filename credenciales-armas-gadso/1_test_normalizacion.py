"""Script para normalizar el Excel sin ejecutar validaciones web."""
from dotenv import load_dotenv
import pandas as pd
import os
import re
import unicodedata

load_dotenv()

EXCEL_DESNORMALIZADO = os.path.join("data", "credenciales-desnormalizado.xlsx")
EXCEL_NORMALIZADO = os.path.join("data", "credenciales-normalizado.xlsx")


def quitar_tildes(texto: str) -> str:
    t = str(texto or "")
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def normalizar_nombre(texto: str) -> str:
    t = quitar_tildes(texto)
    t = re.sub(r"\s+", " ", t).strip().upper()
    return t


def limpiar_tildes_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Quita tildes de todas las columnas de texto para mantener salida ASCII."""
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(lambda v: quitar_tildes(v) if pd.notna(v) else v)
    return df


def normalizar_nro_documento(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if re.fullmatch(r"\d+(\.0+)?", texto):
        solo = texto.split(".")[0]
    else:
        solo = re.sub(r"\D", "", texto)
    if not solo:
        return ""
    if len(solo) <= 8:
        return solo.zfill(8)
    return solo


def normalizar_id(valor) -> str:
    s = str(valor or "").strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s


def construir_clave_registro(row: pd.Series) -> str:
    """Clave estable para preservar ID cuando no hay documento."""
    nro = normalizar_nro_documento(row.get("nro_documento", ""))
    if nro:
        return f"DOC|{nro}"

    ap_pat = normalizar_nombre(row.get("apelido paterno", ""))
    ap_mat = normalizar_nombre(row.get("apellido materno", ""))
    nombres = normalizar_nombre(row.get("nombres", ""))
    fecha = str(row.get("fecha", "") or "").strip()
    return f"NODOC|{ap_pat}|{ap_mat}|{nombres}|{fecha}"


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
    df_aux = df_aux.sort_values(
        by=["__sin_fecha", "__fecha_sort"],
        ascending=[True, True],
        kind="stable",
    ).reset_index(drop=True)
    return df_aux.drop(columns=["__sin_fecha", "__fecha_sort"], errors="ignore")


def normalizar_excel_test():
    """Normaliza el Excel desnormalizado preservando estado/detalle/tipo_doc existentes."""
    print("🔄 Leyendo Excel desnormalizado...")
    df = pd.read_excel(EXCEL_DESNORMALIZADO, dtype=str)
    df_normalizado = df.copy()

    print(f"📋 Columnas encontradas: {df_normalizado.columns.tolist()}")
    print(f"📊 Total de registros: {len(df_normalizado)}\n")

    # Estandarizar nombre de columna del documento
    if "nro_documento" not in df_normalizado.columns:
        if "nro_doc" in df_normalizado.columns:
            df_normalizado["nro_documento"] = df_normalizado["nro_doc"]
        elif "dni" in df_normalizado.columns:
            df_normalizado["nro_documento"] = df_normalizado["dni"]

    columnas_alias_doc = [c for c in ["nro_doc", "dni"] if c in df_normalizado.columns]
    if columnas_alias_doc:
        df_normalizado = df_normalizado.drop(columns=columnas_alias_doc)

    if "nro_documento" in df_normalizado.columns:
        df_normalizado["nro_documento"] = df_normalizado["nro_documento"].apply(normalizar_nro_documento)
        print("✅ nro_documento normalizado (preserva ceros a la izquierda)")

    # Fecha en formato dd/mm/aa. No mantener hora ni marca temporal.
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

    # Normalizar nombres: mayúsculas, sin tildes, sin dobles espacios
    nombre_cols = ["apelido paterno", "apellido materno", "nombres"]
    for col in nombre_cols:
        if col in df_normalizado.columns:
            df_normalizado[col] = df_normalizado[col].apply(normalizar_nombre)
    print("✅ Nombres normalizados (sin tildes, mayúsculas)")

    if all(col in df_normalizado.columns for col in nombre_cols):
        df_normalizado["nombre_completo"] = (
            df_normalizado["apelido paterno"].fillna("") + " "
            + df_normalizado["apellido materno"].fillna("") + " "
            + df_normalizado["nombres"].fillna("")
        ).str.replace(r"\s+", " ", regex=True).str.strip()

    # tipo_doc: preservar si ya existe y completar vacios por longitud de nro_documento
    if "nro_documento" in df_normalizado.columns:
        if "tipo_doc" not in df_normalizado.columns:
            df_normalizado["tipo_doc"] = ""
        tipo_actual = df_normalizado["tipo_doc"].fillna("").astype(str).str.strip()
        tipo_calculado = df_normalizado["nro_documento"].apply(tipo_doc_por_nro)
        df_normalizado["tipo_doc"] = tipo_actual.where(tipo_actual != "", tipo_calculado)

    # Preservar estado/detalle/tipo_doc/id del normalizado existente
    if "estado" not in df_normalizado.columns:
        df_normalizado["estado"] = ""
    if "detalle_validacion" not in df_normalizado.columns:
        df_normalizado["detalle_validacion"] = ""
    if "id" not in df_normalizado.columns:
        df_normalizado["id"] = ""

    mapa_id_por_doc = {}
    mapa_id_por_clave = {}
    ultimo_id_numerico = 0

    if os.path.exists(EXCEL_NORMALIZADO):
        try:
            df_existente = pd.read_excel(EXCEL_NORMALIZADO, dtype=str)
            if "nro_documento" not in df_existente.columns:
                if "nro_doc" in df_existente.columns:
                    df_existente["nro_documento"] = df_existente["nro_doc"]
                elif "dni" in df_existente.columns:
                    df_existente["nro_documento"] = df_existente["dni"]
            if "nro_documento" in df_existente.columns and "nro_documento" in df_normalizado.columns:
                df_existente["nro_documento"] = df_existente["nro_documento"].apply(normalizar_nro_documento)

                if "id" not in df_existente.columns:
                    df_existente["id"] = ""
                df_existente["id"] = df_existente["id"].apply(normalizar_id)

                if len(df_existente) > 0:
                    ids_numericos = pd.to_numeric(df_existente["id"], errors="coerce").dropna()
                    if len(ids_numericos) > 0:
                        ultimo_id_numerico = int(ids_numericos.max())

                df_existente["__clave_registro"] = df_existente.apply(construir_clave_registro, axis=1)
                existentes_con_doc = df_existente[df_existente["nro_documento"].astype(str).str.strip() != ""]
                if len(existentes_con_doc) > 0:
                    mapa_id_por_doc = (
                        existentes_con_doc.drop_duplicates("nro_documento")
                        .set_index("nro_documento")["id"]
                        .to_dict()
                    )

                mapa_id_por_clave = (
                    df_existente.drop_duplicates("__clave_registro")
                    .set_index("__clave_registro")["id"]
                    .to_dict()
                )

                mapa_estado = df_existente.drop_duplicates("nro_documento").set_index("nro_documento")["estado"] if "estado" in df_existente.columns else {}
                mapa_detalle = df_existente.drop_duplicates("nro_documento").set_index("nro_documento")["detalle_validacion"] if "detalle_validacion" in df_existente.columns else {}
                mapa_tipo_doc = df_existente.drop_duplicates("nro_documento").set_index("nro_documento")["tipo_doc"] if "tipo_doc" in df_existente.columns else {}
                mapa_id = {}
                if "id" in df_existente.columns and "id" in df_normalizado.columns:
                    mapa_id = df_existente.drop_duplicates("id").set_index("id")

                for idx, row in df_normalizado.iterrows():
                    nro = row.get("nro_documento", "")
                    id_actual = row.get("id", "")
                    estado_prev = str(mapa_estado.get(nro, "") or "").strip() if hasattr(mapa_estado, "get") else ""
                    detalle_prev = str(mapa_detalle.get(nro, "") or "").strip() if hasattr(mapa_detalle, "get") else ""
                    tipo_prev = str(mapa_tipo_doc.get(nro, "") or "").strip() if hasattr(mapa_tipo_doc, "get") else ""

                    if mapa_id is not None and len(mapa_id) > 0 and str(id_actual) in mapa_id.index.astype(str):
                        fila_id = mapa_id.loc[mapa_id.index.astype(str) == str(id_actual)].iloc[0]
                        estado_por_id = str(fila_id.get("estado", "") or "").strip()
                        detalle_por_id = str(fila_id.get("detalle_validacion", "") or "").strip()
                        tipo_por_id = str(fila_id.get("tipo_doc", "") or "").strip()
                        estado_prev = estado_por_id or estado_prev
                        detalle_prev = detalle_por_id or detalle_prev
                        tipo_prev = tipo_por_id or tipo_prev

                    if estado_prev:
                        df_normalizado.at[idx, "estado"] = estado_prev
                    if detalle_prev:
                        df_normalizado.at[idx, "detalle_validacion"] = detalle_prev
                    if tipo_prev and str(df_normalizado.at[idx, "tipo_doc"] or "").strip() == "":
                        df_normalizado.at[idx, "tipo_doc"] = tipo_prev

            print("✅ Estado, detalle_validacion, tipo_doc e IDs previos preservados desde el normalizado existente")
        except Exception as e:
            print(f"⚠️ No se pudo preservar estado/detalle/tipo_doc/id existentes: {e}")

    # Asignar IDs estables: conserva IDs antiguos y solo crea nuevos para registros nuevos.
    df_normalizado["__clave_registro"] = df_normalizado.apply(construir_clave_registro, axis=1)
    ids_asignados = []
    ids_en_uso = set()
    ids_existentes_actuales = df_normalizado["id"].apply(normalizar_id).tolist()

    for idx, row in df_normalizado.iterrows():
        id_actual = normalizar_id(ids_existentes_actuales[idx])
        nro = normalizar_nro_documento(row.get("nro_documento", ""))
        clave = row.get("__clave_registro", "")

        id_preservado = ""
        if nro and nro in mapa_id_por_doc:
            id_preservado = normalizar_id(mapa_id_por_doc.get(nro, ""))
        if not id_preservado and clave in mapa_id_por_clave:
            id_preservado = normalizar_id(mapa_id_por_clave.get(clave, ""))
        if not id_preservado:
            id_preservado = id_actual

        if not id_preservado or id_preservado in ids_en_uso:
            ultimo_id_numerico += 1
            id_preservado = str(ultimo_id_numerico)

        ids_asignados.append(id_preservado)
        ids_en_uso.add(id_preservado)

    df_normalizado["id"] = ids_asignados
    df_normalizado = df_normalizado.drop(columns=["__clave_registro"], errors="ignore")

    # Mover id al inicio para mantener orden consistente del Excel.
    columnas_ordenadas = ["id"] + [c for c in df_normalizado.columns if c != "id"]
    df_normalizado = df_normalizado[columnas_ordenadas]
    print("✅ IDs alineados: existentes conservados y nuevos asignados sin alterar los anteriores")

    print("\n✏️ Vista previa de datos normalizados:")
    cols_preview = [
        c for c in [
            "nro_documento",
            "tipo_doc",
            "contraseña",
            "apelido paterno",
            "apellido materno",
            "nombres",
            "fecha",
            "estado",
            "detalle_validacion",
        ]
        if c in df_normalizado.columns
    ]
    print(df_normalizado[cols_preview].head(2))

    df_normalizado = ordenar_por_fecha_asc(df_normalizado)
    df_normalizado = limpiar_tildes_dataframe(df_normalizado)
    df_normalizado.to_excel(EXCEL_NORMALIZADO, index=False)
    print(f"\n✅ Excel normalizado guardado en {EXCEL_NORMALIZADO}")
    return df_normalizado


if __name__ == "__main__":
    print("="*70)
    print("  PRUEBA DE NORMALIZACIÓN")
    print("="*70 + "\n")
    
    df_resultado = normalizar_excel_test()
    
    print("\n" + "="*70)
    print("  RESULTADO FINAL")
    print("="*70)
    print(f"Registros procesados: {len(df_resultado)}")
    print(f"Columnas del resultado: {df_resultado.columns.tolist()}")
    print("="*70)
