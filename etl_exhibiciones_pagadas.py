import re
import io
import os
import urllib.parse
import requests
import msal
import numpy as np  
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

# --- LIBRERÍAS PROPIAS Y RESOLUCIÓN DE RUTAS ---
import paths
import periodo_resolver as pr

# --- LIBRERÍAS REQUERIDAS PARA SUPABASE ---
from supabase import create_client

# --- CARGAR CREDENCIALES DESDE EL ARCHIVO EXTERNO .ENV (Solo Local) ---
load_dotenv()

# =============================================================================
# 1. CONFIGURACIÓN GLOBAL & AZURE API
# =============================================================================

TENANT_ID = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET, SUPABASE_URL, SUPABASE_KEY]):
    raise ValueError("❌ ERROR: Faltan credenciales esenciales. Verifica tu archivo .env o tus GitHub Secrets.")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

BASE_NAME_DETALLE  = "Resultado_exhibiciones_pagadas"
BASE_NAME_AGRUPADO = "Rexhibiciones_pagadas_agrupado"

# =============================================================================
# FUNCIONES AUXILIARES DE MICROSOFT GRAPH API
# =============================================================================
def obtener_token_azure():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    scopes = ["https://graph.microsoft.com/.default"]
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=scopes)
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Error de autenticación en Azure: {result.get('error_description')}")

def obtener_site_id(headers):
    url = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{paths.SHAREPOINT_SITE_NAME}"
    res = requests.get(url, headers=headers).json()
    if "id" not in res:
        raise Exception(f"No se encontró el sitio SharePoint '{paths.SHAREPOINT_SITE_NAME}'.")
    return res["id"]

def obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_carpeta):
    ruta_formateada = urllib.parse.quote(ruta_carpeta)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_formateada}:/children"
    
    response = requests.get(url, headers=headers)
    res_json = response.json()
    
    if response.status_code != 200:
        print(f"❌ Error de la API de Graph ({response.status_code}): {res_json.get('error', {}).get('message')}")
    
    return res_json.get("value", [])

def subir_archivo_a_sharepoint(headers, site_id, ruta_carpeta, nombre_archivo, dataframe):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False)
    buffer.seek(0)
    
    url_subida = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_carpeta}/{nombre_archivo}:/content"
    headers_subida = {**headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
    
    response = requests.put(url_subida, headers=headers_subida, data=buffer.getvalue())
    if response.status_code in [200, 201]:
        print(f"  ✅ SHAREPOINT: Archivo guardado con éxito -> {ruta_carpeta}/{nombre_archivo}")
    else:
        print(f"  ❌ Error al subir a SharePoint ({nombre_archivo}): {response.text}")

def limpiar_texto_llave(serie):
    return (serie.astype(str).str.replace(r'\.0$', '', regex=True)
            .str.strip().str.upper()
            .str.replace('Á','A').str.replace('É','E').str.replace('Í','I')
            .str.replace('Ó','O').str.replace('Ú','U')
            .replace('NAN', ''))

# ==============================================================================
# 3. LÓGICA DE NEGOCIO (Planeado vs Ejecutado)
# ==============================================================================

def preparar_datos_agrupado(df, col_cant_pagada, col_imp_ok, col_marca_plan, col_tipo_plan, col_causal, col_mar_contra, col_tipo_contra, col_cant_contra):
    df["CANTIDAD_PLANEADA"] = pd.to_numeric(df[col_cant_pagada], errors='coerce').fillna(0)
    df[col_imp_ok] = df[col_imp_ok].fillna("No")
    
    df['MARCA_FINAL'] = df[col_marca_plan]
    df['TIPO_FINAL']  = df[col_tipo_plan]
    df['CANTIDAD_EJECUTADA'] = df["CANTIDAD_PLANEADA"]

    causal_std = limpiar_texto_llave(df[col_causal])
    es_contra = causal_std.str.contains("CONTRAPRESTACION A LA NEGOCIACION", na=False)

    df.loc[es_contra, 'MARCA_FINAL'] = df.loc[es_contra, col_mar_contra]
    df.loc[es_contra, 'TIPO_FINAL']  = df.loc[es_contra, col_tipo_contra]
    df.loc[es_contra, 'CANTIDAD_EJECUTADA'] = pd.to_numeric(df.loc[es_contra, col_cant_contra], errors='coerce').fillna(0)
    df.loc[es_contra, col_imp_ok] = "Si"

    imp_final = df[col_imp_ok].astype(str).str.strip().str.upper()
    df.loc[imp_final == "NO", 'CANTIDAD_EJECUTADA'] = 0
    
    return df

# ==============================================================================
# 4. ORQUESTADOR CLOUD
# ==============================================================================

def ejecutar_proceso(spec: pr.PeriodoSpec, headers, site_id):
    print(f"🚀 Iniciando auditoría exhibiciones pagadas ({spec.etiqueta})...")

    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    periodo_str = f"{MESES_ESPANOL[spec.mes].capitalize()} {spec.anio}"

    archivos_exhib = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_EXHIB)
    
    # Búsqueda estricta separando Planning de la Base de Involves
    url_p = next((a["@microsoft.graph.downloadUrl"] for a in archivos_exhib if "PLANNING" in a["name"].upper() and not "BASE" in a["name"].upper() and (str(spec.mes) in a["name"] or MESES_ESPANOL[spec.mes] in a["name"].upper())), None)
    if not url_p:
        url_p = next((a["@microsoft.graph.downloadUrl"] for a in archivos_exhib if "PLANNING" in a["name"].upper() and not "BASE" in a["name"].upper()), None)

    url_i = next((a["@microsoft.graph.downloadUrl"] for a in archivos_exhib if ("BASE" in a["name"].upper() or "RESPUESTA" in a["name"].upper()) and not "PLANNING" in a["name"].upper() and (str(spec.mes) in a["name"] or MESES_ESPANOL[spec.mes] in a["name"].upper())), None)
    if not url_i:
        url_i = next((a["@microsoft.graph.downloadUrl"] for a in archivos_exhib if ("BASE" in a["name"].upper() or "RESPUESTA" in a["name"].upper()) and not "PLANNING" in a["name"].upper()), None)

    if not url_p or not url_i:
        raise FileNotFoundError(f"No se encontraron los archivos correctos de Planning o Base en {paths.RUTA_CARPETA_BASES_EXHIB}")

    nombre_p = next((a["name"] for a in archivos_exhib if a["@microsoft.graph.downloadUrl"] == url_p), "?")
    nombre_i = next((a["name"] for a in archivos_exhib if a["@microsoft.graph.downloadUrl"] == url_i), "?")
    print(f"  ℹ️  Planning seleccionado: {nombre_p}")
    print(f"  ℹ️  Encuesta/Base seleccionada: {nombre_i}")
    if str(spec.mes) not in nombre_p and MESES_ESPANOL[spec.mes] not in nombre_p.upper():
        print(f"  ⚠️  El nombre del Planning no menciona el mes/periodo esperado ({spec.etiqueta}) — "
              f"verifica que no sea el archivo de otro periodo.")

    # Lectura robusta de la cabecera del Planning buscando la fila real de los campos clave
    content_p = requests.get(url_p).content
    df_raw = pd.read_excel(io.BytesIO(content_p), header=None)
    header_row_idx = None

    for idx, row in df_raw.iterrows():
        fila_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
        if any(keyword in fila_str for keyword in ["PUNTO DE VENTA", "PDV"]) and any(keyword in fila_str for keyword in ["TIPO", "MARCA", "PAGADAS"]):
            header_row_idx = idx
            break

    if header_row_idx is None:
        # ANTES: caía en silencio a header_row_idx=0 (fila 0), lo que podía
        # leer el Planning con el encabezado equivocado sin ningún aviso —
        # las columnas "existen" (con nombres basura tipo "Unnamed: 3") pero
        # los datos reales quedan corridos una o más filas, produciendo
        # columnas numéricas vacías/NaN → CANTIDAD_PLANEADA en 0 para todo.
        print(
            f"  ⚠️  No se encontró ninguna fila de encabezado con 'PUNTO DE VENTA/PDV' + "
            f"'TIPO/MARCA/PAGADAS' en el Planning descargado. Usando fila 0 por defecto — "
            f"esto probablemente da columnas/datos desalineados. Primeras 3 filas crudas:\n"
            f"{df_raw.head(3).to_string()}"
        )
        header_row_idx = 0

    df_p = pd.read_excel(io.BytesIO(content_p), header=header_row_idx)
    df_p.columns = [str(c).strip() for c in df_p.columns]
    
    # Búsqueda exacta y flexible de columnas clave en el Planning
    col_pdv = next((c for c in df_p.columns if "PUNTO DE VENTA" in c.upper() or "PDV" in c.upper()), None)
    col_tipo = next((c for c in df_p.columns if "TIPO" in c.upper() and ("PAGADA" in c.upper() or "EXHIBICION" in c.upper())), None)
    col_marca = next((c for c in df_p.columns if "MARCA" in c.upper() and ("PAGADA" in c.upper() or "EXHIBICION" in c.upper())), None)
    # PRIORIDAD explícita (antes era un solo OR que dependía del orden de
    # columnas del Excel, no de cuál keyword era más confiable). "DIGITE" es
    # la más específica -- suele ser la pregunta real donde se captura el
    # número -- así que se prueba primero; "ADICIONALES" puede matchear
    # columnas de texto/etiqueta que no tienen ningún valor numérico.
    candidatos_cant = [c for c in df_p.columns if "DIGITE" in c.upper()]
    if not candidatos_cant:
        candidatos_cant = [c for c in df_p.columns if "ADICIONALES" in c.upper()]
    if not candidatos_cant:
        candidatos_cant = [c for c in df_p.columns
                           if "PAGADA" in c.upper() and ("NUMERO" in c.upper() or "CANTIDAD" in c.upper())]
    col_cant = candidatos_cant[0] if candidatos_cant else None
    if len(candidatos_cant) > 1:
        print(f"  \u2139\ufe0f  Varias columnas candidatas para cantidad: {candidatos_cant} \u2014 usando '{col_cant}'")

    if not all([col_pdv, col_tipo, col_marca, col_cant]):
        raise KeyError(f"❌ Faltan columnas requeridas en el Planning. Disponibles: {list(df_p.columns)}")

    print(f"  ℹ️  Header detectado en fila {header_row_idx} — col_cant='{col_cant}'")
    n_no_nulos = pd.to_numeric(df_p[col_cant], errors='coerce').notna().sum()
    n_mayor_cero = (pd.to_numeric(df_p[col_cant], errors='coerce').fillna(0) > 0).sum()
    print(f"  ℹ️  '{col_cant}': {n_no_nulos}/{len(df_p)} valores numéricos válidos, {n_mayor_cero} > 0")
    if n_mayor_cero == 0:
        print(f"  ⚠️  '{col_cant}' no tiene NINGÚN valor > 0 — CANTIDAD_PLANEADA saldrá en 0 para todo el periodo.")

    df_p["_JOIN_"] = df_p[col_pdv].astype(str).str.strip().str.upper()

    # Leemos la Base de Respuestas de Involves
    df_inv = pd.read_excel(io.BytesIO(requests.get(url_i).content))
    nombre_col_61 = df_inv.columns[61] if df_inv.shape[1] > 61 else "(no existe, hay menos de 62 columnas)"
    serie_marca_inv = df_inv.iloc[:, 61] if df_inv.shape[1] > 61 else pd.Series([""] * len(df_inv))
    print(f"  ℹ️  Marca de la encuesta tomada de la columna #61 ('{nombre_col_61}'). "
          f"Muestra de valores: {serie_marca_inv.dropna().unique()[:5].tolist()}")
    df_inv.columns = [str(c).strip() for c in df_inv.columns]

    col_imp_ok = next((c for c in df_inv.columns if "IMPLEMENTADA" in c.upper() or "ACUERDO" in c.upper()), "La Exhibicion esta implementada de acuerdo con el planning?")
    col_causal = next((c for c in df_inv.columns if "CAUSAL" in c.upper()), "Indique las causales:")
    col_tipo_contra = next((c for c in df_inv.columns if "TIPO" in c.upper() and "CONTRAPRESTACION" in c.upper()), "Seleccionar el Tipo de la exhibicion - CONTRAPRESTACIÓN")
    col_mar_contra = next((c for c in df_inv.columns if "MARCA" in c.upper() and "CONTRAPRESTACION" in c.upper()), "MARCA - CONTRAPRESTACIÓN")
    col_cant_contra = next((c for c in df_inv.columns if "CONTRAPRESTACION" in c.upper() and ("NUMERO" in c.upper() or "DIGITE" in c.upper())), "*Digite el numero de exhibiciones adicionales para este tipo. - CONTRAPRESTACIÓN")
    # El Planning (col_pdv = '*Punto de venta') identifica el PDV por NOMBRE,
    # no por ID numérico. La Encuesta trae varias columnas de PDV ('ID del
    # PDV' numérico, 'PDV' con el nombre, 'Núm. del PDV', 'Código del PDV',
    # etc.) — antes se priorizaba 'ID del PDV' (numérico) porque contenía
    # las palabras "ID" y "PDV", lo cual nunca podía matchear contra un
    # nombre y por eso el merge daba 0% de coincidencias. Se prioriza la
    # columna simple 'PDV' (nombre) en su lugar.
    col_id_inv = "PDV" if "PDV" in df_inv.columns else next(
        (c for c in df_inv.columns if "ID" in c.upper() and "PDV" in c.upper()),
        "ID del PDV" if "ID del PDV" in df_inv.columns else "PDV"
    )
    col_tipo_inv = next((c for c in df_inv.columns if "TIPO" in c.upper() and "PAGADA" in c.upper()), "Seleccionar el Tipo de la exhibicion (Pagadas)")

    # Diagnóstico completo: el Planning identifica el PDV por NOMBRE (col_pdv)
    # y la Encuesta parece identificarlo por un ID numérico (col_id_inv) — si
    # son sistemas de identificación distintos, la llave _LL_ nunca va a
    # matchear sin importar tipo/marca. Se listan TODAS las columnas de
    # ambos archivos para poder elegir la llave común correcta.
    print(f"  ℹ️  col_pdv (Planning) = '{col_pdv}' | col_id_inv (Encuesta) = '{col_id_inv}' | col_tipo_inv (Encuesta) = '{col_tipo_inv}'")
    print(f"  ℹ️  Columnas Planning ({len(df_p.columns)}): {list(df_p.columns)}")
    print(f"  ℹ️  Columnas Encuesta ({len(df_inv.columns)}): {list(df_inv.columns)}")

    df_detalle = df_p.copy()
    
    df_detalle["_LL_"] = (limpiar_texto_llave(df_detalle["_JOIN_"]) + "_" + 
                          limpiar_texto_llave(df_detalle[col_tipo]) + "_" + 
                          limpiar_texto_llave(df_detalle[col_marca]))
    
    df_inv["_LL_"] = (limpiar_texto_llave(df_inv[col_id_inv]) + "_" + 
                      limpiar_texto_llave(df_inv[col_tipo_inv]) + "_" + 
                      limpiar_texto_llave(serie_marca_inv))

    print(f"  ℹ️  Muestra _LL_ Planning: {df_detalle['_LL_'].head(3).tolist()}")
    print(f"  ℹ️  Muestra _LL_ Encuesta: {df_inv['_LL_'].head(3).tolist()}")

    # Universo real de respuestas con la sección "Pagadas" contestada (no en
    # blanco) — si la mayoría de la Encuesta reportó exhibiciones gratis o
    # de contraprestación en vez de pagadas, el techo de matches posibles
    # es bajo de por sí, independientemente de la llave de cruce.
    tipo_pagada_no_vacio = df_inv[col_tipo_inv].astype(str).str.strip().replace("nan", "").ne("")
    print(f"  ℹ️  Filas de Encuesta con '{col_tipo_inv}' contestado (no vacío): "
          f"{tipo_pagada_no_vacio.sum()}/{len(df_inv)}")

    # Muestra de PDV/tipo/marca SIN concatenar, de filas donde sí hay
    # respuesta "Pagadas" — para comparar visualmente contra el Planning y
    # detectar diferencias de formato (tildes, mayúsculas, espacios, sufijos).
    muestra_inv = df_inv.loc[tipo_pagada_no_vacio, [col_id_inv, col_tipo_inv]].head(5).copy()
    muestra_inv["MARCA"] = serie_marca_inv.loc[muestra_inv.index]
    print(f"  ℹ️  Muestra Encuesta (solo con 'Pagadas' contestado):\n{muestra_inv.to_string()}")
    print(f"  ℹ️  Muestra Planning (PDV/Tipo/Marca crudos):\n"
          f"{df_p[[col_pdv, col_tipo, col_marca]].head(5).to_string()}")

    campos_auditoria = [col_imp_ok, col_causal, col_tipo_contra, col_mar_contra, col_cant_contra]
    df_inv_sub = df_inv[["_LL_"] + [c for c in campos_auditoria if c in df_inv.columns]].drop_duplicates(subset=["_LL_"])
    
    df_final = df_detalle.merge(df_inv_sub, on="_LL_", how="left")

    n_match = df_final[col_imp_ok].notna().sum() if col_imp_ok in df_final.columns else 0
    print(f"  ℹ️  Merge Planning↔Encuesta: {n_match}/{len(df_final)} filas encontraron match "
          f"({n_match/len(df_final)*100:.1f}%)")
    if n_match == 0:
        print(f"  ⚠️  NINGUNA fila hizo match — todas quedarán con '{col_imp_ok}'='No' por defecto, "
              f"y CANTIDAD_EJECUTADA saldrá en 0 para todo el periodo. Revisa la muestra de _LL_ arriba: "
              f"si no coinciden, el problema está en col_tipo/col_marca (Planning) o en la columna #61 "
              f"(Encuesta) — probablemente la posición de esa columna cambió.")

    df_final = preparar_datos_agrupado(df_final, col_cant, col_imp_ok, col_marca, col_tipo, col_causal, col_mar_contra, col_tipo_contra, col_cant_contra)

    dim_agrupar = [col_pdv, 'TIPO_FINAL', 'MARCA_FINAL', col_imp_ok, col_causal]
    df_agrupado = df_final.groupby(dim_agrupar, dropna=False)[['CANTIDAD_PLANEADA', 'CANTIDAD_EJECUTADA']].sum().reset_index()

    df_final_clean = df_final.drop(columns=["_LL_", "_JOIN_"], errors='ignore')
    
    nombre_detalle = f"{BASE_NAME_DETALLE} {periodo_str}.xlsx"
    nombre_agrupado = f"{BASE_NAME_AGRUPADO} {periodo_str}.xlsx"

    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_EXHIB, nombre_detalle, df_final_clean)
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_EXHIB, nombre_agrupado, df_agrupado)

    print(f"✅ ¡Hecho! Archivos subidos a SharePoint en {paths.RUTA_CARPETA_SALIDAS_EXHIB}")

def _calcular_captura_planning(spec: pr.PeriodoSpec, headers, site_id) -> pd.DataFrame:
    try:
        archivos_exhib = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_EXHIB)
        url_p = next((a["@microsoft.graph.downloadUrl"] for a in archivos_exhib if "PLANNING" in a["name"].upper() and not "BASE" in a["name"].upper() and (str(spec.mes) in a["name"] or MESES_ESPANOL[spec.mes] in a["name"].upper())), None)
        url_i = next((a["@microsoft.graph.downloadUrl"] for a in archivos_exhib if ("BASE" in a["name"].upper() or "RESPUESTA" in a["name"].upper()) and not "PLANNING" in a["name"].upper() and (str(spec.mes) in a["name"] or MESES_ESPANOL[spec.mes] in a["name"].upper())), None)

        if not url_p or not url_i:
            return pd.DataFrame(columns=['MES','AÑO','EMPLEADO','CAPTURA_PLANEADA','CAPTURA_EJECUTADA','CUMPLIMIENTO_CAPTURA'])

        df_plan = pd.read_excel(io.BytesIO(requests.get(url_p).content), engine='openpyxl')
        df_enc = pd.read_excel(io.BytesIO(requests.get(url_i).content), engine='openpyxl')
    except Exception as e:
        print(f"  ⚠️ Captura planning error: {e}")
        return pd.DataFrame(columns=['MES','AÑO','EMPLEADO','CAPTURA_PLANEADA','CAPTURA_EJECUTADA','CUMPLIMIENTO_CAPTURA'])

    col_pdv_p = next((c for c in df_plan.columns if "PUNTO DE VENTA" in str(c).upper() or "PDV" in str(c).upper()), None)
    col_emp_p = next((c for c in df_plan.columns if "EMPLEADO" in str(c).upper() or "SUPERVISOR" in str(c).upper()), None)
    col_pdv_e = 'PDV' if 'PDV' in df_enc.columns else None

    if not (col_pdv_p and col_emp_p and col_pdv_e):
        return pd.DataFrame(columns=['MES','AÑO','EMPLEADO','CAPTURA_PLANEADA','CAPTURA_EJECUTADA','CUMPLIMIENTO_CAPTURA'])

    def _norm(s): return s.astype(str).str.strip().str.upper()
    df_plan['_PDV_K']  = _norm(df_plan[col_pdv_p])
    df_plan['EMPLEADO'] = _norm(df_plan[col_emp_p])
    df_enc['_PDV_K']   = _norm(df_enc[col_pdv_e])

    pdvs_capturados = set(df_enc['_PDV_K'].dropna().unique())

    asign = df_plan[['EMPLEADO', '_PDV_K']].drop_duplicates()
    asign['CAPTURA_PLANEADA']  = 1
    asign['CAPTURA_EJECUTADA'] = asign['_PDV_K'].isin(pdvs_capturados).astype(int)

    agg = (asign.groupby('EMPLEADO', as_index=False)
                .agg(CAPTURA_PLANEADA=('CAPTURA_PLANEADA', 'sum'),
                     CAPTURA_EJECUTADA=('CAPTURA_EJECUTADA', 'sum')))
    agg['CUMPLIMIENTO_CAPTURA'] = (
        agg['CAPTURA_EJECUTADA'] / agg['CAPTURA_PLANEADA']
    ).where(agg['CAPTURA_PLANEADA'] != 0, 0)
    agg['MES'] = spec.mes
    agg['AÑO'] = spec.anio
    return agg[['MES','AÑO','EMPLEADO','CAPTURA_PLANEADA','CAPTURA_EJECUTADA','CUMPLIMIENTO_CAPTURA']]

def generar_resumen_kpi_exhibiciones_pagadas(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- KPI (V3): RESUMEN EXHIBICIONES PAGADAS por empleado  ({spec.etiqueta}) ---")
    periodo_str = f"{MESES_ESPANOL[spec.mes].capitalize()} {spec.anio}"
    nombre = f"{BASE_NAME_DETALLE} {periodo_str}.xlsx"

    archivos_salida = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_EXHIB)
    url_reporte = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salida if nombre in a["name"]), None)

    if not url_reporte:
        print(f"❌ No existe en SharePoint: {nombre}")
        return

    df = pd.read_excel(io.BytesIO(requests.get(url_reporte).content), engine='openpyxl')
    print(f"  Leyendo desde nube: {nombre} ({len(df)} filas)")

    if 'CANTIDAD_PLANEADA' not in df.columns or 'CANTIDAD_EJECUTADA' not in df.columns:
        print(f"  ⚠️  El detalle no tiene columnas CANTIDAD_PLANEADA/CANTIDAD_EJECUTADA. "
              f"Columnas disponibles: {list(df.columns)}")

    df['CANTIDAD_PLANEADA']  = pd.to_numeric(df.get('CANTIDAD_PLANEADA'),  errors='coerce').fillna(0)
    df['CANTIDAD_EJECUTADA'] = pd.to_numeric(df.get('CANTIDAD_EJECUTADA'), errors='coerce').fillna(0)

    print(f"  ℹ️  CANTIDAD_PLANEADA > 0: {(df['CANTIDAD_PLANEADA'] > 0).sum()}/{len(df)} filas | "
          f"CANTIDAD_EJECUTADA > 0: {(df['CANTIDAD_EJECUTADA'] > 0).sum()}/{len(df)} filas")
    if (df['CANTIDAD_PLANEADA'] > 0).sum() == 0:
        print(f"  ⚠️  CANTIDAD_PLANEADA es 0 en TODAS las filas del detalle — el problema viene "
              f"de {nombre} (revisa ejecutar_proceso, no este paso de KPI).")

    df['PLANEADO_REC']  = (df['CANTIDAD_PLANEADA']  > 0).astype(int)
    df['EJECUTADO_REC'] = (df['CANTIDAD_EJECUTADA'] > 0).astype(int)

    col_empleado = next((c for c in df.columns if "EMPLEADO" in str(c).upper() or "SUPERVISOR" in str(c).upper()), None)
    if not col_empleado:
        print(f"❌ No se encontró la columna de empleado en el reporte procesado.")
        return

    df['MES'] = spec.mes
    df['AÑO'] = spec.anio

    resumen = (
        df.groupby(['MES', 'AÑO', col_empleado], dropna=False)
          .agg(PLANEADO=('PLANEADO_REC',  'sum'),
               EJECUTADO=('EJECUTADO_REC', 'sum'))
          .reset_index()
    )
    resumen['CUMPLIMIENTO'] = (
        resumen['EJECUTADO'] / resumen['PLANEADO']
    ).where(resumen['PLANEADO'] != 0, 0)
    resumen = resumen.rename(columns={col_empleado: 'EMPLEADO'})
    resumen['EMPLEADO'] = resumen['EMPLEADO'].astype(str).str.strip().str.upper()

    cap = _calcular_captura_planning(spec, headers, site_id)
    if not cap.empty:
        resumen = resumen.merge(cap, on=['MES','AÑO','EMPLEADO'], how='outer')
        for c in ('PLANEADO', 'EJECUTADO', 'CUMPLIMIENTO'):
            if c in resumen.columns:
                resumen[c] = resumen[c].fillna(0)
        for c in ('CAPTURA_PLANEADA', 'CAPTURA_EJECUTADA', 'CUMPLIMIENTO_CAPTURA'):
            if c in resumen.columns:
                resumen[c] = resumen[c].fillna(0)
        resumen['MES'] = resumen['MES'].fillna(spec.mes).astype(int)
        resumen['AÑO'] = resumen['AÑO'].fillna(spec.anio).astype(int)

    cols_orden = ['MES','AÑO','EMPLEADO',
                  'PLANEADO','EJECUTADO','CUMPLIMIENTO',
                  'CAPTURA_PLANEADA','CAPTURA_EJECUTADA','CUMPLIMIENTO_CAPTURA']
    cols_orden = [c for c in cols_orden if c in resumen.columns]
    resumen = resumen[cols_orden].sort_values(['AÑO', 'MES', 'EMPLEADO']).reset_index(drop=True)

    nombre_kpi = paths.EXHIB_PAG_OUT_KPIS.name
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_EXHIB, nombre_kpi, resumen)

    try:
        print(f"  🚀 Preparando cargue del KPI de Exhibiciones Pagadas a Supabase...")
        df_kpi = resumen.copy()
        df_kpi.columns = df_kpi.columns.str.strip().str.lower()
        if 'año' in df_kpi.columns:
            df_kpi = df_kpi.rename(columns={'año': 'anio'})
        
        df_kpi = df_kpi.replace([np.inf, -np.inf], 0)
        df_kpi_limpio = df_kpi.astype(object).where(pd.notnull(df_kpi), None)
        
        registros_json = df_kpi_limpio.to_dict(orient="records")
        supabase.table("exhibiciones_pagadas_kpis").upsert(registros_json).execute()
        print(f"  ✅ ÉXITO CLOUD: {len(registros_json)} filas subidas de forma exitosa a Supabase.")
    except Exception as ex_cloud:
        print(f"  ⚠ ADVERTENCIA EN CARGA CLOUD: {ex_cloud}")

# =============================================================================
# 5. EJECUCIÓN PRINCIPAL
# =============================================================================

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL Exhibiciones Pagadas — Cloud Only")
    parser.add_argument("--solo", nargs="+", choices=["full", "kpi"], help="Default: corre todo + kpi.")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    pasos = args.solo or ["full", "kpi"]
    specs = pr.periodos_de_args(args)

    print("🔒 Autenticando con Microsoft Graph API...")
    token = obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = obtener_site_id(headers)

    for i, spec in enumerate(specs, 1):
        print(f"\n🎯 ETL Exh Pagadas — procesando periodo {spec.etiqueta}")
        try:
            if "full" in pasos:
                ejecutar_proceso(spec, headers, site_id)
            if "kpi" in pasos:
                generar_resumen_kpi_exhibiciones_pagadas(spec, headers, site_id)
        except Exception as e:
            print(f"\n❌ ERROR en {spec.etiqueta}: {e}")
            raise