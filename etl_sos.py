import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

import pandas as pd
import numpy as np
import os
import io
import urllib.parse
import requests
import msal
import re
import unicodedata
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

CLAVE_ARCHIVO_ENCUESTA = "Respuestas"
CLAVE_ARCHIVO_TARGET = "Colombia-sos-target"

LISTA_17_CATEGORIAS = [
    "SHAMPOO BEBE EN ADULTOS", "SHAMPOO BEBE", "JABONES SOLIDOS BEBE", 
    "JABONES LIQUIDOS BEBE", "CREMAS CORPORALES BEBE", "ASEO DEL BEBE", 
    "TOALLAS", "TAMPONES", "PROTECTORES", "PROTECCION SOLAR", 
    "JABONES SOLIDOS ADULTOS", "JABONES LIQUIDOS ADULTOS", "FACIALES", 
    "ENJUAGUES BUCALES MASIVOS", "ENJUAGUES BUCALES ESPECIALIZADOS", 
    "ENJUAGUES BUCALES", "CREMAS CORPORALES ADULTO"
]

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

def eliminar_tildes(texto):
    if pd.isna(texto): return texto
    texto = str(texto)
    s = ''.join(c for c in unicodedata.normalize('NFD', texto)
                if unicodedata.category(c) != 'Mn')
    return s.upper().strip()

# =============================================================================
# 3. FUNCIONES DE PROCESAMIENTO (PURAMENTE CLOUD)
# =============================================================================

def ejecutar_paso_1_consolidar_pt(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 1: Consolidando Plan de Trabajo ({spec.etiqueta}) ---")
    from shared_loader import COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR

    COLUMNAS_FINALES_PT = ["ID_PDV_INVOLVES", "NOMBRE_PDV", "VENTAS_PROMEDIO_MES", "ACRONIMO",
                           "CEDULA", "NOMBRE", "COD_MERCADERISTA", "ROL", "SUPERVISOR_LIDER",
                           "FUENTE", "OBSERVACION"]
    OBS_A_ROL = {
        "REPORTA GESTOR": "GESTOR",
        "REPORTA SUPERVISOR": "SUPERVISOR",
        "REPORTA GENERADOR DE DEMANDA": "GENERADOR DE DEMANDA",
    }

    archivos_en_carpeta = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_PT)
    num_mes = f"{spec.mes:02d}"                     
    nombre_mes = MESES_ESPANOL[spec.mes].upper()  

    url_directo = None
    url_ism = None

    for archivo in archivos_en_carpeta:
        nombre_archivo_upper = archivo["name"].upper()
        if "DIRECTO" in nombre_archivo_upper and (num_mes in nombre_archivo_upper or nombre_mes in nombre_archivo_upper):
            url_directo = archivo["@microsoft.graph.downloadUrl"]
        if "ISM" in nombre_archivo_upper and (num_mes in nombre_archivo_upper or nombre_mes in nombre_archivo_upper):
            url_ism = archivo["@microsoft.graph.downloadUrl"]

    if not url_directo:
        url_directo = next((a["@microsoft.graph.downloadUrl"] for a in archivos_en_carpeta if "DIRECTO" in a["name"].upper()), None)
    if not url_ism:
        url_ism = next((a["@microsoft.graph.downloadUrl"] for a in archivos_en_carpeta if "ISM" in a["name"].upper()), None)

    if not url_directo or not url_ism:
        raise FileNotFoundError(f"No se localizaron los archivos base en la carpeta cloud: {paths.RUTA_CARPETA_PT}")

    def leer_y_normalizar_cloud(url_download, hoja, fuente):
        content = requests.get(url_download).content
        with pd.ExcelFile(io.BytesIO(content)) as xls:
            df_val = pd.read_excel(xls, sheet_name="Captura de modulos")
            df_val.columns = df_val.columns.str.strip().str.upper()
            col_filtro = "ESPACIOS_FINAL" if fuente == "ISM" else "ESPACIOS"
            obs_col = next((c for c in df_val.columns if c.startswith('OBSERVAC')), None)
            if obs_col is None:
                print(f"⚠ {fuente}: hoja Captura de modulos sin OBSERVACIÓN")
                return pd.DataFrame()

            df_val_filtrado = df_val[df_val[col_filtro] == 1].copy()
            if "SUB CANAL" in df_val_filtrado.columns:
                n_antes_dc = len(df_val_filtrado)
                df_val_filtrado = df_val_filtrado[
                    df_val_filtrado["SUB CANAL"].astype(str).str.strip().str.upper() != "DISCOUNTER"
                ]
                n_dc = n_antes_dc - len(df_val_filtrado)
                if n_dc > 0:
                    print(f"  {fuente}: filtro DISCOUNTER → {n_dc} PDVs excluidos")

            pdvs_obs = (
                df_val_filtrado[["ID PDV INVOLVES", obs_col]]
                .rename(columns={obs_col: "OBSERVACION"})
            )
            pdvs_obs["OBSERVACION"] = pdvs_obs["OBSERVACION"].astype(str).str.strip().str.upper()
            pdvs_obs["ROL_ESPERADO"] = pdvs_obs["OBSERVACION"].map(OBS_A_ROL).fillna("")

            df = pd.read_excel(xls, sheet_name=hoja)
            df.columns = df.columns.str.strip().str.upper()
            df = df[df["ID PDV INVOLVES"].isin(pdvs_obs["ID PDV INVOLVES"])].copy()
            df.rename(columns={col: COLUMNAS_ESTANDAR[col] for col in df.columns if col in COLUMNAS_ESTANDAR}, inplace=True)
            df = df.merge(pdvs_obs, left_on="ID_PDV_INVOLVES", right_on="ID PDV INVOLVES", how="left")
            df["ROL"] = df["ROL"].astype(str).str.strip().str.upper()
            n_antes = len(df)
            df = df[df["ROL"] == df["ROL_ESPERADO"]].copy()
            print(f"  {fuente}: filtro OBSERVACIÓN → {n_antes} → {len(df)} filas (responsables reales)")
            df["FUENTE"] = fuente
            for col in COLUMNAS_FINALES_PT:
                if col not in df.columns: df[col] = ""
            return df[COLUMNAS_FINALES_PT]

    df_dir = leer_y_normalizar_cloud(url_directo, "Plan de trabajo", "DIRECTO")
    df_ism = leer_y_normalizar_cloud(url_ism, "Plan de trabajo CIF", "ISM")
    
    periodo = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    if not df_dir.empty or not df_ism.empty:
        df_total = pd.concat([df_dir, df_ism], ignore_index=True)
        df_total['ID_PDV_INVOLVES'] = df_total['ID_PDV_INVOLVES'].astype(str).str.strip()
        df_total = df_total.drop_duplicates(
            subset=['ID_PDV_INVOLVES', 'NOMBRE'],
        ).reset_index(drop=True)

        nombre_pt_final = f"Plan_Trabajo_SOS_Completo_{periodo}.xlsx"
        subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_SOS, nombre_pt_final, df_total)
        print(f"✅ EXITO: Plan unificado guardado en SharePoint: {nombre_pt_final}")
        return df_total, periodo
    return pd.DataFrame(), periodo

def ejecutar_paso_2_consolidar_encuestas(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 2: Consolidando Encuestas SOS ({spec.etiqueta}) ---")
    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_SOS)
    
    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    urls_encuestas = [
        a["@microsoft.graph.downloadUrl"] for a in archivos_origen 
        if CLAVE_ARCHIVO_ENCUESTA in a["name"] and periodo_nom in a["name"].upper()
    ]
    if not urls_encuestas:
        urls_encuestas = [
            a["@microsoft.graph.downloadUrl"] for a in archivos_origen 
            if CLAVE_ARCHIVO_ENCUESTA in a["name"] or "EJECUCION PARTICIPACIONES" in a["name"].upper()
        ]

    lista_dfs = [pd.read_excel(io.BytesIO(requests.get(u).content)) for u in urls_encuestas]
    df_enc_total = pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()
    
    nombre_encuesta_salida = f"Encuesta_Sos_Consolidada_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_SOS, nombre_encuesta_salida, df_enc_total)
    print(f"✅ EXITO: Encuestas unidas en SharePoint ({len(urls_encuestas)} archivo(s)).")
    return df_enc_total, nombre_encuesta_salida

def ejecutar_paso_3_cumplimiento_captura(df_pt, df_enc, spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 3: Generando Reporte de Cumplimiento ({spec.etiqueta}) ---")
    if df_pt.empty or df_enc.empty:
        # Cargar de SharePoint si los dfs vienen vacíos (por ejecución por partes)
        periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
        archivos = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_SOS)
        url_pt = next((a["@microsoft.graph.downloadUrl"] for a in archivos if f"Plan_Trabajo_SOS_Completo_{periodo_nom}" in a["name"]), None)
        url_enc = next((a["@microsoft.graph.downloadUrl"] for a in archivos if f"Encuesta_Sos_Consolidada_{periodo_nom}" in a["name"]), None)
        if url_pt and url_enc:
            df_pt = pd.read_excel(io.BytesIO(requests.get(url_pt).content))
            df_enc = pd.read_excel(io.BytesIO(requests.get(url_enc).content))
        else:
            return pd.DataFrame()

    def extraer_categoria_estricta(texto):
        if pd.isna(texto): return None
        t = re.sub(r'[^A-Z0-9\s]', ' ', str(texto).upper())
        t = " ".join(t.split())
        if "SHAMPOO BEBE EN ADULTOS" in t: return "SHAMPOO BEBE EN ADULTOS"
        if re.search(r'\bSHAMPOO BEBE\b', t): return "SHAMPOO BEBE"
        for cat in LISTA_17_CATEGORIAS:
            if cat not in ["SHAMPOO BEBE", "SHAMPOO BEBE EN ADULTOS"] and cat in t: return cat
        return None

    df_enc['CATEGORIA_LIMPIA'] = df_enc['Rótulo de la encuesta'].apply(extraer_categoria_estricta)
    df_enc_valida = df_enc.dropna(subset=['CATEGORIA_LIMPIA']).copy()

    from shared_loader import id_a_str
    df_pt['ID_PDV_INVOLVES'] = id_a_str(df_pt['ID_PDV_INVOLVES'])
    df_enc_valida['ID del PDV'] = id_a_str(df_enc_valida['ID del PDV'])

    dict_encuesta = df_enc_valida.groupby('ID del PDV')['CATEGORIA_LIMPIA'].unique().to_dict()
    set_oficial = set(LISTA_17_CATEGORIAS)
    res = []
    
    for id_pdv in df_pt['ID_PDV_INVOLVES'].unique():
        cats = dict_encuesta.get(id_pdv, [])
        faltantes = sorted(list(set_oficial - set(cats))) if len(cats) < 17 else []
        res.append({
            'ID_PDV_INVOLVES': id_pdv, 
            'CONTEO_CATEGORIAS': len(cats), 
            'CATEGORIAS_FALTANTES': ", ".join(faltantes)
        })
    
    df_final = pd.merge(df_pt, pd.DataFrame(res), on='ID_PDV_INVOLVES', how='left')
    
    if 'FUENTE' in df_final.columns:
        idx_fuente = df_final.columns.get_loc('FUENTE')
        df_final.insert(idx_fuente + 1, 'CAPTURA_PLANEADA', 1)
    else:
        df_final['CAPTURA_PLANEADA'] = 1

    df_final['CAPTURA_EJECUTADA'] = df_final['CONTEO_CATEGORIAS'].apply(lambda x: 1 if x == 17 else 0)
    df_final.loc[df_final['CONTEO_CATEGORIAS'] == 0, 'CATEGORIAS_FALTANTES'] = "SIN CAPTURAS (FALTAN 17)"

    df_final['MES'] = spec.mes
    df_final['AÑO'] = spec.anio

    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    nombre_cumplimiento = f"Cumplimiento_Captura_SOS_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_SOS, nombre_cumplimiento, df_final)
    print(f"✅ EXITO: Reporte de capturas generado en SharePoint.")
    return df_final

def ejecutar_paso_4_normalizar_target_dinamico(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 4: Normalizando Target ({spec.etiqueta}) ---")
    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_SOS)
    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    
    url_target = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_TARGET in a["name"] and periodo_nom in a["name"].upper()), None)
    if not url_target:
        url_target = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_TARGET in a["name"] or f"-{spec.anio}{spec.mes:02d}" in a["name"]), None)

    if not url_target:
        raise FileNotFoundError("No se encontró el archivo Target en SharePoint.")

    df_target = pd.read_excel(io.BytesIO(requests.get(url_target).content))
    df_target.columns = [str(c).strip().upper() for c in df_target.columns]
    
    col_id_target = "ID INVOLVES"
    col_cat = next((c for c in df_target.columns if 'CATEGOR' in c), None)
    if col_cat:
        df_target[col_cat] = df_target[col_cat].astype(str).str.upper().replace(
            "ENJUAGUES BUCALES TOTALES", "ENJUAGUES BUCALES"
        )
    
    for col in [col_cat, 'MARCA', 'NOMBRE DEL PDV']:
        if col in df_target.columns:
            df_target[col] = df_target[col].apply(eliminar_tildes)

    nombre_target_norm = f"Colombia_sos_target_Normalizado_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_SOS, nombre_target_norm, df_target)
    print(f"✅ EXITO: Target normalizado subido a SharePoint.")
    return df_target

def ejecutar_paso_5_cruce_triple_y_calculo(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 5: Generando Reporte Final (Cruce Triple y Cálculos) ({spec.etiqueta}) ---")
    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    
    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_SOS)
    url_target = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if f"Colombia_sos_target_Normalizado_{periodo_nom}" in a["name"]), None)
    url_enc = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if f"Encuesta_Sos_Consolidada_{periodo_nom}" in a["name"]), None)

    if not url_target or not url_enc:
        nombres_disponibles = [a["name"] for a in archivos_origen]
        raise FileNotFoundError(
            f"❌ No se encontró el Target normalizado o la Encuesta consolidada para '{periodo_nom}' en {paths.RUTA_CARPETA_BASES_SOS}.\n"
            f"Archivos encontrados:\n" + "\n".join(f" - {n}" for n in nombres_disponibles)
        )

    df_target = pd.read_excel(io.BytesIO(requests.get(url_target).content))
    df_enc = pd.read_excel(io.BytesIO(requests.get(url_enc).content))

    def extraer_cat_cruce(texto):
        t = eliminar_tildes(texto)
        if "SHAMPOO BEBE EN ADULTOS" in t: return "SHAMPOO BEBE EN ADULTOS"
        if "SHAMPOO BEBE" in t: return "SHAMPOO BEBE"
        for c in LISTA_17_CATEGORIAS:
            if c in t: return c
        return t

    from shared_loader import id_a_str
    df_enc['CAT_CRUCE'] = df_enc['Rótulo de la encuesta'].apply(extraer_cat_cruce)
    df_enc['ID_CRUCE'] = id_a_str(df_enc['ID del PDV'])
    df_enc['MARCA_CRUCE'] = df_enc['Marca'].apply(eliminar_tildes)

    # Detección flexible de columnas en la encuesta (ignorando tildes y símbolos)
    cols_enc_limpias = {eliminar_tildes(c): c for c in df_enc.columns}
    
    col_universo = next((cols_enc_limpias[c] for c in cols_enc_limpias if "UNIVERSO" in c and "CMS" in c), None)
    col_marca_cms = next((cols_enc_limpias[c] for c in cols_enc_limpias if ("CMS" in c or "CENTIMETROS" in c) and "MARCA" in c), None)

    if not col_universo or not col_marca_cms:
        raise KeyError(
            f"❌ No se pudieron identificar las columnas de CMS en la encuesta consolidada.\n"
            f"Columnas disponibles: {list(df_enc.columns)}"
        )
    
    df_enc_resumen = df_enc.groupby(['ID_CRUCE', 'CAT_CRUCE', 'MARCA_CRUCE']).agg({
        col_universo: 'max',
        col_marca_cms: 'max',
        'PDV': 'first'
    }).reset_index()

    col_cat_t = next((c for c in df_target.columns if 'CATEGOR' in c), None)
    df_target['ID INVOLVES'] = id_a_str(df_target['ID INVOLVES'])

    df_final = pd.merge(
        df_target,
        df_enc_resumen,
        left_on=['ID INVOLVES', col_cat_t, 'MARCA'],
        right_on=['ID_CRUCE', 'CAT_CRUCE', 'MARCA_CRUCE'],
        how='left'
    )

    df_final[col_universo] = pd.to_numeric(df_final[col_universo], errors='coerce').fillna(0)
    df_final[col_marca_cms] = pd.to_numeric(df_final[col_marca_cms], errors='coerce').fillna(0)
    df_final['PARTICIPACION_SOS'] = df_final.apply(
        lambda x: x[col_marca_cms] / x[col_universo] if x[col_universo] > 0 else 0, axis=1
    )

    df_final = df_final.rename(columns={
        'PDV': 'PUNTO DE VENTA',
        col_cat_t: 'CATEGORÍA DE PRODUCTO'
    })

    df_final["MES"] = spec.mes
    df_final["AÑO"] = spec.anio

    nombre_reporte_final = f"Reporte_SOS_Final_Calculado_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_SOS, nombre_reporte_final, df_final)
    print(f"✅ EXITO: Reporte Final SOS subido a SharePoint.")
    return df_final

def generar_resumen_kpi_sos(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 6 (V3): RESUMEN KPI SOS por gestor ({spec.etiqueta}) ---")
    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    
    archivos_salida = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_SOS)
    url_cumpl = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salida if f"Cumplimiento_Captura_SOS_{periodo_nom}" in a["name"]), None)
    
    if not url_cumpl:
        print(f"❌ No existe cumplimiento en SharePoint para {periodo_nom}")
        return

    df = pd.read_excel(io.BytesIO(requests.get(url_cumpl).content), engine='openpyxl')
    
    from shared_loader import calcular_kpi_simple_y_escribir
    n = calcular_kpi_simple_y_escribir(
        df_origen=df,
        col_planeado="CAPTURA_PLANEADA",
        col_ejecutado="CAPTURA_EJECUTADA",
        ruta_kpi_mes=str(paths.SOS_OUT_KPIS),
        ruta_kpi_historico=str(paths.SOS_OUT_KPIS_HISTORICO),
        nombre_cumplimiento="CUMPLIMIENTO",
        nombres_renombre=("PLANEADO", "EJECUTADO"),
    )
    print(f"  ✅ Mes activo: {n} gestores.")

    if os.path.exists(str(paths.SOS_OUT_KPIS)):
        df_kpi_local = pd.read_excel(str(paths.SOS_OUT_KPIS))
        subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_SOS, paths.SOS_OUT_KPIS.name, df_kpi_local)

    try:
        if os.path.exists(str(paths.SOS_OUT_KPIS)):
            print(f"  🚀 Preparando cargue del KPI de SOS consolidado a Supabase...")
            df_kpi = pd.read_excel(str(paths.SOS_OUT_KPIS))
            df_kpi.columns = df_kpi.columns.str.strip().str.lower()
            
            # Forzamos el nombre exacto que espera la tabla de Supabase ('anio')
            if 'año' in df_kpi.columns:
                df_kpi = df_kpi.rename(columns={'año': 'anio'})
            elif 'ano' in df_kpi.columns:
                df_kpi = df_kpi.rename(columns={'ano': 'anio'})
            
            df_kpi = df_kpi.replace([np.inf, -np.inf], 0)
            df_kpi_limpio = df_kpi.astype(object).where(pd.notnull(df_kpi), None)
            
            registros_json = df_kpi_limpio.to_dict(orient="records")
            supabase.table("sos_kpis").upsert(registros_json).execute()
            print(f"  ✅ ÉXITO CLOUD: {len(registros_json)} filas subidas a Supabase.")
    except Exception as ex_cloud:
        print(f"  ⚠ ADVERTENCIA EN CARGA SUPABASE: {ex_cloud}")

# =============================================================================
# 4. EJECUCIÓN PRINCIPAL
# =============================================================================

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL SOS — Eficacia (100% Cloud: SharePoint & Supabase)")
    parser.add_argument("--solo", nargs="+", choices=["pt", "enc", "cumpl", "target", "final", "kpi"], help="Pasos específicos")
    parser.add_argument("--full", action="store_true", help="Correr cadena completa")
    
    try:
        pr.cli_add_periodos_arg(parser)
    except Exception:
        parser.add_argument("--mes", type=int, default=datetime.now().month, help="Mes numérico")
        parser.add_argument("--anio", type=int, default=datetime.now().year, help="Año")

    args = parser.parse_args()
    
    try:
        specs = pr.periodos_de_args(args)
    except Exception:
        class SpecMock:
            def __init__(self, m, a):
                self.mes = m
                self.anio = a
                self.etiqueta = f"{MESES_ESPANOL[m]}_{a}"
            def __str__(self):
                return f"{self.mes:02d}/{self.anio}"
        specs = [SpecMock(getattr(args, 'mes', datetime.now().month), getattr(args, 'anio', datetime.now().year))]

    # 🛠️ CAMBIO CLAVE: Por defecto corre TODOS los pasos en orden (pt, enc, cumpl, target, final, kpi)
    pasos = ["pt", "enc", "cumpl", "target", "final", "kpi"] if (args.full or not args.solo) else args.solo

    print("🔒 Autenticando con Microsoft Graph API...")
    token = obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = obtener_site_id(headers)

    for i, spec in enumerate(specs, 1):
        print(f"\n🎯 ETL SOS — procesando periodo {spec.etiqueta} ({spec})")
        try:
            df_pt = pd.DataFrame()
            df_enc = pd.DataFrame()

            if "pt" in pasos:
                df_pt, _ = ejecutar_paso_1_consolidar_pt(spec, headers, site_id)
            if "enc" in pasos:
                df_enc, _ = ejecutar_paso_2_consolidar_encuestas(spec, headers, site_id)
            if "cumpl" in pasos:
                ejecutar_paso_3_cumplimiento_captura(df_pt, df_enc, spec, headers, site_id)
            if "target" in pasos:
                ejecutar_paso_4_normalizar_target_dinamico(spec, headers, site_id)
            if "final" in pasos:
                ejecutar_paso_5_cruce_triple_y_calculo(spec, headers, site_id)
            if "kpi" in pasos:
                generar_resumen_kpi_sos(spec, headers, site_id)

            print(f"\n🏁 PROCESO TERMINADO ({spec.etiqueta}).")
        except Exception as e:
            print(f"\n❌ OCURRIÓ UN ERROR CRÍTICO en {spec.etiqueta}: {e}")
            if len(specs) <= 1:
                raise