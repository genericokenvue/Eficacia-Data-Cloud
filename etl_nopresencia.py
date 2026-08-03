import re
import io
import os
import urllib.parse
import requests
import msal
import numpy as np  
import pandas as pd
from datetime import datetime, timedelta
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

CLAVE_ARCHIVO_NP = "Respuestas"

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

# =============================================================================
# 2. CONFIGURACIÓN DE COLUMNAS Y DICCIONARIOS
# =============================================================================

from shared_loader import (
    COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR,
    COLUMNAS_SUPERSET as COLUMNAS_FINALES_PT,
    COLUMNAS_NUMERICAS,
)

# =============================================================================
# 3. FUNCIONES DEL BLOQUE 1 (CONSOLIDACIÓN PT) — ¡SIN FILTRO DISCOUNTER!
# =============================================================================

def ejecutar_paso_1_consolidar_pt_np(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 1: Consolidando Plan de Trabajo No Presencia ({spec.etiqueta}) ---")

    archivos_en_carpeta = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_PT)
    
    num_mes = f"{spec.mes:02d}"                         
    nombre_mes = MESES_ESPANOL[spec.mes].upper()  

    url_directo = None
    url_ism = None

    for archivo in archivos_en_carpeta:
        nombre_archivo_upper = archivo["name"].upper()
        
        if "DIRECTO" in nombre_archivo_upper and (num_mes in nombre_archivo_upper or nombre_mes in nombre_archivo_upper):
            url_directo = archivo["@microsoft.graph.downloadUrl"]
            print(f"  📂 Archivo DIRECTO detectado en RUTA_CARPETA_PT: {archivo['name']}")
            
        if "ISM" in nombre_archivo_upper and (num_mes in nombre_archivo_upper or nombre_mes in nombre_archivo_upper):
            url_ism = archivo["@microsoft.graph.downloadUrl"]
            print(f"  📂 Archivo ISM detectado en RUTA_CARPETA_PT: {archivo['name']}")

    if not url_directo:
        url_directo = next((a["@microsoft.graph.downloadUrl"] for a in archivos_en_carpeta if "DIRECTO" in a["name"].upper()), None)
    if not url_ism:
        url_ism = next((a["@microsoft.graph.downloadUrl"] for a in archivos_en_carpeta if "ISM" in a["name"].upper()), None)

    if not url_directo or not url_ism:
        raise FileNotFoundError(f"No se localizaron los archivos base en la carpeta cloud: {paths.RUTA_CARPETA_PT}")

    def leer_y_normalizar_cloud_np(url_download, hoja, fuente):
        content = requests.get(url_download).content
        with pd.ExcelFile(io.BytesIO(content)) as xls:
            df_val = pd.read_excel(xls, sheet_name="Captura de modulos")
            df_val.columns = df_val.columns.str.strip().str.upper()
            col_filtro = "NO PRESENCIA_FINAL" if fuente == "ISM" else "NO PRESENCIA"
            if col_filtro not in df_val.columns: return pd.DataFrame()

            obs_col = next((c for c in df_val.columns if c.startswith('OBSERVAC')), None)
            if obs_col is None:
                print(f"⚠ {fuente}: hoja Captura de modulos sin OBSERVACIÓN")
                return pd.DataFrame()

            df_val_filtrado = df_val[df_val[col_filtro] == 1].copy()

            OBS_A_ROL = {
                "REPORTA GESTOR": "GESTOR",
                "REPORTA SUPERVISOR": "SUPERVISOR",
                "REPORTA GENERADOR DE DEMANDA": "GENERADOR DE DEMANDA",
            }

            pdvs_obs = df_val_filtrado[["ID PDV INVOLVES", obs_col]].rename(columns={obs_col: "OBSERVACION"})
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
            for col in COLUMNAS_NUMERICAS:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce')

            if "FECHA" in df.columns:
                df["FECHA"] = pd.to_datetime(df["FECHA"], errors='coerce').dt.strftime("%d/%m/%Y")

            for col in COLUMNAS_FINALES_PT:
                if col not in df.columns: df[col] = ""

            df = df[COLUMNAS_FINALES_PT].copy()
            for col in df.columns:
                if col not in COLUMNAS_NUMERICAS and col != "FECHA":
                    df[col] = df[col].astype(str).str.strip().replace(['nan', 'NaT', 'None'], '')

            return df

    df_dir = leer_y_normalizar_cloud_np(url_directo, "Plan de trabajo", "DIRECTO")
    df_ism = leer_y_normalizar_cloud_np(url_ism, "Plan de trabajo CIF", "ISM")

    periodo = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"

    if not df_dir.empty or not df_ism.empty:
        df_total = pd.concat([df_dir, df_ism], ignore_index=True)
        nombre_pt_final = f"Plan_Trabajo_NoPresencia_Completo_{periodo}.xlsx"
        subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_NP, nombre_pt_final, df_total)
        print(f"✅ EXITO: Plan unificado No Presencia guardado en SharePoint: {nombre_pt_final}")
        return df_total, periodo

    return pd.DataFrame(), periodo

# =============================================================================
# 4. FUNCIONES DEL BLOQUE 2 (ENCUESTAS Y MATRIZ)
# =============================================================================

def procesar_no_presencia(periodo_nom, spec: pr.PeriodoSpec, headers, site_id):
    print("\n--- INICIANDO: Bloque No Presencia (Encuestas Básicas) ---")
    
    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_NP)
    url_encuesta = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_NP in a["name"] and periodo_nom in a["name"].upper()), None)
    if not url_encuesta:
        url_encuesta = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_NP in a["name"]), None)
        
    if not url_encuesta:
        raise FileNotFoundError(f"No se encontró el archivo cloud de encuestas No Presencia en: {paths.RUTA_CARPETA_BASES_NP}")

    content_enc = requests.get(url_encuesta).content
    df = pd.read_excel(io.BytesIO(content_enc))
    df['Fecha de la encuesta'] = pd.to_datetime(df['Fecha de la encuesta'], dayfirst=True)
    
    fecha_minima = df['Fecha de la encuesta'].min()
    lunes_inicio = fecha_minima - timedelta(days=fecha_minima.weekday())
    semana_base = lunes_inicio.isocalendar()[1]

    def calcular_semana_dinamica(fecha):
        try: return f"SEMANA {(fecha.isocalendar()[1] - semana_base) + 1}"
        except: return "N/A"

    df['SEMANA'] = df['Fecha de la encuesta'].apply(calcular_semana_dinamica)
    df['Fecha de la encuesta'] = df['Fecha de la encuesta'].dt.date

    columnas_requeridas = [
        "ID de la encuesta", "Marca", "Categoría de producto", "Supercategoría", 
        "Línea de producto", "Producto (SKU)", "Código de barras", "ID del PDV", 
        "PDV", "Fecha de la encuesta", "Empleado", "Perfil de acceso", 
        "¿EL PRODUCTO ESTÁ AGOTADO?", "¿CUÁL ES LA CAUSAL DE QUE EL PRODUCTO ESTÁ AGOTADO?", 
        "SEMANA"
    ]
    df_resultado = df[[c for c in columnas_requeridas if c in df.columns]].copy()
    
    nombre_salida = f"NO_PRESENCIA_PROCESADO_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_NP, nombre_salida, df_resultado)
    return df_resultado

def procesar_plan_de_trabajo_semanas(df_pt, periodo_nom):
    print("\n--- INICIANDO: Bloque Plan de Trabajo (Semanas) ---")
    df = df_pt.copy()
    df['FECHA'] = pd.to_datetime(df['FECHA'], dayfirst=True)
    df = df.dropna(subset=['FECHA']).copy()
    
    fecha_minima = df['FECHA'].min()
    lunes_inicio = fecha_minima - timedelta(days=fecha_minima.weekday())
    semana_base = lunes_inicio.isocalendar()[1]

    df['SEMANA'] = df['FECHA'].apply(lambda x: f"SEMANA {(x.isocalendar()[1] - semana_base) + 1}")
    df['FECHA'] = df['FECHA'].dt.date
    return df

def generar_matriz_seguimiento(df_np, df_pt, periodo_nom, spec, headers, site_id):
    print("\n--- GENERANDO MATRIZ COMPARATIVA FINAL NO PRESENCIA ---")
    
    lista_capturas = [df_np[['ID del PDV']]]
    
    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_NP)
    for marca in ['ARA', 'D1', 'OXXO']:
        url_dinamico = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if marca in a["name"].upper() and periodo_nom in a["name"].upper()), None)
        if url_dinamico:
            print(f"📥 Archivo dinámico detectado en la nube e integrando: {marca}")
            try:
                content_din = requests.get(url_dinamico).content
                df_extra = pd.read_excel(io.BytesIO(content_din))
                df_extra.columns = df_extra.columns.str.strip().str.upper()
                if 'ID DEL PDV' in df_extra.columns:
                    df_extra = df_extra.rename(columns={'ID DEL PDV': 'ID del PDV'})
                    lista_capturas.append(df_extra[['ID del PDV']])
            except Exception as e:
                print(f"❌ No se pudo procesar el archivo adicional de {marca}: {e}")

    df_capturas_totales = pd.concat(lista_capturas, ignore_index=True)

    if spec is not None:
        df_pt['MES'] = spec.mes
        df_pt['AÑO'] = spec.anio
    else:
        fechas_dt = pd.to_datetime(df_pt['FECHA'], dayfirst=True, errors='coerce')
        df_pt['MES'] = fechas_dt.dt.month.fillna(0).astype(int)
        df_pt['AÑO'] = fechas_dt.dt.year.fillna(0).astype(int)

    from shared_loader import id_a_str
    df_pt['ID_PDV_INVOLVES'] = id_a_str(df_pt['ID_PDV_INVOLVES'])
    df_pt['NOMBRE'] = df_pt['NOMBRE'].astype(str).str.strip().str.upper()
    df_capturas_totales['ID del PDV'] = id_a_str(df_capturas_totales['ID del PDV'])

    plan_agg = (
        df_pt.assign(_PLAN=1)
             .groupby(
                 ['ID_PDV_INVOLVES', 'NOMBRE_PDV', 'NOMBRE',
                  'ROL', 'SUPERVISOR_LIDER', 'MES', 'AÑO'],
                 dropna=False, as_index=False,
             )
             .agg(PLANEADO_MES=('_PLAN', 'max'))
    )

    cap_agg = (
        df_capturas_totales.assign(_CAP=1)
             .groupby(['ID del PDV'], dropna=False, as_index=False)
             .agg(CAPTURA_MES=('_CAP', 'max'))
             .rename(columns={'ID del PDV': 'ID_PDV_INVOLVES'})
    )

    matriz = plan_agg.merge(cap_agg, on=['ID_PDV_INVOLVES'], how='left')
    matriz['CAPTURA_MES'] = matriz['CAPTURA_MES'].fillna(0).astype(int)
    matriz['PLANEADO_MES'] = matriz['PLANEADO_MES'].fillna(0).astype(int)
    matriz['%_CUMPLIMIENTO_MES'] = np.where(
        matriz['PLANEADO_MES'] > 0,
        matriz['CAPTURA_MES'] / matriz['PLANEADO_MES'],
        0,
    )

    columnas_finales_reporte = [
        'ID_PDV_INVOLVES', 'NOMBRE_PDV', 'NOMBRE', 'ROL', 'SUPERVISOR_LIDER',
        'PLANEADO_MES', 'CAPTURA_MES', '%_CUMPLIMIENTO_MES', 'MES', 'AÑO'
    ]
    matriz = matriz[columnas_finales_reporte]

    nombre_matriz = f"REPORTE_NO_PRESENCIA_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_NP, nombre_matriz, matriz)
    print(f"🏁 FINALIZADO: Reporte generado en SharePoint: {nombre_matriz}")

# =============================================================================
# 5. FUNCIONES DE ANÁLISIS Y KPI EN CLOUD
# =============================================================================

def generar_analisis_agotados(periodo_nom, headers, site_id):
    print("\n--- GENERANDO ANÁLISIS DE AGOTADOS (INVOLVES) ---")
    archivos_salida = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_NP)
    url_np = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salida if f"NO_PRESENCIA_PROCESADO_{periodo_nom}" in a["name"]), None)
    
    if not url_np:
        print("❌ No se encontró el archivo procesado de No Presencia para agotados.")
        return

    content_np = requests.get(url_np).content
    df = pd.read_excel(io.BytesIO(content_np))
    df['Fecha de la encuesta'] = pd.to_datetime(df['Fecha de la encuesta'])
    df['MES'] = df['Fecha de la encuesta'].dt.month
    df['AÑO'] = df['Fecha de la encuesta'].dt.year
    df['ES_AGOTADO'] = df['¿EL PRODUCTO ESTÁ AGOTADO?'].apply(lambda x: 1 if str(x).strip().upper() == 'SI' else 0)
    df['PRODUCTOS_MEDIDOS'] = 1

    analisis = df.groupby(['MES', 'AÑO', 'ID del PDV', 'PDV', 'Marca']).agg({
        'PRODUCTOS_MEDIDOS': 'sum',
        'ES_AGOTADO': 'sum'
    }).reset_index()

    analisis.rename(columns={'ES_AGOTADO': 'CANT_AGOTADOS'}, inplace=True)
    analisis['%_AGOTADOS'] = (analisis['CANT_AGOTADOS'] / analisis['PRODUCTOS_MEDIDOS']).fillna(0)

    cols = [c for c in analisis.columns if c not in ['MES', 'AÑO']] + ['MES', 'AÑO']
    analisis = analisis[cols]

    nombre_agotados = f"ANALISIS_AGOTADOS_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_NP, nombre_agotados, analisis)
    print(f"✅ ÉXITO: Reporte de Agotados guardado en SharePoint.")

def generar_resumen_kpi_no_presencia(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO KPI (V3): RESUMEN NP por gestor ({spec.etiqueta}) ---")
    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    
    archivos_salida = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_NP)
    url_reporte = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salida if f"REPORTE_NO_PRESENCIA_{periodo_nom}" in a["name"]), None)
    
    if not url_reporte:
        print(f"❌ No existe el reporte del periodo en SharePoint: REPORTE_NO_PRESENCIA_{periodo_nom}")
        return
        
    content_rep = requests.get(url_reporte).content
    df = pd.read_excel(io.BytesIO(content_rep), engine='openpyxl')
    print(f"  Leyendo desde cloud: REPORTE_NO_PRESENCIA_{periodo_nom}.xlsx ({len(df)} filas)")

    # Ver nota en etl_precios.py: el histórico vive en disco local dentro de
    # calcular_kpi_simple_y_escribir — hay que descargarlo de SharePoint
    # ANTES de llamarla, o en GitHub Actions se resetearía cada corrida.
    nombre_hist = paths.NP_OUT_KPIS_HISTORICO.name
    url_hist = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salida if a["name"] == nombre_hist), None)
    if url_hist:
        os.makedirs(os.path.dirname(str(paths.NP_OUT_KPIS_HISTORICO)), exist_ok=True)
        with open(str(paths.NP_OUT_KPIS_HISTORICO), "wb") as f:
            f.write(requests.get(url_hist).content)
        print(f"  ⏳ Histórico descargado desde SharePoint: {nombre_hist}")
    else:
        print(f"  ℹ️  No hay histórico previo en SharePoint ({nombre_hist}) — se creará uno nuevo")

    from shared_loader import calcular_kpi_simple_y_escribir
    n = calcular_kpi_simple_y_escribir(
        df_origen=df,
        col_planeado="PLANEADO_MES",
        col_ejecutado="CAPTURA_MES",
        ruta_kpi_mes=str(paths.NP_OUT_KPIS),
        ruta_kpi_historico=str(paths.NP_OUT_KPIS_HISTORICO),
        nombre_cumplimiento="EJECUCION",
        nombres_renombre=("SUMA PLANEADO", "SUMA CAPTURADO"),
    )
    print(f"  ✅ Mes activo: {n} gestores.")

    if os.path.exists(str(paths.NP_OUT_KPIS)):
        df_kpi_local = pd.read_excel(str(paths.NP_OUT_KPIS))
        subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_NP, paths.NP_OUT_KPIS.name, df_kpi_local)

    if os.path.exists(str(paths.NP_OUT_KPIS_HISTORICO)):
        df_hist_local = pd.read_excel(str(paths.NP_OUT_KPIS_HISTORICO))
        subir_archivo_a_sharepoint(headers, site_id, paths.RUTA_CARPETA_SALIDAS_NP, nombre_hist, df_hist_local)
        print(f"  ✅ Histórico actualizado subido a SharePoint: {nombre_hist}")

    try:
        if os.path.exists(str(paths.NP_OUT_KPIS)):
            print(f"  🚀 Preparando cargue del KPI de No Presencia consolidado a Supabase...")
            df_kpi = pd.read_excel(str(paths.NP_OUT_KPIS))
            
            # Limpieza exhaustiva de nombres de columnas para Supabase
            df_kpi.columns = (
                df_kpi.columns.str.strip()
                .str.lower()
                .str.replace(" ", "_")
                .str.replace("á", "a")
                .str.replace("é", "e")
                .str.replace("í", "i")
                .str.replace("ó", "o")
                .str.replace("ú", "u")
            )
            
            if 'año' in df_kpi.columns:
                df_kpi = df_kpi.rename(columns={'año': 'anio'})
            
            df_kpi = df_kpi.replace([np.inf, -np.inf], 0)
            df_kpi_limpio = df_kpi.astype(object).where(pd.notnull(df_kpi), None)
            
            registros_json = df_kpi_limpio.to_dict(orient="records")
            
            supabase.table("no_presencia_kpis").upsert(registros_json).execute()
            print(f"  ✅ ÉXITO CLOUD: {len(registros_json)} filas subidas de forma exitosa a Supabase.")
    except Exception as ex_cloud:
        print(f"  ⚠ ADVERTENCIA EN CARGA CLOUD: {ex_cloud}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL NO PRESENCIA — Eficacia Cloud (Sprint 16.1: multi-periodo)")
    parser.add_argument("--solo", nargs="+", choices=["full", "kpi"],
                        help="Por default: corre todo + kpi. Usa --solo kpi para solo KPI.")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    pasos = args.solo or ["full", "kpi"]
    specs = pr.periodos_de_args(args)

    print("🔒 Autenticando con Microsoft Graph API...")
    token = obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = obtener_site_id(headers)

    for i, spec in enumerate(specs, 1):
        print("=" * 50)
        print(f" INICIANDO ETL UNIFICADO NO PRESENCIA CLOUD ({spec.etiqueta})")
        print("=" * 50)

        try:
            periodo_final = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
            df_pt = pd.DataFrame()

            if "full" in pasos:
                df_pt, periodo_final = ejecutar_paso_1_consolidar_pt_np(spec, headers, site_id)

                if df_pt.empty:
                    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, paths.RUTA_CARPETA_BASES_NP)
                    nombre_buscado = f"Plan_Trabajo_NoPresencia_Completo_{periodo_final}.xlsx"
                    url_pt_periodo = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if a["name"] == nombre_buscado), None)
                    
                    if url_pt_periodo:
                        content_pt = requests.get(url_pt_periodo).content
                        df_pt = pd.read_excel(io.BytesIO(content_pt))
                        print(f"📂 Cargado PT No Presencia del periodo desde SharePoint: {nombre_buscado}")

                if not df_pt.empty:
                    df_np_proc = procesar_no_presencia(periodo_final, spec, headers, site_id)
                    df_pt_proc = procesar_plan_de_trabajo_semanas(df_pt, periodo_final)

                    if not df_np_proc.empty and not df_pt_proc.empty:
                        generar_matriz_seguimiento(df_np_proc, df_pt_proc, periodo_final, spec=spec, headers=headers, site_id=site_id)
                        generar_analisis_agotados(periodo_final, headers, site_id)
            
            if "kpi" in pasos:
                generar_resumen_kpi_no_presencia(spec, headers, site_id)

            print(f"\n🏁 PROCESO TERMINADO ({spec.etiqueta})")

        except Exception as e:
            print(f"\n❌ ERROR en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise