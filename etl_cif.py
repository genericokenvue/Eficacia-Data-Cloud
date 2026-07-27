import sys
# Forzar UTF-8 en stdout (Windows cp1252 explota con emojis 📅 ⚠️ ❌)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

import pandas as pd
import numpy as np
import io
import os
import urllib.parse
import requests
import msal
from datetime import datetime, timedelta

# ─────────────────────────────────────────────
# CONFIGURACIÓN GLOBAL & AZURE API / SUPABASE
# ─────────────────────────────────────────────
import paths
import periodo_resolver as pr
from dotenv import load_dotenv
from supabase import create_client

load_dotenv()

TENANT_ID = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET, SUPABASE_URL, SUPABASE_KEY]):
    raise ValueError("❌ ERROR: Faltan credenciales esenciales. Verifica tu archivo .env o tus GitHub Secrets.")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# ─────────────────────────────────────────────
# FUNCIONES AUXILIARES DE MICROSOFT GRAPH API
# ─────────────────────────────────────────────
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
    # Convertir a texto plano y limpiar barras invertidas de Windows por si acaso
    ruta_str = str(ruta_carpeta).strip().replace("\\", "/").lstrip('/').rstrip('/')
    
    if not ruta_str:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    else:
        # Codificar cada segmento por separado para mantener las barras '/' intactas
        ruta_formateada = "/".join(urllib.parse.quote(segment) for segment in ruta_str.split("/"))
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_formateada}:/children"
    
    response = requests.get(url, headers=headers)
    res_json = response.json()
    
    if response.status_code != 200:
        print(f"❌ Error de la API de Graph ({response.status_code}) en ruta '{ruta_str}': {res_json.get('error', {}).get('message')}")
        return []
    
    return res_json.get("value", [])

def descargar_archivo_bytes(url_download):
    response = requests.get(url_download)
    if response.status_code == 200:
        return response.content
    raise Exception(f"Error al descargar archivo desde SharePoint: {response.status_code}")

def subir_archivo_a_sharepoint(headers, site_id, ruta_carpeta, nombre_archivo, dataframe_o_bytes, es_excel=True):
    if es_excel:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            dataframe_o_bytes.to_excel(writer, index=False)
        buffer.seek(0)
        content = buffer.getvalue()
        content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        content = dataframe_o_bytes
        content_type = "text/csv; charset=utf-8"

    url_subida = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_carpeta}/{nombre_archivo}:/content"
    headers_subida = {**headers, "Content-Type": content_type}
    
    response = requests.put(url_subida, headers=headers_subida, data=content)
    if response.status_code in [200, 201]:
        print(f"  ✅ SHAREPOINT: Archivo guardado con éxito -> {ruta_carpeta}/{nombre_archivo}")
    else:
        print(f"  ❌ Error al subir a SharePoint ({nombre_archivo}): {response.text}")

# ─────────────────────────────────────────────
# MAPEO COLUMNAS  ── importado desde shared_loader (fuente canónica)
# ─────────────────────────────────────────────
from shared_loader import (
    COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR,
    COLUMNAS_SUPERSET            as COLUMNAS_FINALES,
    COLUMNAS_NUMERICAS,
)

# ─────────────────────────────────────────────
# BLOQUE 1: CONSOLIDACIÓN PLAN DE TRABAJO DIRECTO Y PROXIMITY
# ─────────────────────────────────────────────
def leer_y_normalizar_cloud(url_download, hoja, fuente):
    content = descargar_archivo_bytes(url_download)
    with pd.ExcelFile(io.BytesIO(content)) as xls:
        if hoja not in xls.sheet_names:
            print(f"⚠️  Hoja '{hoja}' no encontrada en {fuente}")
            return pd.DataFrame()
        df = pd.read_excel(xls, sheet_name=hoja)
    
    df.columns = df.columns.str.strip()
    df.rename(columns={col: COLUMNAS_ESTANDAR[col] for col in df.columns if col in COLUMNAS_ESTANDAR}, inplace=True)
    df["FUENTE"] = fuente

    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors='coerce').dt.strftime("%d/%m/%Y")

    for col in COLUMNAS_NUMERICAS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df[[c for c in COLUMNAS_FINALES if c in df.columns]].copy()
    
    for col in df.columns:
        if col not in COLUMNAS_NUMERICAS:
            df[col] = df[col].astype(str).str.strip().replace(['nan', 'NaT', 'None'], '')

    print(f"   [{fuente}] {len(df)} filas procesadas.")
    return df

def ejecutar_paso_1(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 1: CONSOLIDACIÓN CLOUD  ({spec.etiqueta})")
    print("="*40)
    
    ruta_pt_sharepoint = getattr(paths, 'RUTA_CARPETA_PT_CIF', getattr(paths, 'RUTA_CARPETA_PT', "Equipo Información/BI/INVOLVES/PLAN DE TRABAJO"))
    archivos_pt = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_pt_sharepoint)
    num_mes = f"{spec.mes:02d}"
    nombre_mes = pr.MESES_ESPANOL[spec.mes].upper() if hasattr(pr, 'MESES_ESPANOL') else ""

    url_directo = None
    url_ism = None

    for archivo in archivos_pt:
        nombre_upper = archivo["name"].upper()
        if "DIRECTO" in nombre_upper and (num_mes in nombre_upper or nombre_mes in nombre_upper):
            url_directo = archivo["@microsoft.graph.downloadUrl"]
        if "ISM" in nombre_upper and (num_mes in nombre_upper or nombre_mes in nombre_upper):
            url_ism = archivo["@microsoft.graph.downloadUrl"]

    if not url_directo:
        url_directo = next((a["@microsoft.graph.downloadUrl"] for a in archivos_pt if "DIRECTO" in a["name"].upper()), None)
    if not url_ism:
        url_ism = next((a["@microsoft.graph.downloadUrl"] for a in archivos_pt if "ISM" in a["name"].upper()), None)

    if not url_directo or not url_ism:
        raise FileNotFoundError(f"Paso 1 ({spec.etiqueta}): faltan PT del periodo en la nube.")

    df_dir = leer_y_normalizar_cloud(url_directo, "Plan de trabajo", "DIRECTO")
    df_ism = leer_y_normalizar_cloud(url_ism, "Plan de trabajo CIF", "ISM")
    
    df_total = pd.concat([df_dir, df_ism], ignore_index=True)
    
    for col in COLUMNAS_FINALES:
        if col not in df_total.columns:
            df_total[col] = ""
            
    periodo_nom = f"{pr.MESES_ESPANOL[spec.mes]}_{spec.anio}" if hasattr(pr, 'MESES_ESPANOL') else f"{spec.mes:02d}_{spec.anio}"
    nombre_consolidado = f"Consolidado_PT_{periodo_nom}.csv"
    
    ruta_bases_sharepoint = getattr(paths, 'RUTA_CARPETA_BASES_CIF', getattr(paths, 'RUTA_CARPETA_BASES', "Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS"))
    csv_buffer = io.StringIO()
    df_total[COLUMNAS_FINALES].to_csv(csv_buffer, index=False, encoding="utf-8-sig", sep=";", decimal=",")
    subir_archivo_a_sharepoint(headers, site_id, ruta_bases_sharepoint, nombre_consolidado, csv_buffer.getvalue().encode('utf-8-sig'), es_excel=False)
    print(f"\n✅ ÉXITO: Consolidado guardado en SharePoint: {ruta_bases_sharepoint}/{nombre_consolidado}")

# ─────────────────────────────────────────────
# BLOQUE 2: AGRUPACIÓN PLAN DE TRABAJO (DuckDB / Pandas in-memory)
# ─────────────────────────────────────────────
def ejecutar_paso_2(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "="*40)
    print(" INICIANDO PASO 2: AGRUPACIÓN EN LA NUBE")
    print("="*40)
    
    periodo_nom = f"{pr.MESES_ESPANOL[spec.mes]}_{spec.anio}" if hasattr(pr, 'MESES_ESPANOL') else f"{spec.mes:02d}_{spec.anio}"
    nombre_consolidado = f"Consolidado_PT_{periodo_nom}.csv"
    nombre_agrupado = f"Agrupado_PT_{periodo_nom}.csv"

    ruta_bases_sharepoint = getattr(paths, 'RUTA_CARPETA_BASES_CIF', getattr(paths, 'RUTA_CARPETA_BASES', "Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS"))
    archivos_bases = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_bases_sharepoint)
    url_con = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_consolidado in a["name"]), None)

    if not url_con:
        print(f"❌ ERROR: No existe el archivo consolidado {nombre_consolidado} en SharePoint.")
        return

    content = descargar_archivo_bytes(url_con)
    df = pd.read_csv(io.BytesIO(content), sep=";", decimal=",", encoding="utf-8-sig", 
                     dtype={'ID_PDV_INVOLVES': str, 'CEDULA': str})

    import sqlite3
    conexion = sqlite3.connect(':memory:')
    df.to_sql('plan_consolidado', conexion, index=False, if_exists='replace')

    query = """
    SELECT 
        ID_PDV_INVOLVES, ACRONIMO,
        MAX(NOMBRE_PDV) AS NOMBRE_PDV,
        MAX(VENTAS_PROMEDIO_MES) AS VENTAS_PROMEDIO_MES,
        MAX(CEDULA) AS CEDULA,
        MAX(NOMBRE) AS NOMBRE,
        MAX(COD_MERCADERISTA) AS COD_MERCADERISTA,
        COUNT(FECHA) AS CANTIDAD_VISITAS,
        SUM(TIEMPO_SERVICIO) AS TIEMPO_SERVICIO,
        (SUM(TIEMPO_SERVICIO) / 60.0) AS TIEMPO_SERVICIO_HORAS,
        MAX(FREC_SEMANAL) AS FREC_SEMANAL,
        MAX(FREC_MENSUAL) AS FREC_MENSUAL,
        MAX(HORAS_X_VISITA) AS HORAS_X_VISITA,
        MAX(HORAS_MES) AS HORAS_MES,
        MAX(ROL) AS ROL,
        MAX(SUPERVISOR_LIDER) AS SUPERVISOR_LIDER,
        MAX(FUENTE) AS FUENTE
    FROM plan_consolidado
    GROUP BY ACRONIMO, ID_PDV_INVOLVES
    """

    df_agrupado = pd.read_sql_query(query, conexion)
    conexion.close()

    csv_buffer = io.StringIO()
    df_agrupado.to_csv(csv_buffer, index=False, sep=";", decimal=",", encoding="utf-8-sig")
    subir_archivo_a_sharepoint(headers, site_id, ruta_bases_sharepoint, nombre_agrupado, csv_buffer.getvalue().encode('utf-8-sig'), es_excel=False)
    print(f"✅ ÉXITO: Agrupado guardado en SharePoint: {ruta_bases_sharepoint}/{nombre_agrupado}")


# ─────────────────────────────────────────────
# BLOQUE 3: INVOLVES VISITAS REALIZADAS Y VALIDAS
# ─────────────────────────────────────────────
def ejecutar_paso_3(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 3: INVOLVES NUBE (CRUCE PRIORIZANDO BASE VENTAS)  ({spec.etiqueta})")
    print("="*40)

    periodo_nom = f"{pr.MESES_ESPANOL[spec.mes]}_{spec.anio}" if hasattr(pr, 'MESES_ESPANOL') else f"{spec.mes:02d}_{spec.anio}"
    nombre_consolidado = f"Consolidado_PT_{periodo_nom}.csv"
    nombre_output_involves = f"Involves_Procesado_{periodo_nom}.csv"

    ruta_bases_sharepoint = getattr(paths, 'RUTA_CARPETA_BASES_CIF', getattr(paths, 'RUTA_CARPETA_BASES', "Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS"))
    archivos_bases = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_bases_sharepoint)
    
    url_plan = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_consolidado in a["name"]), None)
    url_colab = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if "colaboradores" in a["name"].lower()), None)
    url_ventas = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if "ventas" in a["name"].lower()), None)
    
    # Búsqueda en la nube adaptada para evitar rutas locales (pr.cif_involves)
    url_inv = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if "informe-gerencial-visitas" in a["name"].lower() and (str(spec.mes) in a["name"] or str(spec.anio) in a["name"])), None)
    
    if not url_inv:
        # Buscar en la carpeta específica de bases de respuestas del periodo en SharePoint
        ruta_bases_respuestas_cif = f"Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS/CIF/{spec.anio}"
        archivos_gen = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_bases_respuestas_cif)
        url_inv = next((a["@microsoft.graph.downloadUrl"] for a in archivos_gen if "informe-gerencial-visitas" in a["name"].lower() or periodo_nom.lower() in a["name"].lower()), None)

    if not all([url_plan, url_colab, url_ventas, url_inv]):
        print("❌ ERROR: Faltan archivos requeridos en SharePoint para el Paso 3.")
        return

    df_maestro_plan = pd.read_csv(io.BytesIO(descargar_archivo_bytes(url_plan)), sep=";", decimal=",", encoding="utf-8-sig")
    df_maestro_plan['ID_PDV_INVOLVES'] = pd.to_numeric(df_maestro_plan['ID_PDV_INVOLVES'], errors='coerce').fillna(0).astype(int).astype(str)
    df_maestro_plan = df_maestro_plan[['NOMBRE_PDV', 'ID_PDV_INVOLVES']].drop_duplicates(subset=['NOMBRE_PDV'])
    df_maestro_plan.rename(columns={'NOMBRE_PDV': 'P_NOMBRE', 'ID_PDV_INVOLVES': 'ID_PLAN'}, inplace=True)

    df_maestro_ventas = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_ventas)))
    df_maestro_ventas['ID'] = pd.to_numeric(df_maestro_ventas['ID'], errors='coerce').fillna(0).astype(int).astype(str)
    df_maestro_ventas = df_maestro_ventas[['Nombre del PDV', 'ID']].drop_duplicates(subset=['Nombre del PDV'])
    df_maestro_ventas.rename(columns={'Nombre del PDV': 'V_NOMBRE', 'ID': 'ID_VENTAS'}, inplace=True)

    df_colab = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_colab)))[['Nombre del empleado', 'Usuario', 'Perfil de acceso']].drop_duplicates(subset=['Nombre del empleado'])
    df_colab.rename(columns={'Nombre del empleado': 'EMPLEADO_KEY', 'Usuario': 'USUARIO_INVOLVES', 'Perfil de acceso': 'PERFIL_ACCESO'}, inplace=True)

    df_inv_raw = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_inv)), sheet_name="Sheet0")
    df = df_inv_raw.merge(df_colab, left_on='Empleado', right_on='EMPLEADO_KEY', how='left')
    
    columnas_interes = {
        'Punto de venta': 'PUNTO_DE_VENTA',
        'Empleado': 'EMPLEADO',
        'USUARIO_INVOLVES': 'USUARIO_INVOLVES', 
        'PERFIL_ACCESO': 'PERFIL_ACCESO',       
        'Fecha de la visita': 'FECHA_VISITA',
        'Tipo de check-in': 'TIPO_CHECKIN',
        'Primer check-in manual': 'CHECKIN_MANUAL',
        'Último check-out manual': 'CHECKOUT_MANUAL',
        'Check-in automático': 'CHECKIN_AUTO',
        'Check-out automático': 'CHECKOUT_AUTO',
        'Justificación': 'JUSTIFICACION',
        'Justificación auditada el': 'JUSTIFICADO_EL',
        'Quién justificó': 'QUIEN_JUSTIFICO'
    }
    
    df = df[[c for c in columnas_interes.keys() if c in df.columns]].rename(columns=columnas_interes)
    df = df[df['TIPO_CHECKIN'].isin(['Manual', 'Manual y GPS'])].copy()

    df = df.merge(df_maestro_ventas, left_on='PUNTO_DE_VENTA', right_on='V_NOMBRE', how='left')
    df = df.merge(df_maestro_plan, left_on='PUNTO_DE_VENTA', right_on='P_NOMBRE', how='left')
    df['ID_PDV_INVOLVES'] = df['ID_VENTAS'].fillna(df['ID_PLAN'])

    time_cols = ['CHECKIN_MANUAL', 'CHECKOUT_MANUAL', 'CHECKIN_AUTO', 'CHECKOUT_AUTO']
    for col in time_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format='%H:%M:%S', errors='coerce')

    sin_checkout = df['CHECKOUT_MANUAL'].isna()
    es_2359 = (
        (df['CHECKOUT_MANUAL'].notna() & df['CHECKOUT_MANUAL'].dt.hour.eq(23) & df['CHECKOUT_MANUAL'].dt.minute.eq(59)) |
        (df['CHECKOUT_AUTO'].notna()   & df['CHECKOUT_AUTO'].dt.hour.eq(23)   & df['CHECKOUT_AUTO'].dt.minute.eq(59))
    )

    t_manual = (df['CHECKOUT_MANUAL'] - df['CHECKIN_MANUAL']).dt.total_seconds().div(60).clip(lower=0)
    t_auto   = (df['CHECKOUT_AUTO']   - df['CHECKIN_MANUAL']).dt.total_seconds().div(60).clip(lower=0)

    hay_auto        = df['CHECKOUT_AUTO'].notna()
    dif_auto_manual = (df['CHECKOUT_AUTO'] - df['CHECKOUT_MANUAL']).dt.total_seconds().div(60)
    usar_auto       = hay_auto & (dif_auto_manual.abs() > 60)
    tiempo_base     = np.where(usar_auto, t_auto, t_manual)

    df['TIEMPO_SERVICIO_MIN'] = np.where(sin_checkout | es_2359, 0.0, tiempo_base).clip(min=0)
    df['VISITA_OK'] = df['TIEMPO_SERVICIO_MIN'] >= 10

    for col in time_cols:
        if col in df.columns:
            df[col] = df[col].dt.strftime('%H:%M:%S')

    columnas_a_quitar = ['EMPLEADO_KEY', 'V_NOMBRE', 'ID_VENTAS', 'P_NOMBRE', 'ID_PLAN']
    cols_finales = ['ID_PDV_INVOLVES'] + [c for c in df.columns if c not in columnas_a_quitar and c != 'ID_PDV_INVOLVES']
    
    df_detalle_final = df[cols_finales].drop_duplicates()

    csv_buffer = io.StringIO()
    df_detalle_final.to_csv(csv_buffer, index=False, sep=";", decimal=",", encoding="utf-8-sig")
    subir_archivo_a_sharepoint(headers, site_id, ruta_bases_sharepoint, nombre_output_involves, csv_buffer.getvalue().encode('utf-8-sig'), es_excel=False)
    print(f"✅ Paso 3 en nube finalizado.")


# ─────────────────────────────────────────────
# BLOQUE 4: INVOLVES VISITAS JUSTIFICADAS
# ─────────────────────────────────────────────
def ejecutar_paso_4(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 4: JUSTIFICACIONES NUBE  ({spec.etiqueta})")
    print("="*40)

    periodo_nom = f"{pr.MESES_ESPANOL[spec.mes]}_{spec.anio}" if hasattr(pr, 'MESES_ESPANOL') else f"{spec.mes:02d}_{spec.anio}"
    nombre_consolidado = f"Consolidado_PT_{periodo_nom}.csv"
    nombre_output_involves = f"Involves_Procesado_{periodo_nom}.csv"

    ruta_bases_sharepoint = getattr(paths, 'RUTA_CARPETA_BASES_CIF', getattr(paths, 'RUTA_CARPETA_BASES', "Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS"))
    archivos_bases = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_bases_sharepoint)
    
    url_plan = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_consolidado in a["name"]), None)
    url_ventas = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if "ventas" in a["name"].lower()), None)
    url_causales = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if "causales" in a["name"].lower()), None)
    url_proc_inv = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_output_involves in a["name"]), None)
    
    ruta_involves_spec = str(pr.cif_involves(spec))
    nombre_inv_file = os.path.basename(ruta_involves_spec)
    url_inv = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_inv_file.lower() in a["name"].lower()), None)

    if not all([url_plan, url_ventas, url_causales, url_inv, url_proc_inv]):
        print("❌ ERROR: Faltan archivos requeridos en SharePoint para el Paso 4.")
        return

    df_causales_maestro = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_causales)))
    CAUSALES_CONFIG = dict(zip(
        df_causales_maestro['Motivo para no realizar la visita'].str.strip().str.upper(),
        df_causales_maestro['Validez']
    ))

    df_inv_raw = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_inv)), sheet_name="Sheet0")
    df_just = df_inv_raw[
        (df_inv_raw['Tipo de check-in'] == 'Sin check-in') & 
        (df_inv_raw['Motivo para no realizar la visita'].notnull())
    ].copy()

    if df_just.empty:
        print("⚠️ No hay motivos de inasistencia para procesar.")
        return

    columnas_just = {
        'Punto de venta': 'PUNTO_DE_VENTA', 
        'Empleado': 'EMPLEADO',
        'Fecha de la visita': 'FECHA_VISITA', 
        'Tipo de check-in': 'TIPO_CHECKIN',
        'Motivo para no realizar la visita': 'MOTIVO_REAL',
        'Justificación': 'OBSERVACION_JUSTIFICACION', 
        'Justificación auditada el': 'JUSTIFICADO_EL',
        'Quién justificó': 'QUIEN_JUSTIFICO'
    }
    df_just = df_just[[c for c in columnas_just.keys() if c in df_just.columns]].rename(columns=columnas_just)

    df_plan = pd.read_csv(io.BytesIO(descargar_archivo_bytes(url_plan)), sep=";", decimal=",", encoding="utf-8-sig")
    df_plan['ID_PDV_INVOLVES'] = pd.to_numeric(df_plan['ID_PDV_INVOLVES'], errors='coerce').fillna(0).astype(int).astype(str)
    
    df_maestro_info = df_plan[['NOMBRE', 'NOMBRE_PDV', 'ID_PDV_INVOLVES', 'CEDULA', 'ROL', 'HORAS_X_VISITA']].drop_duplicates()
    df_maestro_validador = df_plan[['NOMBRE', 'CEDULA']].drop_duplicates(subset=['NOMBRE'])

    df_maestro_ventas = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_ventas)))
    df_maestro_ventas['ID'] = pd.to_numeric(df_maestro_ventas['ID'], errors='coerce').fillna(0).astype(int).astype(str)
    df_maestro_ventas = df_maestro_ventas[['Nombre del PDV', 'ID']].drop_duplicates(subset=['Nombre del PDV'])

    df_just = df_just.merge(
        df_maestro_info, 
        left_on=['EMPLEADO', 'PUNTO_DE_VENTA'], 
        right_on=['NOMBRE', 'NOMBRE_PDV'], 
        how='left'
    )

    df_just = df_just.merge(df_maestro_ventas, left_on='PUNTO_DE_VENTA', right_on='Nombre del PDV', how='left')
    df_just['ID_FINAL'] = df_just['ID'].fillna(df_just['ID_PDV_INVOLVES'])

    df_just = df_just.merge(
        df_maestro_validador, 
        left_on='QUIEN_JUSTIFICO', 
        right_on='NOMBRE', 
        how='left', 
        suffixes=('', '_VAL')
    )
    
    df_just.rename(columns={'CEDULA_VAL': 'CEDULA_QUIEN_JUSTIFICO', 'ID_FINAL': 'ID_PDV_RESULTADO'}, inplace=True)
    cols_drop = ['NOMBRE_x', 'NOMBRE_y', 'NOMBRE_PDV', 'NOMBRE', 'ID', 'ID_PDV_INVOLVES', 'Nombre del PDV']
    df_just.drop(columns=[c for c in cols_drop if c in df_just.columns], inplace=True, errors='ignore')
    df_just.rename(columns={'ID_PDV_RESULTADO': 'ID_PDV_INVOLVES'}, inplace=True)

    df_just['VALIDEZ_JUSTIFICACION'] = df_just['MOTIVO_REAL'].apply(
        lambda x: CAUSALES_CONFIG.get(str(x).strip().upper(), 0)
    )
    
    df_just['TIEMPO_SERVICIO_MIN'] = np.where(
        df_just['VALIDEZ_JUSTIFICACION'] == 1,
        df_just['HORAS_X_VISITA'] * 60,
        0
    )
    df_just['VISITA_OK'] = df_just['TIEMPO_SERVICIO_MIN'] > 0

    df_paso3 = pd.read_csv(io.BytesIO(descargar_archivo_bytes(url_proc_inv)), sep=";", decimal=",", dtype={'ID_PDV_INVOLVES': str})
    
    df_final_total = pd.concat([df_paso3, df_just], ignore_index=True, sort=False)
    
    columnas_a_excluir = ['VALIDEZ_JUSTIFICACION', 'HORAS_X_VISITA']
    cols_finales = [c for c in df_final_total.columns if c not in columnas_a_excluir]
    
    csv_buffer = io.StringIO()
    df_final_total[cols_finales].drop_duplicates().to_csv(csv_buffer, index=False, sep=";", decimal=",", encoding="utf-8-sig")
    subir_archivo_a_sharepoint(headers, site_id, ruta_bases_sharepoint, nombre_output_involves, csv_buffer.getvalue().encode('utf-8-sig'), es_excel=False)
    print(f"✅ Paso 4 en nube finalizado.")


# ─────────────────────────────────────────────
# BLOQUE 5: INVOLVES AGRUPAR ARCHIVO INVOLVES
# ─────────────────────────────────────────────
def ejecutar_paso_5(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "="*40)
    print(" INICIANDO PASO 5: CONSOLIDACIÓN INVOLVES NUBE")
    print("="*40)
    
    periodo_nom = f"{pr.MESES_ESPANOL[spec.mes]}_{spec.anio}" if hasattr(pr, 'MESES_ESPANOL') else f"{spec.mes:02d}_{spec.anio}"
    nombre_output_involves = f"Involves_Procesado_{periodo_nom}.csv"
    nombre_agrupado_final = f"Involves_Agrupado_Final_{periodo_nom}.csv"

    ruta_bases_sharepoint = getattr(paths, 'RUTA_CARPETA_BASES_CIF', getattr(paths, 'RUTA_CARPETA_BASES', "Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS"))
    archivos_bases = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_bases_sharepoint)
    url_proc_inv = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_output_involves in a["name"]), None)

    if not url_proc_inv:
        print(f"❌ ERROR: No se encuentra el archivo {nombre_output_involves} en SharePoint.")
        return

    try:
        df = pd.read_csv(io.BytesIO(descargar_archivo_bytes(url_proc_inv)), sep=";", decimal=",", encoding="utf-8-sig")

        df['ID_PDV_INVOLVES'] = pd.to_numeric(df['ID_PDV_INVOLVES'], errors='coerce').fillna(0).astype(int).astype(str)
        df['EMPLEADO'] = df['EMPLEADO'].astype(str).str.strip()
        df['PERFIL_ACCESO'] = df['PERFIL_ACCESO'].fillna('').astype(str).str.upper().str.strip()

        mask_titular = (df['PERFIL_ACCESO'] == 'GESTOR') & (~df['EMPLEADO'].str.contains('SUPERNUMERARIO', case=False, na=False))
        df_titulares = df[mask_titular].sort_values('TIEMPO_SERVICIO_MIN', ascending=False)
        mapa_titulares = df_titulares.drop_duplicates('ID_PDV_INVOLVES').set_index('ID_PDV_INVOLVES')['EMPLEADO'].to_dict()

        es_backup         = df['PERFIL_ACCESO'] == 'BACKUP'
        es_supernumerario = df['EMPLEADO'].str.contains('SUPERNUMERARIO', case=False, na=False)
        es_apoyo          = es_backup | es_supernumerario

        titular_mapeado = df['ID_PDV_INVOLVES'].map(mapa_titulares)
        tiene_titular   = titular_mapeado.notna()

        reemplazar = es_apoyo & tiene_titular
        df['EMPLEADO_UNIFICADO'] = np.where(reemplazar, titular_mapeado, df['EMPLEADO'])
        df['PERFIL_UNIFICADO']   = np.where(reemplazar, 'GESTOR',        df['PERFIL_ACCESO'])

        df_final = df.groupby(['ID_PDV_INVOLVES', 'EMPLEADO_UNIFICADO', 'PERFIL_UNIFICADO']).agg({
            'PUNTO_DE_VENTA': 'first',
            'USUARIO_INVOLVES': 'first',
            'TIEMPO_SERVICIO_MIN': 'sum',
            'VISITA_OK': 'sum'
        }).reset_index()

        df_final = df_final[[
            'ID_PDV_INVOLVES', 'PUNTO_DE_VENTA', 'EMPLEADO_UNIFICADO', 
            'USUARIO_INVOLVES', 'PERFIL_UNIFICADO', 'TIEMPO_SERVICIO_MIN', 'VISITA_OK'
        ]]
        
        df_final.columns = [
            'ID_PDV_INVOLVES', 'PUNTO_DE_VENTA', 'EMPLEADO', 
            'USUARIO_INVOLVES', 'PERFIL_ACCESO', 'SUMA_TIEMPO_SERVICIO_MIN', 'CONTEO_VISITAS_OK'
        ]

        csv_buffer = io.StringIO()
        df_final.to_csv(csv_buffer, index=False, sep=";", decimal=",", encoding="utf-8-sig")
        subir_archivo_a_sharepoint(headers, site_id, ruta_bases_sharepoint, nombre_agrupado_final, csv_buffer.getvalue().encode('utf-8-sig'), es_excel=False)
        
        print(f"✅ ¡Proceso Limpio en Nube! IDs corregidos y guardados en SharePoint.")

    except Exception as e:
        print(f"❌ Error en Paso 5: {e}")


# ─────────────────────────────────────────────
# BLOQUE 6: ARCHIVO ALERTA CUANDO EL TIEMPO DE SERVICIO MAYOR A 540 MIN
# ─────────────────────────────────────────────
def ejecutar_paso_6(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "="*40)
    print(" INICIANDO PASO 6: GENERACIÓN DE ALERTAS NUBE")
    print("="*40)

    periodo_nom = f"{pr.MESES_ESPANOL[spec.mes]}_{spec.anio}" if hasattr(pr, 'MESES_ESPANOL') else f"{spec.mes:02d}_{spec.anio}"
    nombre_output_involves = f"Involves_Procesado_{periodo_nom}.csv"
    nombre_alerta = f"Alertas_Tiempos_{periodo_nom}.csv"

    ruta_bases_sharepoint = getattr(paths, 'RUTA_CARPETA_BASES_CIF', getattr(paths, 'RUTA_CARPETA_BASES', "Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS"))
    archivos_bases = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_bases_sharepoint)
    url_proc_inv = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_output_involves in a["name"]), None)

    if not url_proc_inv:
        print(f"❌ ERROR: No se encontró el archivo base en SharePoint.")
        return

    df = pd.read_csv(io.BytesIO(descargar_archivo_bytes(url_proc_inv)), sep=";", decimal=",", encoding="utf-8-sig")

    df['TIEMPO_NUM'] = pd.to_numeric(df['TIEMPO_SERVICIO_MIN'], errors='coerce').fillna(0)
    df_alertas = df[df['TIEMPO_NUM'] > 540].copy()

    if not df_alertas.empty:
        columnas_a_quitar = [
            'JUSTIFICACION', 'JUSTIFICADO_EL', 'QUIEN_JUSTIFICO', 
            'MOTIVO_REAL', 'OBSERVACION_JUSTIFICACION', 'CEDULA', 
            'ROL', 'NOMBRE_VAL', 'CEDULA_QUIEN_JUSTIFICO'
        ]
        
        df_alertas.drop(columns=[c for c in columnas_a_quitar if c in df_alertas.columns], inplace=True)
        if 'TIEMPO_NUM' in df_alertas.columns:
            df_alertas.drop(columns=['TIEMPO_NUM'], inplace=True)

        try:
            ruta_salidas_sharepoint = getattr(paths, 'RUTA_CARPETA_SALIDAS_CIF', getattr(paths, 'RUTA_CARPETA_SALIDAS', "Equipo Información/BI/INVOLVES/SALIDAS"))
            csv_buffer = io.StringIO()
            df_alertas.to_csv(csv_buffer, index=False, sep=";", decimal=",", encoding="utf-8-sig")
            subir_archivo_a_sharepoint(headers, site_id, ruta_salidas_sharepoint, nombre_alerta, csv_buffer.getvalue().encode('utf-8-sig'), es_excel=False)
            
            print(f"⚠️ AUDITORÍA: Se detectaron {len(df_alertas)} registros con tiempos > 9h.")
            print(f"📁 Archivo de alertas generado en SharePoint: {ruta_salidas_sharepoint}/{nombre_alerta}")
        except Exception as e:
            print(f"❌ ERROR al guardar alertas: {e}")
    else:
        print("✅ No se detectaron tiempos superiores a 9 horas.")

    print("="*40)


# ─────────────────────────────────────────────
# BLOQUE 7: ARCHIVO FINAL SALIDA CIF
# ─────────────────────────────────────────────
def ejecutar_paso_7(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 7: ESTRUCTURACIÓN FINAL NUBE  ({spec.etiqueta})")
    print("="*40)

    val_mes  = spec.mes
    val_anio = spec.anio
    periodo_nom = f"{pr.MESES_ESPANOL[spec.mes]}_{spec.anio}" if hasattr(pr, 'MESES_ESPANOL') else f"{spec.mes:02d}_{spec.anio}"

    nombre_agrupado_pt = f"Agrupado_PT_{periodo_nom}.csv"
    nombre_agrupado_inv = f"Involves_Agrupado_Final_{periodo_nom}.csv"

    ruta_bases_sharepoint = getattr(paths, 'RUTA_CARPETA_BASES_CIF', getattr(paths, 'RUTA_CARPETA_BASES', "Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS"))
    archivos_bases = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_bases_sharepoint)
    url_plan = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_agrupado_pt in a["name"]), None)
    url_vis = next((a["@microsoft.graph.downloadUrl"] for a in archivos_bases if nombre_agrupado_inv in a["name"]), None)

    if not url_plan or not url_vis:
        print("❌ ERROR: Faltan archivos agrupados en SharePoint para el Paso 7.")
        return

    df_plan = pd.read_csv(io.BytesIO(descargar_archivo_bytes(url_plan)), sep=";", decimal=",", encoding="utf-8-sig")
    df_visitas = pd.read_csv(io.BytesIO(descargar_archivo_bytes(url_vis)), sep=";", decimal=",", encoding="utf-8-sig")

    for df_temp, col in [(df_visitas, 'ID_PDV_INVOLVES'), (df_visitas, 'USUARIO_INVOLVES'),
                         (df_plan, 'ID_PDV_INVOLVES'), (df_plan, 'CEDULA')]:
        if col in df_temp.columns:
            df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0).astype(int).astype(str).str.strip()

    import re as _re
    def _norm_nombre(s):
        return _re.sub(r'\s+', ' ', str(s)).strip().upper()

    df_plan['_NOMBRE_KEY']    = df_plan['NOMBRE'].map(_norm_nombre)
    df_visitas['_NOMBRE_KEY'] = df_visitas['EMPLEADO'].map(_norm_nombre)

    def _first_non_null(s):
        s = s.dropna()
        return s.iloc[0] if len(s) else None

    df_visitas_unicas_nombre = df_visitas.groupby(
        ['ID_PDV_INVOLVES', '_NOMBRE_KEY'], as_index=False, dropna=False
    ).agg(
        PUNTO_DE_VENTA=('PUNTO_DE_VENTA', 'first'),
        USUARIO_INVOLVES=('USUARIO_INVOLVES', _first_non_null),
        PERFIL_ACCESO=('PERFIL_ACCESO', _first_non_null),
        SUMA_TIEMPO_SERVICIO_MIN=('SUMA_TIEMPO_SERVICIO_MIN', 'sum'),
        CONTEO_VISITAS_OK=('CONTEO_VISITAS_OK', 'sum'),
    )

    df_via_nombre = df_plan.merge(
        df_visitas_unicas_nombre,
        on=['ID_PDV_INVOLVES', '_NOMBRE_KEY'],
        how='left',
        suffixes=('', '_VIS'),
    )

    cols_visitas_solo = [c for c in df_visitas.columns if c not in ('ID_PDV_INVOLVES', '_NOMBRE_KEY')]
    mask_no_match = df_via_nombre['SUMA_TIEMPO_SERVICIO_MIN'].isna()

    if mask_no_match.any():
        df_visitas_unicas_cedula = df_visitas.groupby(
            ['ID_PDV_INVOLVES', 'USUARIO_INVOLVES'], as_index=False, dropna=False
        ).agg(
            PUNTO_DE_VENTA=('PUNTO_DE_VENTA', 'first'),
            EMPLEADO=('EMPLEADO', _first_non_null),
            _NOMBRE_KEY=('_NOMBRE_KEY', _first_non_null),
            PERFIL_ACCESO=('PERFIL_ACCESO', _first_non_null),
            SUMA_TIEMPO_SERVICIO_MIN=('SUMA_TIEMPO_SERVICIO_MIN', 'sum'),
            CONTEO_VISITAS_OK=('CONTEO_VISITAS_OK', 'sum'),
        )
        plan_rescate = df_via_nombre.loc[mask_no_match, df_plan.columns.tolist()].copy()
        rescate = plan_rescate.merge(
            df_visitas_unicas_cedula,
            left_on=['ID_PDV_INVOLVES', 'CEDULA'],
            right_on=['ID_PDV_INVOLVES', 'USUARIO_INVOLVES'],
            how='left',
            suffixes=('', '_VIS'),
        )
        rescate_reset = rescate.reset_index(drop=True)
        mask_indices = df_via_nombre.index[mask_no_match]
        for col in cols_visitas_solo:
            if col in rescate_reset.columns:
                df_via_nombre.loc[mask_indices, col] = rescate_reset[col].values

    df_final = df_via_nombre
    df_final.drop(columns=['_NOMBRE_KEY'], inplace=True, errors='ignore')

    df_final['SUMA_TIEMPO_SERVICIO_MIN'] = df_final['SUMA_TIEMPO_SERVICIO_MIN'].fillna(0)
    df_final['CONTEO_VISITAS_OK'] = df_final['CONTEO_VISITAS_OK'].fillna(0)
    df_final['SUMA_TIEMPO_SERVICIO_HORAS_REAL'] = df_final['SUMA_TIEMPO_SERVICIO_MIN'] / 60

    columnas_nuevas = {
        'TIEMPO_SERVICIO': 'TIEMPO_SERVICIO_PLANEADO',
        'TIEMPO_SERVICIO_HORAS': 'TIEMPO_SERVICIO_HORAS_PLANEADO',
        'FREC_SEMANAL': 'FREC_SEMANAL_PLANEADO',
        'FREC_MENSUAL': 'FREC_MENSUAL_PLANEADO',
        'HORAS_X_VISITA': 'HORAS_X_VISITA_PLANEADO',
        'HORAS_MES': 'HORAS_MES_PLANEADO',
        'SUMA_TIEMPO_SERVICIO_MIN': 'SUMA_TIEMPO_SERVICIO_MIN_REAL',
        'CONTEO_VISITAS_OK': 'VISITAS_REAL'
    }
    df_final.rename(columns=columnas_nuevas, inplace=True)

    df_final['COBERTURA'] = (df_final['VISITAS_REAL'] >= 1).astype(int)

    df_final['INTENSIDAD'] = np.where(
        df_final['HORAS_MES_PLANEADO'].astype(float) > 0,
        df_final['SUMA_TIEMPO_SERVICIO_HORAS_REAL'] / df_final['HORAS_MES_PLANEADO'],
        0.0,
    )

    mask_gestor = df_final['ROL'].astype(str).str.upper().str.contains('GESTOR', na=False)
    df_final['INTENSIDAD GESTORES'] = 0.0
    df_final.loc[mask_gestor, 'INTENSIDAD GESTORES'] = np.where(
        df_final.loc[mask_gestor, 'TIEMPO_SERVICIO_HORAS_PLANEADO'].astype(float) > 0,
        df_final.loc[mask_gestor, 'SUMA_TIEMPO_SERVICIO_HORAS_REAL']
            / df_final.loc[mask_gestor, 'TIEMPO_SERVICIO_HORAS_PLANEADO'],
        0.0,
    )

    den_frec = (
        df_final['FREC_SEMANAL_PLANEADO'].astype(float)
        * df_final['FREC_MENSUAL_PLANEADO'].astype(float)
    )
    df_final['FRECUENCIA'] = np.where(
        den_frec > 0, df_final['VISITAS_REAL'] / den_frec, 0.0
    )

    df_final['FRECUENCIA GESTORES'] = 0.0
    df_final.loc[mask_gestor, 'FRECUENCIA GESTORES'] = np.where(
        df_final.loc[mask_gestor, 'CANTIDAD_VISITAS'].astype(float) > 0,
        df_final.loc[mask_gestor, 'VISITAS_REAL']
            / df_final.loc[mask_gestor, 'CANTIDAD_VISITAS'],
        0.0,
    )

    df_final['MES'] = val_mes
    df_final['AÑO'] = val_anio

    columnas_ordenadas = [
        'ID_PDV_INVOLVES', 'ACRONIMO', 'NOMBRE_PDV', 'VENTAS_PROMEDIO_MES',
        'CEDULA', 'NOMBRE', 'COD_MERCADERISTA', 'CANTIDAD_VISITAS',
        'TIEMPO_SERVICIO_PLANEADO', 'TIEMPO_SERVICIO_HORAS_PLANEADO',
        'FREC_SEMANAL_PLANEADO', 'FREC_MENSUAL_PLANEADO',
        'HORAS_X_VISITA_PLANEADO', 'HORAS_MES_PLANEADO',
        'SUMA_TIEMPO_SERVICIO_MIN_REAL', 'SUMA_TIEMPO_SERVICIO_HORAS_REAL',
        'VISITAS_REAL',
        'COBERTURA', 'INTENSIDAD', 'INTENSIDAD GESTORES',
        'FRECUENCIA', 'FRECUENCIA GESTORES',
        'ROL', 'SUPERVISOR_LIDER', 'FUENTE', 'MES', 'AÑO',
    ]

    df_resultado = df_final[[c for c in columnas_ordenadas if c in df_final.columns]].copy()

    # Manejo de histórico en SharePoint para UPSERT multi-periodo
    ruta_salidas_sharepoint = getattr(paths, 'RUTA_CARPETA_SALIDAS_CIF', getattr(paths, 'RUTA_CARPETA_SALIDAS', "Equipo Información/BI/INVOLVES/SALIDAS"))
    archivos_salidas = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_salidas_sharepoint)
    nombre_cif_out = os.path.basename(str(paths.CIF_OUT_CIF))
    url_cif_prev = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salidas if nombre_cif_out.lower() in a["name"].lower()), None)

    if url_cif_prev:
        try:
            df_prev = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_cif_prev)))
            if "MES" in df_prev.columns and "AÑO" in df_prev.columns:
                periodos_actuales = set(
                    (int(m), int(a)) for m, a in zip(df_resultado["MES"], df_resultado["AÑO"])
                    if pd.notna(m) and pd.notna(a)
                )
                df_prev["_mes_int"]  = pd.to_numeric(df_prev["MES"],  errors="coerce")
                df_prev["_anio_int"] = pd.to_numeric(df_prev["AÑO"], errors="coerce")
                mask_keep = ~df_prev.apply(
                    lambda r: (r["_mes_int"], r["_anio_int"]) in periodos_actuales,
                    axis=1,
                )
                df_prev = df_prev[mask_keep].drop(columns=["_mes_int", "_anio_int"])
                df_resultado = pd.concat([df_prev, df_resultado], ignore_index=True, sort=False)
        except Exception as e:
            print(f"⚠️ No se pudo leer archivo previo en SharePoint ({e}); se sobrescribe.")

    subir_archivo_a_sharepoint(headers, site_id, ruta_salidas_sharepoint, nombre_cif_out, df_resultado, es_excel=True)
    print(f"✅ ¡LISTO! Excel generado y subido a SharePoint para el periodo {val_mes}/{val_anio}.")
    print("="*40)


# ─────────────────────────────────────────────
# BLOQUE 8 (V3): RESUMEN KPIS POR GESTOR Y CARGA SUPABASE
# ─────────────────────────────────────────────
def generar_resumen_kpis_8(spec: pr.PeriodoSpec, headers, site_id):
    print("\n" + "=" * 40)
    print(f" INICIANDO PASO 8: RESUMEN KPIS POR GESTOR (V3) NUBE  ({spec.etiqueta})")
    print("=" * 40)

    ruta_salidas_sharepoint = getattr(paths, 'RUTA_CARPETA_SALIDAS_CIF', getattr(paths, 'RUTA_CARPETA_SALIDAS', "Equipo Información/BI/INVOLVES/SALIDAS"))
    archivos_salidas = obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_salidas_sharepoint)
    nombre_cif_out = os.path.basename(str(paths.CIF_OUT_CIF))
    url_cif = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salidas if nombre_cif_out.lower() in a["name"].lower()), None)

    if not url_cif:
        print(f"❌ No existe el detalle CIF en SharePoint.")
        return

    df = pd.read_excel(io.BytesIO(descargar_archivo_bytes(url_cif)), engine='openpyxl')

    if 'MES' in df.columns and 'AÑO' in df.columns:
        m_int = pd.to_numeric(df['MES'], errors='coerce')
        a_int = pd.to_numeric(df['AÑO'], errors='coerce')
        mask_periodo = (m_int == spec.mes) & (a_int == spec.anio)
        if not mask_periodo.any():
            raise ValueError(
                f"Paso 8: el detalle CIF no contiene filas para el "
                f"periodo {spec.mes}/{spec.anio}. Re-ejecuta el paso 7."
            )
        df = df[mask_periodo].copy()

    def _limpiar(serie: pd.Series) -> pd.Series:
        s = serie.astype(str).str.strip().str.replace('%', '').str.replace(',', '.')
        n = pd.to_numeric(s, errors='coerce')
        if n.dropna().mean() > 1.5:
            n = n / 100
        return n.fillna(0)

    df['VENTAS_PROMEDIO_MES'] = pd.to_numeric(
        df['VENTAS_PROMEDIO_MES'].astype(str).str.replace(',', '.'),
        errors='coerce',
    ).fillna(0)
    df['COBERTURA']  = _limpiar(df['COBERTURA'])
    df['INTENSIDAD'] = _limpiar(df['INTENSIDAD'])
    df['FRECUENCIA'] = _limpiar(df['FRECUENCIA'])
    df['MES']  = pd.to_numeric(df.get('MES'),  errors='coerce').fillna(0).astype(int)
    df['AÑO']  = pd.to_numeric(df.get('AÑO'),  errors='coerce').fillna(0).astype(int)

    df['_prod_cob'] = df['VENTAS_PROMEDIO_MES'] * df['COBERTURA']
    df['_prod_int'] = df['VENTAS_PROMEDIO_MES'] * df['INTENSIDAD']
    df['_prod_fre'] = df['VENTAS_PROMEDIO_MES'] * df['FRECUENCIA']

    def _supervisor_dominante(s):
        s = s.dropna().astype(str).str.strip()
        s = s[(s != '') & (s != '0')]
        if s.empty:
            return ''
        return s.mode().iloc[0]

    resumen = (
        df.groupby(['MES', 'AÑO', 'NOMBRE'], dropna=False)
          .agg(SUPERVISOR_LIDER=('SUPERVISOR_LIDER', _supervisor_dominante),
               VENTAS_PROMEDIO_MES=('VENTAS_PROMEDIO_MES', 'sum'),
               prod_cob=('_prod_cob', 'sum'),
               prod_int=('_prod_int', 'sum'),
               prod_fre=('_prod_fre', 'sum'))
          .reset_index()
    )

    resumen['CUMPLIMIENTO COBERTURA']  = np.where(
        resumen['VENTAS_PROMEDIO_MES'] != 0,
        resumen['prod_cob'] / resumen['VENTAS_PROMEDIO_MES'], 0)
    resumen['CUMPLIMIENTO INTENSIDAD'] = np.where(
        resumen['VENTAS_PROMEDIO_MES'] != 0,
        resumen['prod_int'] / resumen['VENTAS_PROMEDIO_MES'], 0)
    resumen['CUMPLIMIENTO FRECUENCIA'] = np.where(
        resumen['VENTAS_PROMEDIO_MES'] != 0,
        resumen['prod_fre'] / resumen['VENTAS_PROMEDIO_MES'], 0)

    resumen['TOTAL'] = (
        resumen['CUMPLIMIENTO COBERTURA']  * 0.5
        + resumen['CUMPLIMIENTO INTENSIDAD'] * 0.2
        + resumen['CUMPLIMIENTO FRECUENCIA'] * 0.3
    )

    cols_final = ['MES', 'AÑO', 'SUPERVISOR_LIDER', 'NOMBRE',
                  'VENTAS_PROMEDIO_MES',
                  'CUMPLIMIENTO COBERTURA', 'CUMPLIMIENTO INTENSIDAD',
                  'CUMPLIMIENTO FRECUENCIA', 'TOTAL']
    resumen_final = resumen[cols_final].sort_values(
        ['AÑO', 'MES', 'SUPERVISOR_LIDER', 'NOMBRE']
    ).reset_index(drop=True)

    if not resumen_final.empty:
        anio_max = int(resumen_final['AÑO'].max())
        mes_max  = int(resumen_final[resumen_final['AÑO'] == anio_max]['MES'].max())
        del_periodo_activo = (
            (resumen_final['AÑO'] == anio_max) & (resumen_final['MES'] == mes_max)
        )
        resumen_activo = resumen_final[del_periodo_activo].copy()
    else:
        anio_max, mes_max = 0, 0
        resumen_activo = resumen_final.copy()

    nombre_kpi_out = os.path.basename(str(paths.CIF_OUT_KPIS))
    subir_archivo_a_sharepoint(headers, site_id, ruta_salidas_sharepoint, nombre_kpi_out, resumen_activo, es_excel=True)

    # ── CARGA A SUPABASE (Con normalización automática de columnas) ──
    try:
        if not resumen_activo.empty:
            print(f" 🚀 Preparando cargue de KPIs CIF a Supabase...")
            df_kpi_sup = resumen_activo.copy()
            
            # Normalización estricta de columnas respetando la estructura original de Supabase
            df_kpi_sup.columns = (
                df_kpi_sup.columns.str.strip()
                .str.lower()
                .str.replace(' ', '_')
            )
            
            # Aseguramos que se mantenga como 'año' si la tabla de Supabase lo requiere con eñe
            if 'anio' in df_kpi_sup.columns:
                df_kpi_sup = df_kpi_sup.rename(columns={'anio': 'año'})
            
            df_kpi_sup = df_kpi_sup.replace([np.inf, -np.inf], 0)
            df_kpi_limpio = df_kpi_sup.astype(object).where(pd.notnull(df_kpi_sup), None)
            
            registros_json = df_kpi_limpio.to_dict(orient="records")
            
            # Nombre de tabla Supabase adaptado para CIF KPIs (ej. 'cif_kpis')
            supabase.table("cif_kpis").upsert(registros_json).execute()
            print(f"  ✅ ÉXITO CLOUD: {len(registros_json)} filas subidas exitosamente a Supabase.")
    except Exception as ex_cloud:
        print(f"  ⚠ ADVERTENCIA EN CARGA CLOUD (Supabase): {ex_cloud}")

    print("=" * 40)


# ─────────────────────────────────────────────
# CONTROL DE EJECUCIÓN NUBE
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="ETL CIF NUBE — Eficacia (Microsoft Graph + Supabase)")
    parser.add_argument("--solo", nargs="+",
                        choices=["consolidacion", "agrupacion", "visitas", "visitas2",
                                  "agrupacion_visitas", "alerta", "final", "kpis"],
                        help="Correr solo los pasos indicados")
    parser.add_argument("--full", action="store_true",
                        help="Correr la cadena completa en orden estricto")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()

    specs = pr.periodos_de_args(args)

    # Orden estricto y lógico de la ETL
    orden_estricto = [
        "consolidacion", 
        "agrupacion", 
        "visitas", 
        "visitas2",
        "agrupacion_visitas", 
        "alerta", 
        "final", 
        "kpis"
    ]

    if args.full or not args.solo:
        pasos = orden_estricto
    else:
        # Respeta el orden estricto aunque el usuario elija subconjuntos
        pasos = [p for p in orden_estricto if p in args.solo]

    print("🔒 Autenticando con Microsoft Graph API...")
    token = obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = obtener_site_id(headers)

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL CIF NUBE — procesando periodo {spec.etiqueta} ({spec})")

        try:
            if "consolidacion" in pasos:        ejecutar_paso_1(spec, headers, site_id)
            if "agrupacion" in pasos:           ejecutar_paso_2(spec, headers, site_id)
            if "visitas" in pasos:              ejecutar_paso_3(spec, headers, site_id)
            if "visitas2" in pasos:             ejecutar_paso_4(spec, headers, site_id)
            if "agrupacion_visitas" in pasos:   ejecutar_paso_5(spec, headers, site_id)
            if "alerta" in pasos:               ejecutar_paso_6(spec, headers, site_id)
            if "final" in pasos:                ejecutar_paso_7(spec, headers, site_id)
            if "kpis" in pasos:                 generar_resumen_kpis_8(spec, headers, site_id)

            print(f"\n🏁 PROCESO TERMINADO EN NUBE  ({spec.etiqueta}).")

        except Exception as e:
            print(f"\n❌ OCURRIÓ UN ERROR CRÍTICO en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise