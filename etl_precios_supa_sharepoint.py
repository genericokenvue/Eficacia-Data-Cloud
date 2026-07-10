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

# --- LIBRERÍAS REQUERIDAS PARA SUPABASE ---
from supabase import create_client

# --- CARGAR CREDENCIALES DESDE EL ARCHIVO EXTERNO .ENV (Solo Local) ---
load_dotenv()

# =============================================================================
# 1. CONFIGURACIÓN GLOBAL & AZURE API (SISTEMA DE ENTORNOS INTEGRADO)
# =============================================================================

# 🔑 Credenciales mapeadas de manera segura:
TENANT_ID = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")

# 🗄️ Conexión de Supabase:
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

# Validación estricta antes de iniciar el flujo
if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET, SUPABASE_URL, SUPABASE_KEY]):
    raise ValueError("❌ ERROR: Faltan credenciales esenciales. Verifica tu archivo .env o tus GitHub Secrets.")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# 📂 Configuración del SharePoint de la empresa:
SHAREPOINT_SITE_NAME = "JJ451"  
RUTA_CARPETA_PT = "Equipo Información/BI/INVOLVES/Plan_De_Trabajo"
RUTA_CARPETA_BASES = "Equipo Información/BI/INVOLVES/Precios_Bases"
RUTA_CARPETA_SALIDAS = "Equipo Información/BI/INVOLVES/Salidas"

LISTA_CATEGORIAS_PRECIOS = [
    "PROTECCION FEMENINA", "JABONES DE TOCADOR", "ASEO DEL BEBE",
    "CREMAS CORPORALES", "CUIDADO FACIAL", "ENJUAGUE BUCAL", "CREMAS DENTALES"
]

MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# --- RUTAS locales o resueltas para archivos temporales/históricos ---
import paths
import periodo_resolver as pr

# Patrón relajado para soportar "Respuestas de encuestas..." y "Respuestas Precios"
CLAVE_ARCHIVO_ENCUESTA = "Respuestas"

# =============================================================================
# FUNCIONES AUXILIARES DE MICROSOFT GRAPH API (SHAREPOINT CONNECT)
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
    url = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{SHAREPOINT_SITE_NAME}"
    res = requests.get(url, headers=headers).json()
    if "id" not in res:
        raise Exception(f"No se encontró el sitio SharePoint '{SHAREPOINT_SITE_NAME}'.")
    return res["id"]

def obtener_archivos_carpeta_sharepoint(headers, site_id, ruta_carpeta):
    ruta_formateada = urllib.parse.quote(ruta_carpeta)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_formateada}:/children"
    
    response = requests.get(url, headers=headers)
    res_json = response.json()
    
    if response.status_code != 200:
        print(f"❌ Error de la API de Graph ({response.status_code}): {res_json.get('error', {}).get('message')}")
        print("\n🔍 Listando carpetas disponibles en la raíz del sitio para verificar nombres:")
        url_raiz = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        raiz_json = requests.get(url_raiz, headers=headers).json()
        for item in raiz_json.get('value', []):
            print(f" -> [{item.get('name')}] (Tipo: {item.get('folder', 'Archivo')})")
    
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
# 2. FUNCIONES DE PROCESAMIENTO
# =============================================================================

def ejecutar_paso_1_consolidar_pt(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 1: Consolidando Plan de Trabajo  ({spec.etiqueta}) ---")

    from shared_loader import COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR

    COLUMNAS_FINALES_PT = [
        "ID_PDV_INVOLVES", "NOMBRE_PDV", "VENTAS_PROMEDIO_MES", "ACRONIMO",
        "CEDULA", "NOMBRE", "COD_MERCADERISTA", "ROL", "SUPERVISOR_LIDER", "FUENTE",
        "OBSERVACION",
    ]

    OBS_A_ROL = {
        "REPORTA GESTOR":              "GESTOR",
        "REPORTA SUPERVISOR":          "SUPERVISOR",
        "REPORTA GENERADOR DE DEMANDA":"GENERADOR DE DEMANDA",
    }

    archivos_en_carpeta = obtener_archivos_carpeta_sharepoint(headers, site_id, RUTA_CARPETA_PT)
    
    num_mes = f"{spec.mes:02d}"                  
    nombre_mes = MESES_ESPANOL[spec.mes].upper()  

    url_directo = None
    url_ism = None

    for archivo in archivos_en_carpeta:
        nombre_archivo_upper = archivo["name"].upper()
        
        if "DIRECTO" in nombre_archivo_upper and (num_mes in nombre_archivo_upper or nombre_mes in nombre_archivo_upper):
            url_directo = archivo["@microsoft.graph.downloadUrl"]
            print(f"  📂 Archivo DIRECTO detectado en Cloud: {archivo['name']}")
            
        if "ISM" in nombre_archivo_upper and (num_mes in nombre_archivo_upper or nombre_mes in nombre_archivo_upper):
            url_ism = archivo["@microsoft.graph.downloadUrl"]
            print(f"  📂 Archivo ISM detectado en Cloud: {archivo['name']}")

    if not url_directo:
        url_directo = next((a["@microsoft.graph.downloadUrl"] for a in archivos_en_carpeta if "DIRECTO" in a["name"].upper()), None)
    if not url_ism:
        url_ism = next((a["@microsoft.graph.downloadUrl"] for a in archivos_en_carpeta if "ISM" in a["name"].upper()), None)

    if not url_directo or not url_ism:
        raise FileNotFoundError(f"No se localizaron los archivos base en la carpeta cloud {RUTA_CARPETA_PT}")

    def leer_y_normalizar_cloud(url_download, hoja, fuente):
        content = requests.get(url_download).content
        with pd.ExcelFile(io.BytesIO(content)) as xls:
            df_val = pd.read_excel(xls, sheet_name="Captura de modulos")
            df_val.columns = df_val.columns.str.strip().str.upper()
            col_filtro = "PRECIOS_FINAL" if fuente == "ISM" else "PRECIOS"
            if col_filtro not in df_val.columns: return pd.DataFrame()

            obs_col = next((c for c in df_val.columns if c.startswith('OBSERVAC')), None)
            if obs_col is None:
                print(f"⚠ {fuente}: hoja Captura de modulos sin columna OBSERVACIÓN")
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

            pdvs_obs = df_val_filtrado[["ID PDV INVOLVES", obs_col]].rename(columns={obs_col: "OBSERVACION"})
            pdvs_obs["OBSERVACION"] = pdvs_obs["OBSERVACION"].astype(str).str.strip().str.upper()
            pdvs_obs["ROL_ESPERADO"] = pdvs_obs["OBSERVACION"].map(OBS_A_ROL).fillna("")

            df = pd.read_excel(xls, sheet_name=hoja)
            df.columns = df.columns.str.strip().str.upper()

            df = df[df["ID PDV INVOLVES"].isin(pdvs_obs["ID PDV INVOLVES"])].copy()
            df.rename(columns={col: COLUMNAS_ESTANDAR[col] for col in df.columns if col in COLUMNAS_ESTANDAR}, inplace=True)

            df = df.merge(pdvs_obs, left_on="ID_PDV_INVOLVES", right_on="ID PDV INVOLVES", how="left")

            df["ROL"] = df["ROL"].astype(str).str.strip().str.upper()
            mask_responsable = df["ROL"] == df["ROL_ESPERADO"]
            n_antes = len(df)
            df = df[mask_responsable].copy()
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

        df_total_unificado = df_total.drop_duplicates(
            subset=['ID_PDV_INVOLVES', 'NOMBRE'],
        ).reset_index(drop=True)

        nombre_pt_final = f"Plan_Trabajo_Precios_{periodo}.xlsx"
        subir_archivo_a_sharepoint(headers, site_id, RUTA_CARPETA_BASES, nombre_pt_final, df_total_unificado)

        print(f"✅ EXITO: Plan unificado guardado en SharePoint: {nombre_pt_final}")
        return df_total_unificado, periodo
    
    return pd.DataFrame(), periodo

def generar_reporte_captura_precios(df_pt, periodo_nom, spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 2: Cruzando Capturas por Categorías ({periodo_nom}) ---")
    
    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, RUTA_CARPETA_BASES)
    url_encuesta = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_ENCUESTA in a["name"] and periodo_nom in a["name"].upper()), None)
    
    if not url_encuesta:
        url_encuesta = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_ENCUESTA in a["name"]), None)
        
    if not url_encuesta:
        raise FileNotFoundError(f"No se encontró el archivo cloud de encuestas en: {RUTA_CARPETA_BASES}")

    content_enc = requests.get(url_encuesta).content
    df_enc = pd.read_excel(io.BytesIO(content_enc))

    def extraer_cat(texto):
        if pd.isna(texto): return None
        t = re.sub(r'[^A-Z0-9\s]', ' ', str(texto).upper())
        t = " ".join(t.split())
        for cat in LISTA_CATEGORIAS_PRECIOS:
            if cat in t: return cat
        return None

    df_enc['CATEGORIA_LIMPIA'] = df_enc['Rótulo de la encuesta'].apply(extraer_cat)
    df_enc_valida = df_enc.dropna(subset=['CATEGORIA_LIMPIA']).copy()
    
    from shared_loader import id_a_str
    df_pt['ID_PDV_INVOLVES'] = id_a_str(df_pt['ID_PDV_INVOLVES'])
    df_enc_valida['ID del PDV'] = id_a_str(df_enc_valida['ID del PDV'])

    df_enc_valida['ROTULO_UPPER'] = df_enc_valida['Rótulo de la encuesta'].astype(str).str.upper()
    
    pdvs_mayoristas = df_enc_valida.groupby('ID del PDV')['ROTULO_UPPER'].apply(
        lambda x: any('MAYORISTA' in r for r in x)
    ).to_dict()

    dict_enc = df_enc_valida.groupby('ID del PDV')['CATEGORIA_LIMPIA'].unique().to_dict()
    set_oficial_estandar = set(LISTA_CATEGORIAS_PRECIOS)
    res = []

    for id_pdv in df_pt['ID_PDV_INVOLVES'].unique():
        cats = dict_enc.get(id_pdv, [])
        set_cats = set(cats)
        
        es_mayorista = pdvs_mayoristas.get(id_pdv, False)
        categorias_requeridas = 1 if es_mayorista else len(set_oficial_estandar)
        
        if es_mayorista:
            faltantes = [] if len(set_cats) >= 1 else ["AL MENOS 1 CATEGORÍA"]
            conteo = len(set_cats)
            msg_faltantes = "NINGUNA" if len(set_cats) >= 1 else "REQUIERE 1 CATEGORÍA"
        else:
            faltantes = sorted(list(set_oficial_estandar - set_cats))
            conteo = len(set_cats)
            msg_faltantes = ", ".join(faltantes) if faltantes else "NINGUNA"
            
        if conteo == 0: 
            msg_faltantes = f"SIN CAPTURAS (FALTAN {categorias_requeridas})"

        res.append({
            'ID_PDV_INVOLVES': id_pdv, 
            'CAPTURA_PLANEADA': 1, 
            'CONTEO_CATEGORIAS': conteo, 
            'CATEGORIAS_FALTANTES': msg_faltantes,
            'CAPTURA_EJECUTADA': 1 if conteo >= categorias_requeridas else 0
        })
    
    df_final = pd.merge(df_pt, pd.DataFrame(res), on='ID_PDV_INVOLVES', how='left')
    df_final['CAPTURA_PLANEADA'] = df_final['CAPTURA_PLANEADA'].fillna(1).astype(int)
    df_final['CAPTURA_EJECUTADA'] = df_final['CAPTURA_EJECUTADA'].fillna(0).astype(int)

    df_final['MES'] = spec.mes
    df_final['AÑO'] = spec.anio

    cols_out = list(df_pt.columns) + [
        'CAPTURA_PLANEADA', 'CONTEO_CATEGORIAS', 'CATEGORIAS_FALTANTES',
        'CAPTURA_EJECUTADA', 'MES', 'AÑO',
    ]
    
    nombre_matriz = f"REPORTE_CAPTURA_PRECIOS_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, RUTA_CARPETA_SALIDAS, nombre_matriz, df_final[cols_out])
    print(f"✅ Reporte capturas unificado generado en SharePoint.")

def generar_analisis_precios(periodo_nom, spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO 3: Generando Análisis Detallado ({periodo_nom}) ---")
    
    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, RUTA_CARPETA_BASES)
    url_encuesta = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_ENCUESTA in a["name"] and periodo_nom in a["name"].upper()), None)
    if not url_encuesta:
        url_encuesta = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if CLAVE_ARCHIVO_ENCUESTA in a["name"]), None)

    content_enc = requests.get(url_encuesta).content
    df = pd.read_excel(io.BytesIO(content_enc))
    mapeo_binario = {'SI': 1, 'SÍ': 1, 'NO': 0, 'no': 0, 'si': 1, 'sí': 1}

    if "Producto (SKU)" in df.columns:
        split_sku = df["Producto (SKU)"].astype(str).str.split(' - ', n=1, expand=True)
        df['CODIGO_SKU'] = split_sku[0].str.strip()
        df['NOMBRE_PRODUCTO'] = split_sku[1].str.strip() if split_sku.shape[1] > 1 else ""

    df['PRESENCIA'] = df["El producto está presente en el PDV?"].astype(str).str.strip().str.upper().map(mapeo_binario).fillna(0).astype(int)
    df['PRECIO_REGULAR'] = pd.to_numeric(df["Digite Precio Regular"], errors='coerce').fillna(0)
    df['LABEL_UBICADO'] = df["El producto cuenta con el Label (Precio impreso) ubicado?"].astype(str).str.strip().str.upper().map(mapeo_binario).fillna(0).astype(int)
    df['HAY_PROMO'] = df["Hay Promociones o Descuentos?"].astype(str).str.strip().str.upper().map(mapeo_binario).fillna(0).astype(int)
    df['PRECIO_PROMO'] = pd.to_numeric(df["Digite Precio Promoción:"], errors='coerce').fillna(0)
    df['TIPO_PROMO'] = df["Seleccione la promoción:"].fillna("SIN PROMOCION")

    cols_finales = ["ID del PDV", "PDV", "Fecha de la encuesta", "Empleado", "Marca", 
                    "CODIGO_SKU", "NOMBRE_PRODUCTO", "PRESENCIA", "PRECIO_REGULAR", 
                    "LABEL_UBICADO", "HAY_PROMO", "PRECIO_PROMO", "TIPO_PROMO"]
    
    nombre_analisis = f"ANALISIS_PRECIOS_{periodo_nom}.xlsx"
    subir_archivo_a_sharepoint(headers, site_id, RUTA_CARPETA_SALIDAS, nombre_analisis, df[cols_finales])
    print(f"✅ Análisis detallado guardado en SharePoint.")

# =============================================================================
# 3. EJECUCIÓN PRINCIPAL Y INTEGRACIÓN CON CLOUD
# =============================================================================

def generar_resumen_kpi_precios(spec: pr.PeriodoSpec, headers, site_id):
    print(f"\n--- PASO KPI (V3): RESUMEN PRECIOS por gestor  ({spec.etiqueta}) ---")
    import paths as _paths
    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    
    archivos_salida = obtener_archivos_carpeta_sharepoint(headers, site_id, RUTA_CARPETA_SALIDAS)
    url_reporte = next((a["@microsoft.graph.downloadUrl"] for a in archivos_salida if f"REPORTE_CAPTURA_PRECIOS_{periodo_nom}" in a["name"]), None)
    
    if not url_reporte:
        print(f"❌ No existe el reporte del periodo en SharePoint: REPORTE_CAPTURA_PRECIOS_{periodo_nom}")
        return
        
    content_rep = requests.get(url_reporte).content
    df = pd.read_excel(io.BytesIO(content_rep), engine='openpyxl')
    print(f"  Leyendo desde cloud: REPORTE_CAPTURA_PRECIOS_{periodo_nom}.xlsx ({len(df)} filas)")

    from shared_loader import calcular_kpi_simple_y_escribir
    n = calcular_kpi_simple_y_escribir(
        df_origen=df,
        col_planeado="CAPTURA_PLANEADA",
        col_ejecutado="CAPTURA_EJECUTADA",
        ruta_kpi_mes=str(_paths.PR_OUT_KPIS),
        ruta_kpi_historico=str(_paths.PR_OUT_KPIS_HISTORICO),
        nombre_cumplimiento="CUMPLIMIENTO",
        nombres_renombre=("PLANEADO", "EJECUTADO"),
    )
    print(f"  ✅ Mes activo: {n} gestores.")

    if os.path.exists(str(_paths.PR_OUT_KPIS)):
        df_kpi_local = pd.read_excel(str(_paths.PR_OUT_KPIS))
        subir_archivo_a_sharepoint(headers, site_id, RUTA_CARPETA_SALIDAS, _paths.PR_OUT_KPIS.name, df_kpi_local)

    try:
        if os.path.exists(str(_paths.PR_OUT_KPIS)):
            print(f"  🚀 Preparando cargue del KPI de Precios consolidado a Supabase...")
            df_kpi = pd.read_excel(str(_paths.PR_OUT_KPIS))
            
            df_kpi.columns = df_kpi.columns.str.strip().str.lower()
            if 'año' in df_kpi.columns:
                df_kpi = df_kpi.rename(columns={'año': 'anio'})
            
            # 🛠️ AJUSTE AQUÍ: Limpiar infinitos y NaNs antes de mapear a diccionarios JSON
            df_kpi.replace([np.inf, -np.inf], 0, inplace=True)
            df_kpi_limpio = df_kpi.where(pd.notnull(df_kpi), None)
            
            registros_json = df_kpi_limpio.to_dict(orient="records")
            
            supabase.table("precios_kpis").upsert(registros_json).execute()
            print(f"  ✅ ÉXITO CLOUD: {len(registros_json)} filas subidas de forma exitosa a Supabase.")
    except Exception as ex_cloud:
        print(f"  ⚠ ADVERTENCIA EN CARGA CLOUD (Los archivos Excel locales se generaron bien): {ex_cloud}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL PRECIOS — Eficacia (Sprint 16.1: multi-periodo a SharePoint con Rutas Quemadas)")
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

    if len(specs) > 1:
        print(f"🎯 ETL PRECIOS — multi-periodo: {len(specs)} meses → "
              f"{', '.join(s.etiqueta for s in specs)}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL PRECIOS — procesando periodo {spec.etiqueta} ({spec})")

        df_pt = pd.DataFrame()
        periodo_final = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"

        try:
            if "full" in pasos:
                df_pt, periodo_final = ejecutar_paso_1_consolidar_pt(spec, headers, site_id)

                if df_pt.empty:
                    archivos_origen = obtener_archivos_carpeta_sharepoint(headers, site_id, RUTA_CARPETA_BASES)
                    nombre_buscado = f"Plan_Trabajo_Precios_{periodo_final}.xlsx"
                    url_pt_periodo = next((a["@microsoft.graph.downloadUrl"] for a in archivos_origen if a["name"] == nombre_buscado), None)
                    
                    if url_pt_periodo:
                        content_pt = requests.get(url_pt_periodo).content
                        df_pt = pd.read_excel(io.BytesIO(content_pt))
                        print(f"📂 Cargado PT del periodo desde SharePoint: {nombre_buscado}")

                if not df_pt.empty:
                    generar_reporte_captura_precios(df_pt, periodo_final, spec, headers, site_id)
                generar_analisis_precios(periodo_final, spec, headers, site_id)

            if "kpi" in pasos:
                generar_resumen_kpi_precios(spec, headers, site_id)

            print(f"\n🏁 PROCESO TERMINADO  ({spec.etiqueta})")

        except Exception as e:
            print(f"\n❌ ERROR en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise