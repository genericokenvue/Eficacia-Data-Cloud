"""
etl_ventas.py
─────────────
Consolida los archivos Excel mensuales de **Ventas D&P** en un único CSV
(`Consolidado_Ventas.csv`) que sirve como data lake interno para Fase 2.

Origen
──────
    paths.DYP_VENTAS_DIR/<año>/<mes>/<archivo>.xlsx

Filtros
───────
  • Sólo años 2024 en adelante (carpetas con nombre "2023" o anterior se omiten).
  • Subcarpetas que contengan "MES ACTUAL" en cualquier nivel se omiten
    (las usa el área para preparar el corte parcial del mes en curso).

Salida
──────
    paths.DYP_OUT_VENTAS  (CSV con delimitador `|`, decimal `,`)

Modo upsert
───────────
  • Si el CSV ya existe, sólo se reprocesa el **último archivo** detectado
    (corte más reciente del mes vigente) y se sobreescriben sus filas
    (MES, AÑO) en el consolidado conservando la historia.
  • Si el CSV no existe, se reconstruye desde cero leyendo todos los Excel.

Esquema esperado en el Excel origen
───────────────────────────────────
Columnas que se mapean a UPPER del consolidado:
  Bodega, Cod./Nombre Supervisor/Vendedor/Cliente, Estado PDM, Profit,
  Categoria/Subcategoria/Marca/Linea, Cod. EAN Producto, Producto, Tipologia,
  Cant./Vta. <Mes>, Cantidades Totales, Ventas Totales,
  Mes/MES, Año/AÑO.
"""

from __future__ import annotations

import os
import re
import io
import urllib.parse
import requests
import msal
import warnings
import locale
import numpy as np
import pandas as pd

import paths
import periodo_resolver as pr

# --- LIBRERÍAS REQUERIDAS PARA SUPABASE ---
from supabase import create_client

warnings.filterwarnings("ignore", category=FutureWarning)

try:
    locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, "Spanish")
    except locale.Error:
        pass

from dotenv import load_dotenv
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

COLUMN_MAPPING_ORIGEN_A_DESTINO = {
    "Bodega": "Bodega",
    "Cod. Supervisor": "Cod. Supervisor", "Supervisor": "Supervisor",
    "Cod. Vendedor":   "Cod. Vendedor",   "Vendedor": "Vendedor",
    "Cod. Cliente":    "Cod. Cliente",    "Cliente": "Cliente",
    "Multicategoria Decil": "Multicategoria Decil",
    "Listerine Decil":      "Listerine Decil",
    "Lubriderm Decil":      "Lubriderm Decil",
    "NTG Decil":            "NTG Decil",
    "Johnson s Decil":      "Johnson s Decil",
    "Estado PDM": "Estado PDM",
    "Profit": "Profit",
    "Categoria": "Categoria", "Subcategoria": "Subcategoria",
    "Marca": "Marca", "Linea": "Linea",
    "Cod. EAN Producto": "Cod. EAN Producto", "Producto": "Producto",
    "Tipologia": "Tipologia",
    "Cant. Enero": "Cant. Enero",         "Vta. Enero":      "Vta. Enero",
    "Cant. Febrero": "Cant. Febrero",     "Vta. Febrero":    "Vta. Febrero",
    "Cant. Marzo": "Cant. Marzo",         "Vta. Marzo":      "Vta. Marzo",
    "Cant. Abril": "Cant. Abril",         "Vta. Abril":      "Vta. Abril",
    "Cant. Mayo": "Cant. Mayo",           "Vta. Mayo":       "Vta. Mayo",
    "Cant. Junio": "Cant. Junio",         "Vta. Junio":      "Vta. Junio",
    "Cant. Julio": "Cant. Julio",         "Vta. Julio":      "Vta. Julio",
    "Cant. Agosto": "Cant. Agosto",       "Vta. Agosto":     "Vta. Agosto",
    "Cant. Septiembre": "Cant. Septiembre","Vta. Septiembre":"Vta. Septiembre",
    "Cant. Octubre": "Cant. Octubre",     "Vta. Octubre":    "Vta. Octubre",
    "Cant. Noviembre": "Cant. Noviembre", "Vta. Noviembre":  "Vta. Noviembre",
    "Cant. Diciembre": "Cant. Diciembre", "Vta. Diciembre":  "Vta. Diciembre",
    "Cantidades Totales.": "Cantidades Totales", "Cantidades Totales": "Cantidades Totales",
    "Ventas Totales.":     "Ventas Totales",     "Ventas Totales":     "Ventas Totales",
    "Mes": "MES", "mes": "MES",
    "Año": "AÑO", "año": "AÑO",
}

COLUMNS_TO_CLEAN_DOT_ZERO = [
    "Cod. Vendedor", "Cod. Supervisor", "Cod. Cliente", "Cod. EAN Producto",
    "Cantidades Totales",
    "Cant. Enero", "Cant. Febrero", "Cant. Marzo", "Cant. Abril",
    "Cant. Mayo", "Cant. Junio", "Cant. Julio", "Cant. Agosto",
    "Cant. Septiembre", "Cant. Octubre", "Cant. Noviembre", "Cant. Diciembre",
]

VTA_COLUMNS = [
    c for c in COLUMN_MAPPING_ORIGEN_A_DESTINO.values()
    if c.startswith("Vta.") or c == "Ventas Totales"
]

DELIMITADOR_CSV = "|"

MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# Archivos fijos / maestros conocidos que no tienen mes/año dinámico
ARCHIVOS_FIJOS_IGNORAR = [
    "base_cupos", "debug_base_cupos_normalizada", "msl & listas target catman", 
    "listas_referencia", "rutero_droguerias"
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

def obtener_archivos_carpeta_sharepoint_recursivo(headers, site_id, anio_proceso=2026):
    ruta_relativa = f"Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS/DYP/{anio_proceso}"
    print(f"  ☁️ Consultando ruta de origen en SharePoint: {ruta_relativa}")
    
    ruta_formateada = urllib.parse.quote(ruta_relativa)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_formateada}:/children"
    
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"❌ Error al listar carpeta en SharePoint ({response.status_code}): {response.text}")
        return []
    
    elementos = response.json().get("value", [])
    archivos_validos = []

    def recorrer_items(items, current_path):
        for item in items:
            nombre = item.get("name", "")
            if "folder" in item:
                sub_ruta_relativa = f"{current_path}/{nombre}"
                sub_ruta_formateada = urllib.parse.quote(sub_ruta_relativa)
                sub_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{sub_ruta_formateada}:/children"
                sub_res = requests.get(sub_url, headers=headers)
                if sub_res.status_code == 200:
                    recorrer_items(sub_res.json().get("value", []), f"{current_path}/{nombre}")
            elif "file" in item:
                ruta_completa = f"{current_path}/{nombre}"
                if "MES ACTUAL" in ruta_completa.upper():
                    continue
                if nombre.endswith(".xlsx") or nombre.endswith(".xls"):
                    archivos_validos.append({
                        "name": nombre,
                        "path": ruta_completa,
                        "@microsoft.graph.downloadUrl": item.get("@microsoft.graph.downloadUrl")
                    })

    recorrer_items(elementos, ruta_relativa)
    return sorted(archivos_validos, key=lambda x: x["path"])


# ─────────────────────────────────────────────────────────────────────────────
# DESCUBRIMIENTO DE ARCHIVOS DESDE CLOUD
# ─────────────────────────────────────────────────────────────────────────────

def _detectar_mes_anio_cloud(download_url: str, nombre_archivo: str, mapping: dict) -> tuple[str | None, str | None]:
    mes_val: str | None = None
    anio_val: str | None = None

    try:
        content = requests.get(download_url).content
        df_head = pd.read_excel(io.BytesIO(content), nrows=10)
        df_head.rename(columns=mapping, inplace=True)
        if "MES" in df_head.columns and not df_head["MES"].dropna().empty:
            mes_val = str(df_head["MES"].dropna().iloc[0]).strip().capitalize()
        if "AÑO" in df_head.columns and not df_head["AÑO"].dropna().empty:
            try:
                anio_val = str(int(df_head["AÑO"].dropna().iloc[0]))
            except Exception:
                anio_val = None
    except Exception:
        pass

    if not mes_val or not anio_val:
        m = re.search(r"([A-Za-z]+)\s(\d{4})", nombre_archivo)
        if m:
            if not mes_val:
                mes_val = m.group(1).capitalize()
            if not anio_val:
                anio_val = m.group(2)

    return mes_val, anio_val


# ─────────────────────────────────────────────────────────────────────────────
# 1. CARGA + CONSOLIDACIÓN EN LA NUBE (VENTAS)
# ─────────────────────────────────────────────────────────────────────────────

def consolidar_ventas(
    archivos_cloud: list[dict],
    spec: pr.PeriodoSpec,
    headers,
    site_id,
    mapping: dict = COLUMN_MAPPING_ORIGEN_A_DESTINO,
    delimitador: str = DELIMITADOR_CSV,
) -> int:
    archivos_ventas = []
    for f in archivos_cloud:
        nombre_lower = f["name"].lower()
        if any(fijo in nombre_lower for fijo in ARCHIVOS_FIJOS_IGNORAR):
            continue
        # Excluir explícitamente impactos y segmentos de ventas
        if "impacto" in nombre_lower or "segmento" in nombre_lower:
            continue
        if "venta" in nombre_lower or "rutero" in nombre_lower:
            archivos_ventas.append(f)

    if not archivos_ventas:
        print("❌ Sin archivos de Ventas/Ruteros para consolidar en la nube.")
        return 0

    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    nombre_consolidado_cloud = f"Consolidado_Ventas_{periodo_nom}.csv"
    ruta_salida_relativa = "Equipo Información/BI/INVOLVES/SALIDAS/DYP"
    
    archivos_salida = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{urllib.parse.quote(ruta_salida_relativa)}:/children",
        headers=headers
    )
    
    consolidado_existe = False
    url_consolidado_previo = None
    if archivos_salida.status_code == 200:
        for item in archivos_salida.json().get("value", []):
            if item.get("name") == nombre_consolidado_cloud:
                consolidado_existe = True
                url_consolidado_previo = item.get("@microsoft.graph.downloadUrl")
                break

    archivos_a_procesar = [archivos_ventas[-1]] if consolidado_existe else archivos_ventas

    bloques: list[pd.DataFrame] = []
    orden_columnas: list[str] | None = None

    if consolidado_existe and url_consolidado_previo:
        try:
            content_prev = requests.get(url_consolidado_previo).content
            df_existente = pd.read_csv(
                io.BytesIO(content_prev), sep=delimitador, decimal=",",
                encoding="utf-8", dtype={"MES": str, "AÑO": str}, low_memory=False,
            )
            orden_columnas = list(df_existente.columns)
            df_filtrado = df_existente[
                (df_existente["MES"].astype(str).str.strip().str.capitalize() != MESES_ESPANOL[spec.mes])
                | (df_existente["AÑO"].astype(str).str.strip() != str(spec.anio))
            ]
            bloques.append(df_filtrado)
            print(f"  Modo upsert cloud (Ventas) — reprocesando periodo activo: {spec.etiqueta}")
        except Exception as e:
            print(f"  ⚠️ No pude leer el consolidado previo de ventas ({e}); reconstruyo desde cero")
            archivos_a_procesar = archivos_ventas
            bloques = []

    for archivo in archivos_a_procesar:
        nombre = archivo["name"]
        url_download = archivo["@microsoft.graph.downloadUrl"]
        mes, anio = _detectar_mes_anio_cloud(url_download, nombre, mapping)
        
        if not mes or not anio:
            if "rutero" in nombre.lower():
                mes, anio = MESES_ESPANOL[spec.mes], str(spec.anio)
            else:
                print(f"  ⚠️ Sin mes/año detectable en ventas, omitido: {nombre}")
                continue

        print(f"  → Procesando ventas desde nube: {nombre}")
        try:
            content = requests.get(url_download).content
            try:
                df = pd.read_excel(io.BytesIO(content))
            except Exception:
                df = pd.read_html(io.BytesIO(content), encoding="latin-1")[0]

            df.rename(columns=mapping, inplace=True)
            df = df.loc[:, ~df.columns.duplicated()]
            df["MES"], df["AÑO"] = mes, anio

            if orden_columnas is None:
                orden_columnas = list(df.columns)
                if "MES" not in orden_columnas:
                    orden_columnas.extend(["MES", "AÑO"])

            df = df[[c for c in orden_columnas if c in df.columns]]

            for col in COLUMNS_TO_CLEAN_DOT_ZERO:
                if col in df.columns:
                    df[col] = (
                        pd.to_numeric(df[col], errors="coerce")
                          .astype(str)
                          .replace(r"\.0$", "", regex=True)
                          .replace("nan", "")
                    )
            for col in VTA_COLUMNS:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

            bloques.append(df)
        except Exception as e:
            print(f"  ❌ Error procesando ventas {nombre}: {e}")

    if not bloques:
        print("❌ No se generó ningún bloque de ventas consolidable.")
        return 0

    df_final = pd.concat(bloques, ignore_index=True)
    df_final["AÑO"] = (
        pd.to_numeric(df_final["AÑO"], errors="coerce")
          .fillna("")
          .astype(pd.StringDtype())
    )

    csv_buffer = io.StringIO()
    df_final.to_csv(csv_buffer, sep=delimitador, index=False, encoding="utf-8", decimal=",")
    csv_bytes = csv_buffer.getvalue().encode("utf-8")

    url_subida_csv = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{urllib.parse.quote(ruta_salida_relativa)}/{nombre_consolidado_cloud}:/content"
    headers_subida = {**headers, "Content-Type": "text/csv; charset=utf-8"}
    res_subida = requests.put(url_subida_csv, headers=headers_subida, data=csv_bytes)
    
    if res_subida.status_code in [200, 201]:
        print(f"  ✅ Consolidado de Ventas guardado en SharePoint: {ruta_salida_relativa}/{nombre_consolidado_cloud} ({len(df_final):,} filas)")
    else:
        print(f"  ❌ Error al subir CSV consolidado de ventas: {res_subida.text}")

    try:
        print(f"  🚀 Preparando cargue de ventas consolidadas a Supabase...")
        df_supabase = df_final.replace([np.inf, -np.inf], 0)
        df_supabase_limpio = df_supabase.astype(object).where(pd.notnull(df_supabase), None)
        registros_json = df_supabase_limpio.to_dict(orient="records")
        
        chunk_size = 1000
        for i in range(0, len(registros_json), chunk_size):
            chunk = registros_json[i:i + chunk_size]
            supabase.table("dyp_ventas").upsert(chunk).execute()
        print(f"  ✅ ÉXITO CLOUD (Ventas): {len(registros_json)} filas subidas a Supabase.")
    except Exception as ex_cloud:
        print(f"  ⚠ ADVERTENCIA EN CARGA SUPABASE (Ventas): {ex_cloud}")

    return len(df_final)


# =============================================================================
# 2. CARGA + CONSOLIDACIÓN EN LA NUBE (IMPACTOS)
# =============================================================================

def consolidar_impactos(
    archivos_cloud: list[dict],
    spec: pr.PeriodoSpec,
    headers,
    site_id,
    delimitador: str = DELIMITADOR_CSV,
) -> int:
    archivos_impactos = []
    for f in archivos_cloud:
        nombre_lower = f["name"].lower()
        # Captura archivos de impactos pero excluye explícitamente los de segmentos
        if ("impacto" in nombre_lower or "cuota" in nombre_lower) and "segmento" not in nombre_lower:
            archivos_impactos.append(f)

    if not archivos_impactos:
        print("❌ Sin archivos de Impactos para consolidar en la nube.")
        return 0

    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    nombre_consolidado_cloud = f"Consolidado_Impactos_{periodo_nom}.csv"
    ruta_salida_relativa = "Equipo Información/BI/INVOLVES/SALIDAS/DYP"

    bloques: list[pd.DataFrame] = []
    for archivo in archivos_impactos:
        nombre = archivo["name"]
        url_download = archivo["@microsoft.graph.downloadUrl"]
        
        print(f"  → Procesando impactos desde nube: {nombre}")
        try:
            content = requests.get(url_download).content
            try:
                df = pd.read_excel(io.BytesIO(content))
            except Exception:
                df = pd.read_html(io.BytesIO(content), encoding="latin-1")[0]

            df = df.loc[:, ~df.columns.duplicated()]
            bloques.append(df)
        except Exception as e:
            print(f"  ❌ Error procesando impactos {nombre}: {e}")

    if not bloques:
        print("❌ No se generó ningún bloque de impactos consolidable.")
        return 0

    df_final = pd.concat(bloques, ignore_index=True)

    csv_buffer = io.StringIO()
    df_final.to_csv(csv_buffer, sep=delimitador, index=False, encoding="utf-8", decimal=",")
    csv_bytes = csv_buffer.getvalue().encode("utf-8")

    url_subida_csv = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{urllib.parse.quote(ruta_salida_relativa)}/{nombre_consolidado_cloud}:/content"
    headers_subida = {**headers, "Content-Type": "text/csv; charset=utf-8"}
    res_subida = requests.put(url_subida_csv, headers=headers_subida, data=csv_bytes)
    
    if res_subida.status_code in [200, 201]:
        print(f"  ✅ Consolidado de Impactos guardado en SharePoint: {ruta_salida_relativa}/{nombre_consolidado_cloud} ({len(df_final):,} filas)")
    else:
        print(f"  ❌ Error al subir CSV consolidado de impactos: {res_subida.text}")

    return len(df_final)


# =============================================================================
# 3. CARGA + CONSOLIDACIÓN EN LA NUBE (IMPACTO_SEGMENTOS)
# =============================================================================

def consolidar_impacto_segmentos(
    archivos_cloud: list[dict],
    spec: pr.PeriodoSpec,
    headers,
    site_id,
    delimitador: str = DELIMITADOR_CSV,
) -> int:
    archivos_segmentos = []
    for f in archivos_cloud:
        nombre_lower = f["name"].lower()
        # Coincidencia exacta con el nombre del archivo generado por la ETL
        if "impacto_segmentos" in nombre_lower or "segmento" in nombre_lower:
            archivos_segmentos.append(f)

    if not archivos_segmentos:
        print("❌ Sin archivos de impacto_segmentos para consolidar en la nube.")
        return 0

    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    nombre_consolidado_cloud = f"Consolidado_Impacto_Segmentos_{periodo_nom}.csv"
    ruta_salida_relativa = "Equipo Información/BI/INVOLVES/SALIDAS/DYP"

    bloques: list[pd.DataFrame] = []
    for archivo in archivos_segmentos:
        nombre = archivo["name"]
        url_download = archivo["@microsoft.graph.downloadUrl"]
        
        print(f"  → Procesando impacto_segmentos desde nube: {nombre}")
        try:
            content = requests.get(url_download).content
            try:
                df = pd.read_excel(io.BytesIO(content))
            except Exception:
                df = pd.read_html(io.BytesIO(content), encoding="latin-1")[0]

            df = df.loc[:, ~df.columns.duplicated()]
            bloques.append(df)
        except Exception as e:
            print(f"  ❌ Error procesando impacto_segmentos {nombre}: {e}")

    if not bloques:
        print("❌ No se generó ningún bloque de impacto_segmentos consolidable.")
        return 0

    df_final = pd.concat(bloques, ignore_index=True)

    csv_buffer = io.StringIO()
    df_final.to_csv(csv_buffer, sep=delimitador, index=False, encoding="utf-8", decimal=",")
    csv_bytes = csv_buffer.getvalue().encode("utf-8")

    url_subida_csv = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{urllib.parse.quote(ruta_salida_relativa)}/{nombre_consolidado_cloud}:/content"
    headers_subida = {**headers, "Content-Type": "text/csv; charset=utf-8"}
    res_subida = requests.put(url_subida_csv, headers=headers_subida, data=csv_bytes)
    
    if res_subida.status_code in [200, 201]:
        print(f"  ✅ Consolidado de impacto_segmentos guardado en SharePoint: {ruta_salida_relativa}/{nombre_consolidado_cloud} ({len(df_final):,} filas)")
    else:
        print(f"  ❌ Error al subir CSV consolidado de impacto_segmentos: {res_subida.text}")

    return len(df_final)


# ─────────────────────────────────────────────────────────────────────────────
# PUNTO DE ENTRADA PRINCIPAL (MULTI-PERIODO)
# ─────────────────────────────────────────────────────────────────────────────

def run() -> int:
    import argparse
    parser = argparse.ArgumentParser(description="ETL VENTAS, IMPACTOS Y SEGMENTOS DYP — Nube")
    pr.cli_add_periodos_arg(parser)
    args, _ = parser.parse_known_args()
    specs = pr.periodos_de_args(args)

    print("🔒 Autenticando con Microsoft Graph API...")
    token = obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = obtener_site_id(headers)

    total_filas = 0
    for spec in specs:
        print(f"\n🎯 Procesando periodo: {spec.etiqueta} (Año: {spec.anio})")
        
        rutas_cloud = obtener_archivos_carpeta_sharepoint_recursivo(headers, site_id, anio_proceso=spec.anio)
        print(f"✅ Archivos encontrados en la nube para {spec.anio}: {len(rutas_cloud)}")
        
        # Ejecutar las tres consolidaciones de manera independiente
        total_filas += consolidar_ventas(rutas_cloud, spec, headers, site_id)
        consolidar_impactos(rutas_cloud, spec, headers, site_id)
        consolidar_impacto_segmentos(rutas_cloud, spec, headers, site_id)

    return total_filas


if __name__ == "__main__":
    run()