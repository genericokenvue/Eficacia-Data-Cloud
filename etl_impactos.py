"""
etl_impactos.py
───────────────
Consolida los archivos Excel mensuales de **Impactos D&P** (Cuota, Venta,
Clientes Visitar/Visitados/Impactar/Impactados/Georeferenciados, %Efectividad)
en un único CSV, **100% en la nube** (SharePoint vía Microsoft Graph API) —
mismo patrón que `etl_ventas.py`.

NOTA: `etl_ventas.py` ya incluye su propia `consolidar_impactos()` y corre
ambas consolidaciones (Ventas + Impactos + Impacto_Segmentos) en un solo
`run()`. Este script queda como el equivalente standalone/independiente por
si quieres correr SOLO Impactos (p. ej. un paso de workflow separado) sin
tocar Ventas.

Origen (SharePoint)
────────────────────
    Equipo Información/BI/INVOLVES/BASES DE RESPUESTAS/DYP/<año>/... (recursivo)
    Se toman los archivos cuyo nombre contenga "impacto" o "cuota",
    excluyendo explícitamente los que contengan "segmento".

Salida (SharePoint)
────────────────────
    Equipo Información/BI/INVOLVES/SALIDAS/DYP/Consolidado_Impactos_<MES>_<AÑO>.csv

Esquema esperado en el Excel origen
───────────────────────────────────
    Supervisor, Asesor, Ciudad, Cuota, Venta,
    Total Clientes a Visitar, Clientes Visitados,
    Total Clientes a Impactar, Clientes Impactados,
    Clientes Georeferenciados,
    % Efect (.0/.1/.2/.3),     ← cuatro columnas con el mismo nombre
    Mes, Año.

`% Efect.0` = Efectividad Venta, `% Efect.1` = Visita,
`% Efect.2` = Impacto, `% Efect.3` = Georef.
"""

from __future__ import annotations

import io
import os
import urllib.parse
import warnings
import locale

import msal
import requests
import numpy as np
import pandas as pd

import paths
import periodo_resolver as pr

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

# ─────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN GLOBAL & AZURE API
# ─────────────────────────────────────────────────────────────────────────────
TENANT_ID     = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")

if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
    raise ValueError("❌ ERROR: Faltan credenciales de Azure. Verifica tu .env o tus GitHub Secrets.")

MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

RUTA_SALIDA_DYP = "Equipo Información/BI/INVOLVES/SALIDAS/DYP"

COLUMN_MAPPING_ORIGEN_A_DESTINO = {
    "Supervisor": "Supervisor", "Asesor": "Asesor", "Ciudad": "Ciudad",
    "Cuota": "Cuota", "Venta": "Venta",
    "Total Clientes a Visitar":  "Total Clientes Visitar",
    "Clientes Visitados":        "Clientes Visitados",
    "Total Clientes a Impactar": "Total Clientes Impactar",
    "Clientes Impactados":       "Clientes Impactados",
    "Clientes Georeferenciados": "Clientes Georeferenciados",
    "% Efect":    "Efectividad Venta (%)",
    "% Efect.1":  "Efectividad Visita (%)",
    "% Efect.2":  "Efectividad Impacto (%)",
    "% Efect.3":  "Efectividad Georef (%)",
    "% Efect_1":  "Efectividad Visita (%)",
    "% Efect_2":  "Efectividad Impacto (%)",
    "% Efect_3":  "Efectividad Georef (%)",
    "Mes": "MES", "mes": "MES",
    "Año": "AÑO", "año": "AÑO",
}

COLUMNS_TO_CLEAN_DOT_ZERO = [
    "Total Clientes Visitar", "Clientes Visitados",
    "Total Clientes Impactar", "Clientes Impactados",
    "Clientes Georeferenciados",
]
VTA_COLUMNS_CASH = ["Cuota", "Venta"]
VTA_COLUMNS_PERCENT = [
    "Efectividad Venta (%)", "Efectividad Visita (%)",
    "Efectividad Impacto (%)", "Efectividad Georef (%)",
]

DELIMITADOR_CSV = "|"


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES AUXILIARES DE MICROSOFT GRAPH API
# ─────────────────────────────────────────────────────────────────────────────

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
    """Igual que en etl_ventas.py: recorre recursivamente BASES DE RESPUESTAS/DYP/<año>,
    excluyendo cualquier ruta con 'MES ACTUAL'."""
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
                        "@microsoft.graph.downloadUrl": item.get("@microsoft.graph.downloadUrl"),
                    })

    recorrer_items(elementos, ruta_relativa)
    return sorted(archivos_validos, key=lambda x: x["path"])


# ─────────────────────────────────────────────────────────────────────────────
# CONSOLIDACIÓN EN LA NUBE (IMPACTOS)
# ─────────────────────────────────────────────────────────────────────────────

def consolidar_impactos(
    archivos_cloud: list[dict],
    spec: pr.PeriodoSpec,
    headers,
    site_id,
    mapping: dict = COLUMN_MAPPING_ORIGEN_A_DESTINO,
    delimitador: str = DELIMITADOR_CSV,
) -> int:
    archivos_impactos = []
    for f in archivos_cloud:
        nombre_lower = f["name"].lower()
        if ("impacto" in nombre_lower or "cuota" in nombre_lower) and "segmento" not in nombre_lower:
            archivos_impactos.append(f)

    if not archivos_impactos:
        print("❌ Sin archivos de Impactos para consolidar en la nube.")
        return 0

    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    nombre_consolidado_cloud = f"Consolidado_Impactos_{periodo_nom}.csv"

    # Modo upsert: si ya existe el consolidado de este periodo, solo se
    # reprocesa el último archivo (corte más reciente) y se reemplazan sus
    # filas (MES, AÑO) en el consolidado, igual que hace etl_ventas.py.
    archivos_salida = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{urllib.parse.quote(RUTA_SALIDA_DYP)}:/children",
        headers=headers,
    )
    consolidado_existe = False
    url_consolidado_previo = None
    if archivos_salida.status_code == 200:
        for item in archivos_salida.json().get("value", []):
            if item.get("name") == nombre_consolidado_cloud:
                consolidado_existe = True
                url_consolidado_previo = item.get("@microsoft.graph.downloadUrl")
                break

    archivos_a_procesar = [archivos_impactos[-1]] if consolidado_existe else archivos_impactos

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
                (df_existente["MES"].astype(str).str.strip().str.capitalize() != MESES_ESPANOL[spec.mes].capitalize())
                | (df_existente["AÑO"].astype(str).str.strip() != str(spec.anio))
            ]
            bloques.append(df_filtrado)
            print(f"  Modo upsert cloud (Impactos) — reprocesando periodo activo: {spec.etiqueta}")
        except Exception as e:
            print(f"  ⚠️ No pude leer el consolidado previo de impactos ({e}); reconstruyo desde cero")
            archivos_a_procesar = archivos_impactos
            bloques = []

    for archivo in archivos_a_procesar:
        nombre = archivo["name"]
        url_download = archivo["@microsoft.graph.downloadUrl"]

        print(f"  → Procesando impactos desde nube: {nombre}")
        try:
            content = requests.get(url_download).content
            try:
                df = pd.read_excel(io.BytesIO(content))
            except Exception:
                df = pd.read_html(io.BytesIO(content), encoding="latin-1")[0]

            df.rename(columns=mapping, inplace=True)
            df = df.loc[:, ~df.columns.duplicated()]

            # Si el Excel no trae MES/AÑO propios, se asignan del periodo activo.
            if "MES" not in df.columns or df["MES"].dropna().empty:
                df["MES"] = MESES_ESPANOL[spec.mes].capitalize()
            if "AÑO" not in df.columns or df["AÑO"].dropna().empty:
                df["AÑO"] = str(spec.anio)

            if orden_columnas is None:
                orden_columnas = list(df.columns)
                for c in ("MES", "AÑO"):
                    if c not in orden_columnas:
                        orden_columnas.append(c)
            df = df[[c for c in orden_columnas if c in df.columns]]

            for col in COLUMNS_TO_CLEAN_DOT_ZERO:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")
                    df[col] = (
                        df[col].astype(str)
                              .replace(r"\.0$", "", regex=True)
                              .replace("nan", "")
                    )

            for col in VTA_COLUMNS_CASH:
                if col in df.columns:
                    s = (df[col].astype(str).str.strip()
                                  .str.replace(",", "", regex=False)
                                  .str.replace(".", "", regex=False))
                    df[col] = pd.to_numeric(s, errors="coerce").fillna(0).astype(np.int64)

            for col in VTA_COLUMNS_PERCENT:
                if col in df.columns:
                    s = (df[col].astype(str).str.strip()
                                  .str.replace(".", "", regex=False)
                                  .str.replace(",", ".", regex=False)
                                  .str.replace(r"[^\d\.\-]", "", regex=True))
                    df[col] = pd.to_numeric(s, errors="coerce")

            bloques.append(df)
        except Exception as e:
            print(f"  ❌ Error procesando {nombre}: {e}")

    if not bloques:
        print("❌ No se generó ningún bloque de impactos consolidable.")
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

    url_subida_csv = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{urllib.parse.quote(RUTA_SALIDA_DYP)}/{nombre_consolidado_cloud}:/content"
    headers_subida = {**headers, "Content-Type": "text/csv; charset=utf-8"}
    res_subida = requests.put(url_subida_csv, headers=headers_subida, data=csv_bytes)

    if res_subida.status_code in (200, 201):
        print(f"  ✅ Consolidado de Impactos guardado en SharePoint: {RUTA_SALIDA_DYP}/{nombre_consolidado_cloud} ({len(df_final):,} filas)")
    else:
        print(f"  ❌ Error al subir CSV consolidado de impactos: {res_subida.text}")

    return len(df_final)


# ─────────────────────────────────────────────────────────────────────────────
# PUNTO DE ENTRADA (MULTI-PERIODO)
# ─────────────────────────────────────────────────────────────────────────────

def run() -> int:
    import argparse
    parser = argparse.ArgumentParser(description="ETL IMPACTOS DYP — Nube")
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
        total_filas += consolidar_impactos(rutas_cloud, spec, headers, site_id)

    return total_filas


if __name__ == "__main__":
    run()