import os
import sys
import io
import re
import tempfile
import requests
import urllib.parse
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv

# Cargar variables del .env
load_dotenv()

# Cargar config.env antes de leer UMBRAL_OK/WARNING desde os.environ
from config_loader import cargar_config
cargar_config()

from alertas_logger import get_logger
_log = get_logger("calcular_cumplimientos")

# alertas_logger ya añade SCRIPTS/ a sys.path
import paths

# KPIs comerciales D&P (venta, impactos, MSL, productos nuevos)
import cumplimiento_dyp

# Tabla maestra de personas (ACRONIMO como llave canónica)
import base_cupos as bcm


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE CONEXIÓN NUBE (Microsoft Graph API usando Drive ID)
# ─────────────────────────────────────────────────────────────────────────────
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

def _obtener_token_azure():
    """Genera el token de acceso para la API de Microsoft Graph."""
    credential = ClientSecretCredential(
        tenant_id=AZURE_TENANT_ID,
        client_id=AZURE_CLIENT_ID,
        client_secret=AZURE_CLIENT_SECRET
    )
    return credential.get_token("https://graph.microsoft.com/.default").token

_DRIVE_ID_CACHE: dict = {}


def _obtener_default_drive_id(token: str) -> str:
    """
    Obtiene el ID del drive conectándose al sitio corporativo específico de
    SharePoint. Usa `paths.SHAREPOINT_SITE_NAME` (ej. "JJ451") para buscar el
    sitio correcto — antes esta función buscaba un texto fijo ("eficacia")
    que no necesariamente coincide con el nombre real del sitio y podía
    resolver el drive equivocado (causa típica de 404 en TODAS las rutas
    candidatas, incluso siendo la ruta relativa correcta).
    """
    if "drive_id" in _DRIVE_ID_CACHE:
        return _DRIVE_ID_CACHE["drive_id"]

    headers = {"Authorization": f"Bearer {token}"}
    nombre_sitio = getattr(paths, "SHAREPOINT_SITE_NAME", "eficacia")

    # Búsqueda del sitio corporativo por el nombre configurado en paths.py
    res_site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites?search={urllib.parse.quote(nombre_sitio)}",
        headers=headers,
    )
    if res_site.status_code == 200:
        sites = res_site.json().get("value", [])
        if sites:
            # Si hay varios resultados, preferir un match exacto de nombre.
            site = next(
                (s for s in sites
                 if nombre_sitio.lower() in (s.get("name", "").lower(), s.get("displayName", "").lower())),
                sites[0],
            )
            site_id = site.get("id")
            print(f"    ℹ️  Sitio SharePoint resuelto: {site.get('displayName') or site.get('name')} "
                  f"({site.get('webUrl')})")
            res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
            if res_drive.status_code == 200:
                drive = res_drive.json()
                print(f"    ℹ️  Drive resuelto: {drive.get('name')} ({drive.get('webUrl')})")
                _DRIVE_ID_CACHE["drive_id"] = drive.get("id")
                return drive.get("id")

    # Plan B: Si no lo encuentra por búsqueda, intenta con el raíz del tenant
    res_root = requests.get("https://graph.microsoft.com/v1.0/sites/root", headers=headers)
    if res_root.status_code == 200:
        site_id = res_root.json().get("id")
        res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
        if res_drive.status_code == 200:
            print("    ⚠️  No se encontró el sitio por nombre; usando el sitio raíz del tenant como fallback.")
            _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
            return res_drive.json().get("id")

    raise Exception(f"No se pudo obtener el Drive ID del sitio SharePoint '{nombre_sitio}': {res_site.text}")

def _limpiar_ruta_graph(ruta_sharepoint: str) -> str:
    """
    Normaliza separadores de la ruta relativa dentro del drive.
    IMPORTANTE: NO remueve 'Equipo Información/' — esa es una carpeta REAL
    dentro de la biblioteca 'Documentos compartidos' del sitio (así lo
    define paths.SHAREPOINT_BASE_DIR), no un alias del nombre del sitio.
    Solo se remueven prefijos que sí son alias genéricos del nombre de la
    biblioteca documental por defecto.
    """
    ruta_sharepoint = ruta_sharepoint.replace("\\", "/")
    for prefijo in [
        "Documentos compartidos/",
        "Shared Documents/",
    ]:
        if ruta_sharepoint.startswith(prefijo):
            ruta_sharepoint = ruta_sharepoint.replace(prefijo, "", 1)
    return ruta_sharepoint.lstrip("/")

def _leer_excel_cloud(ruta_sharepoint: str, descripcion: str) -> pd.DataFrame:
    """Lee un archivo Excel directamente desde SharePoint usando el Drive ID en Graph API con reintentos de ruta."""
    print(f"  ⏳ Leyendo {descripcion} desde SharePoint ({ruta_sharepoint})...")
    token = _obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    
    rutas_candidatas = [
        ruta_limpia,
        f"Documentos compartidos/{ruta_limpia}",
    ]
    
    response = None
    url_usada = ""
    for r in rutas_candidatas:
        ruta_codificada = urllib.parse.quote(r)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{ruta_codificada}:/content"
        print(f"    🔍 Probando URL Graph API: {url}")
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            url_usada = url
            break

    if response is not None and response.status_code == 200:
        print(f"    ✓ Archivo descargado exitosamente.")
        return pd.read_excel(io.BytesIO(response.content))
    else:
        status = response.status_code if response is not None else "N/A"
        text = response.text if response is not None else "Sin respuesta"
        raise Exception(f"Error al descargar {descripcion} desde Graph API (404/Error). Última URL probada: {url_usada} | Status: {status} - {text}")

def _guardar_excel_cloud(df: pd.DataFrame, ruta_sharepoint: str, descripcion: str):
    """Sube un DataFrame convertido a Excel directamente a SharePoint vía Graph API usando el Drive ID."""
    print(f"  ⏳ Guardando {descripcion} en SharePoint ({ruta_sharepoint})...")
    token = _obtener_token_azure()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_bytes = output.getvalue()
    
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    ruta_codificada = urllib.parse.quote(ruta_limpia)
    
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{ruta_codificada}:/content"
    response = requests.put(url, headers=headers, data=excel_bytes)
    if response.status_code in [200, 201]:
        print(f"  ✓ {descripcion} guardado exitosamente en SharePoint")
    else:
        raise Exception(f"Error al subir {descripcion} a Graph API: {response.status_code} - {response.text}")


def _subir_bytes_cloud(data: bytes, ruta_sharepoint: str, descripcion: str,
                       content_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    """
    Sube bytes ya generados (p.ej. un Workbook de openpyxl volcado a BytesIO)
    directamente a SharePoint vía Graph API. Equivalente a
    `_guardar_excel_cloud` pero sin pasar por un DataFrame, para archivos con
    formato/multiples hojas.
    """
    token = _obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": content_type}
    drive_id = _obtener_default_drive_id(token)
    ruta_codificada = urllib.parse.quote(_limpiar_ruta_graph(ruta_sharepoint))
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{ruta_codificada}:/content"
    response = requests.put(url, headers=headers, data=data)
    if response.status_code in (200, 201):
        return True
    raise Exception(f"Error al subir {descripcion} a Graph API: {response.status_code} - {response.text}")


def _listar_hijos_cloud(ruta_carpeta: str) -> list:
    """
    Lista los archivos/subcarpetas de `ruta_carpeta` en SharePoint vía Graph
    API (equivalente en la nube de recorrer un directorio local). Cada
    elemento devuelto es el dict nativo de Graph (trae 'name',
    'lastModifiedDateTime' y 'file' o 'folder' según el tipo de ítem).
    """
    token = _obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_carpeta)
    ruta_codificada = urllib.parse.quote(ruta_limpia)

    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{ruta_codificada}:/children"
    items = []
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(
                f"Error al listar carpeta '{ruta_carpeta}' en Graph API: "
                f"{response.status_code} - {response.text}"
            )
        data = response.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _resolver_exhib_data_dir_cloud() -> str:
    """
    Equivalente en la nube de `paths._resolver_exhib_data_dir()`: la carpeta
    de BASES de Exhibiciones suele contener subcarpetas tipo "01. ...",
    "02. ..." (una por corte/actualización); se toma la más reciente. Si no
    hay subcarpetas que matcheen el patrón, se usa la carpeta raíz tal cual.
    """
    raiz = paths.RUTA_CARPETA_BASES_EXHIB
    try:
        hijos = _listar_hijos_cloud(raiz)
    except Exception:
        return raiz
    candidatos = [
        h for h in hijos
        if h.get("folder") and re.match(r"^\d{2}\.\s", h.get("name", ""))
    ]
    if not candidatos:
        return raiz
    candidatos.sort(key=lambda h: h.get("lastModifiedDateTime", ""), reverse=True)
    return f"{raiz}/{candidatos[0]['name']}"


def _buscar_archivo_cloud(ruta_carpeta: str, patron_regex: str, contexto: str) -> str:
    """
    Localiza UN único archivo cuyo nombre matchee `patron_regex` dentro de
    `ruta_carpeta` en SharePoint. Equivalente en la nube de
    `periodo_resolver._find_unico`. Ignora temporales de Excel (~$*).
    Devuelve la ruta SharePoint completa del archivo encontrado.
    """
    hijos = _listar_hijos_cloud(ruta_carpeta)
    matches = [
        h["name"] for h in hijos
        if h.get("file") and re.match(patron_regex, h.get("name", ""))
        and not h["name"].startswith("~$")
    ]
    if not matches:
        raise FileNotFoundError(
            f"[{contexto}] No se encontró archivo con patrón '{patron_regex}' en {ruta_carpeta}"
        )
    if len(matches) > 1:
        raise RuntimeError(
            f"[{contexto}] Múltiples coincidencias en {ruta_carpeta}: {matches}. Esperaba exactamente uno."
        )
    return f"{ruta_carpeta}/{matches[0]}"


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE RUTAS USANDO EXCLUSIVAMENTE `paths.py`
# ─────────────────────────────────────────────────────────────────────────────
RUTA_CIF = f"{paths.RUTA_CARPETA_SALIDAS_CIF}/Plan de trabajo.xlsx"
RUTA_SOS = f"{paths.RUTA_CARPETA_SALIDAS_SOS}/Cumplimiento_Captura_SOS.xlsx"

DIR_NP_OUT      = paths.RUTA_CARPETA_SALIDAS_NP
DIR_PRECIOS_OUT = paths.RUTA_CARPETA_SALIDAS_PRECIOS
DIR_ALERTAS     = getattr(paths, 'ALERTAS_DIR', f"{paths._SALIDAS_ROOT}/ALERTAS")
RUTA_MAESTRA    = getattr(paths, 'ALERTAS_MAESTRO', f"{DIR_ALERTAS}/maestra_supervisores.xlsx")

# paths.py sólo define el directorio local de ALERTAS (ALERTAS_DIR, usado para
# cachear adjuntos antes de enviarlos por correo). No existe un
# `RUTA_CARPETA_SALIDAS_ALERTAS` en la nube, así que lo construimos aquí
# siguiendo la misma convención que el resto de módulos (_SALIDAS_ROOT/<módulo>).
RUTA_CARPETA_SALIDAS_ALERTAS = f"{paths._SALIDAS_ROOT}/ALERTAS"
RUTA_MAESTRA_CLOUD           = f"{paths._BASES_ROOT}/ALERTAS/MAESTRO_SUPERVISORES.xlsx"
# Carpeta en la nube donde deben quedar los adjuntos por supervisor
# (…/BI/INVOLVES/SALIDAS/ALERTAS/ADJUNTOS)
RUTA_CARPETA_ADJUNTOS_CLOUD  = f"{RUTA_CARPETA_SALIDAS_ALERTAS}/ADJUNTOS"


# ─────────────────────────────────────────────────────────────────────────────
# F) PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────────────────────────

def main() -> dict:
    """
    Ejecuta el cálculo completo conectado a la nube usando paths.py.
    Devuelve dict para alertas_telegram.py y alertas_email.py:
      {df_detalle, df_resumen, rutas_adjuntos, mes, anio}
    """
    print("\n" + "═" * 55)
    print("  CALCULAR CUMPLIMIENTOS — Eficacia (Cloud via paths.py)")
    print("═" * 55)
    _log.info("Iniciando cálculo de cumplimientos en la nube con paths.py")

    # ── Cargar archivos de detalle desde la Nube usando paths ─────────────
    print("\nCargando archivos de detalle:")
    df_cif = _leer_excel_cloud(RUTA_CIF, "CIF (Plan de trabajo.xlsx)")
    df_sos = _leer_excel_cloud(RUTA_SOS, "SOS")
    
    ruta_np_cloud = f"{DIR_NP_OUT}/REPORTE_NO_PRESENCIA.xlsx"
    ruta_pr_cloud = f"{DIR_PRECIOS_OUT}/REPORTE_CAPTURA_PRECIOS.xlsx"
    
    df_np = _leer_excel_cloud(ruta_np_cloud, "No Presencia")
    df_pr = _leer_excel_cloud(ruta_pr_cloud, "Precios")
    
    df_cif_pdv = df_cif

    # ── Cargar Base cupos (tabla maestra de personas, llave ACRONIMO) ────
    print("\nCargando tabla maestra de personas (Base cupos):")
    df_bc   = bcm.cargar()
    universo = bcm.universo_personas(df_bc)
    sups_bc  = bcm.supervisores(df_bc)
    nombres_sups_bc = set(sups_bc["NOMBRE"].astype(str).str.upper())
    
    if "ES_GDD" in df_bc.columns:
        nombres_sups_bc.update(
            df_bc[df_bc["ES_GDD"] == True]["NOMBRE"].astype(str).str.upper()
        )
    if "ES_LIDER" in df_bc.columns:
        nombres_sups_bc.update(
            df_bc[df_bc["ES_LIDER"] == True]["NOMBRE"].astype(str).str.upper()
        )
    idx_bc = bcm.construir_indices(df_bc)
    print(f"  ✓ Universo: {len(universo)} personas activas | {len(sups_bc)} supervisores")

    # ── Maestra de supervisores ──────────────────────────────────────────
    print("\nVerificando maestra de supervisores:")
    generar_maestra(df_np, df_pr, df_sos)

    # ── KPIs V3 ──────────────────────────────────────────────────────────
    print("\nLeyendo KPIs V3 pre-calculados por los ETLs:")
    kpis_v3 = cargar_kpis_v3(idx_bc, nombres_sups_bc)
    df_cif_gest = kpis_v3["cif_gest"]
    df_np_gest  = kpis_v3["np_gest"]
    df_pr_gest  = kpis_v3["pr_gest"]
    df_sos_gest = kpis_v3["sos_gest"]
    df_exp_gest = kpis_v3.get("exp_gest", pd.DataFrame())
    df_egr_gest = kpis_v3.get("egr_gest", pd.DataFrame())
    
    df_cif_sup = pd.DataFrame()
    df_np_sup  = pd.DataFrame()
    df_pr_sup  = pd.DataFrame()
    df_sos_sup = pd.DataFrame()
    print(f"  ✓ CIF        — {len(df_cif_gest)} personas")
    print(f"  ✓ NP         — {len(df_np_gest)} personas")
    print(f"  ✓ Precios    — {len(df_pr_gest)} personas")
    print(f"  ✓ SOS        — {len(df_sos_gest)} personas")
    print(f"  ✓ Exh PAG    — {len(df_exp_gest)} empleados")
    print(f"  ✓ Exh GRATIS — {len(df_egr_gest)} empleados")

    # ── Periodo activo ────────────────────────────────────────────────────
    ahora = datetime.now()
    if not df_cif_gest.empty and "MES" in df_cif_gest.columns:
        mes  = int(pd.to_numeric(df_cif_gest["MES"], errors="coerce").dropna().mode().iloc[0])
        anio = int(pd.to_numeric(df_cif_gest["AÑO"], errors="coerce").dropna().mode().iloc[0])
    else:
        mes, anio = ahora.month, ahora.year
    print(f"  Periodo activo: {mes:02d}/{anio}")

    # ── KPIs D&P ──────────────────────────────────────────────────────────
    kpis_dyp = cumplimiento_dyp.calcular_kpis_dyp(
        mes, anio,
        base_cupos_idx=idx_bc,
        nombres_supervisores_bc=nombres_sups_bc,
    )
    df_dyp_gest = kpis_dyp["gestor"]
    df_dyp_sup  = kpis_dyp["supervisor"]

    # ── Ensamblaje final ──────────────────────────────────────────────────
    print("\nEnsamblando detalle consolidado:")
    df_detalle = ensamblar_detalle(
        universo,
        df_cif_gest, df_cif_sup,
        df_np_gest,  df_np_sup,
        df_pr_gest,  df_pr_sup,
        df_sos_gest, df_sos_sup,
        df_dyp_gest=df_dyp_gest,
        df_dyp_sup=df_dyp_sup,
        df_exp_gest=df_exp_gest,
        df_egr_gest=df_egr_gest,
        mes=mes, anio=anio,
    )

    n_gest = (df_detalle["ES_SUPERVISOR"] == False).sum()
    n_sup  = (df_detalle["ES_SUPERVISOR"] == True).sum()
    print(f"  Personas en detalle: {len(df_detalle)} (gestores: {n_gest} | supervisores: {n_sup})")

    # ── Resumen ───────────────────────────────────────────────────────────
    print("\nCalculando resumen por supervisor:")
    df_resumen = calcular_resumen(df_detalle, universo)
    print(f"  Supervisores en resumen: {len(df_resumen)}")

    # ── Guardar archivos usando las rutas de `paths.py` ────────────────────
    print("\nGuardando archivos de salida en SharePoint:")
    ruta_salida_detalle = f"{paths._SALIDAS_ROOT}/Detalle_Cumplimientos_{anio}_{mes:02d}.xlsx"
    ruta_salida_resumen = f"{paths._SALIDAS_ROOT}/Resumen_Supervisores_{anio}_{mes:02d}.xlsx"
    
    _guardar_excel_cloud(df_detalle, ruta_salida_detalle, "Detalle Consolidado")
    _guardar_excel_cloud(df_resumen, ruta_salida_resumen, "Resumen por Supervisor")

    def _agregar_acr_sup(df, col_sup_origen="SUPERVISOR_LIDER"):
        if df is None or df.empty or col_sup_origen not in df.columns:
            return df
        df = df.copy()
        df["ACRONIMO_SUP"] = df[col_sup_origen].apply(
            lambda s: _resolver_acr_supervisor(s, nombres_sups_bc, idx_bc)
        )
        return df

    df_cif_pdv_e = _agregar_acr_sup(df_cif_pdv)
    df_np_e      = _agregar_acr_sup(df_np)
    df_pr_e      = _agregar_acr_sup(df_pr)
    df_sos_e     = _agregar_acr_sup(df_sos)

    print("\nPre-cargando insumos D&P para hojas de productos nuevos:")
    seg_data = _precargar_segmentos_nuevos(mes, anio, idx_bc, nombres_sups_bc)
    if seg_data.get("rutero") is not None:
        n_sups_con_pdvs = len(seg_data.get("pdvs_por_sup") or {})
        print(f"  ✓ Segmentos listos — supervisores con PDVs en periodo: {n_sups_con_pdvs}")

    rutas_adj = generar_adjuntos_por_supervisor(
        df_detalle, df_cif_pdv_e, df_np_e, df_pr_e, df_sos_e, mes, anio,
        seg_data=seg_data,
    )

    _log.info(
        f"Cálculo completado (Cloud vía paths) — periodo {mes:02d}/{anio} | "
        f"gestores={n_gest} supervisores={n_sup} adjuntos={len(rutas_adj)}"
    )
    print("\n" + "═" * 55)
    print(f"  ✅ Proceso completado en la Nube (SharePoint) — {mes:02d}/{anio}")
    print("═" * 55 + "\n")

    rango_periodo = detectar_rango_periodo(mes, anio)
    if rango_periodo.get("rango_legible"):
        print(f"\n  📆 Periodo de corte detectado: {rango_periodo['rango_legible']}"
              f" | avance esperado del mes: {rango_periodo['avance_esperado_pct']*100:.0f}%")

    return {
        "df_detalle"     : df_detalle,
        "df_resumen"     : df_resumen,
        "rutas_adjuntos" : rutas_adj,
        "mes"            : mes,
        "anio"           : anio,
        "rango_periodo"  : rango_periodo,
        "df_cif_pdv"     : df_cif_pdv_e,
        "df_np"          : df_np_e,
        "df_pr"          : df_pr_e,
        "df_sos"         : df_sos_e,
    }

##############################################################################################################################
##############################################################################################################################
##############################################################################################################################
##############################################################################################################################
##############################################################################################################################






def _ultimo_archivo(directorio: Path, patron: str) -> Path:
    """Devuelve el archivo más reciente que matchea `patron`."""
    candidatos = sorted(
        directorio.glob(patron),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidatos[0] if candidatos else directorio / patron


# Umbrales y ponderación de negocio
UMBRAL_OK      = float(os.environ.get("UMBRAL_OK",      "0.90"))
UMBRAL_WARNING = float(os.environ.get("UMBRAL_WARNING", "0.70"))

# ─────────────────────────────────────────────────────────────────────────────
# UMBRALES POR KPI (Sprint 13.4)  — defaults D17
# ─────────────────────────────────────────────────────────────────────────────
def _umbral(kpi: str, default: float) -> float:
    return float(os.environ.get(f"UMBRAL_OK_{kpi}", str(default)))

UMBRALES_OK = {
    "CIF":                _umbral("CIF",                0.90),
    "NP":                 _umbral("NP",                 1.00),
    "PRECIOS":            _umbral("PRECIOS",            1.00),
    "SOS":                _umbral("SOS",                1.00),
    "EXHIB_PAG":          _umbral("EXHIB_PAG",          1.00),
    "EXHIB_PAG_CAPTURA":  _umbral("EXHIB_PAG_CAPTURA",  1.00),  # Sprint 17
    "EXHIB_GRA_ALTO":     _umbral("EXHIB_GRA_ALTO",     1.00),  # Sprint 17
    "EXHIB_GRA_MEDIO":    _umbral("EXHIB_GRA_MEDIO",    1.00),  # Sprint 17
    "EXHIB_GRATIS_PROM":  _umbral("EXHIB_GRATIS_PROM",  1.00),  # Sprint 17.13 — promedio alto+medio
    "VENTA":              _umbral("VENTA",              0.90),
    "IMPACTOS":           _umbral("IMPACTOS",           0.90),
    "MSL":                _umbral("MSL",                1.00),
    "PROD_NUEVOS":        _umbral("PROD_NUEVOS",        0.70),
    "GLOBAL":             _umbral("GLOBAL",             0.90),
}

# Mapeo de columna del detalle → KPI key en UMBRALES_OK
COL_A_KPI = {
    "CIF_%":              "CIF",
    "NP_%":               "NP",
    "PRECIOS_%":          "PRECIOS",
    "SOS_%":              "SOS",
    "EXHIB_PAG_%":        "EXHIB_PAG",
    "EXHIB_PAG_CAPTURA_%":"EXHIB_PAG_CAPTURA",    # Sprint 17
    "EXHIB_GRA_ALTO_%":   "EXHIB_GRA_ALTO",      # Sprint 17
    "EXHIB_GRA_MEDIO_%":  "EXHIB_GRA_MEDIO",     # Sprint 17
    "EXHIB_GRATIS_PROM_%":"EXHIB_GRATIS_PROM",    # Sprint 17.13
    "VENTA_%":            "VENTA",
    "IMPACTOS_%":         "IMPACTOS",
    "MSL_%":              "MSL",
    "PROD_NUEVOS_%":      "PROD_NUEVOS",
    "CUMPL_GLOBAL_%":     "GLOBAL",
}


def umbral_de(col_o_kpi: str) -> float:
    """Devuelve UMBRAL_OK para una columna del detalle o una KPI key."""
    kpi = COL_A_KPI.get(col_o_kpi, col_o_kpi)
    return UMBRALES_OK.get(kpi, UMBRAL_OK)


# Roles usados como filtros en adjuntos / detalle.
ROL_GESTOR     = "GESTOR"
ROL_SUPERVISOR = "SUPERVISOR"


# ─────────────────────────────────────────────────────────────────────────────
# RANGO DE FECHAS DEL PERIODO (Sprint 13.1 - Adaptado a Nube)
# ─────────────────────────────────────────────────────────────────────────────

def detectar_rango_periodo(mes: int, anio: int) -> dict:
    """
    Detecta el rango de fechas activas del periodo desde
    el reporte de visitas procesado en SharePoint (FECHA_VISITA).

    Devuelve dict con:
      fecha_inicio_dt, fecha_fin_dt    — datetime.date | None
      fecha_inicio, fecha_fin          — strings "DD de MES"
      rango_legible                    — "1 al 6 de abril" o "1 de marzo al 6 de abril"
      dias_transcurridos, dias_mes_total, avance_esperado_pct
    """
    from datetime import date
    import calendar

    rango = {
        "fecha_inicio_dt": None, "fecha_fin_dt": None,
        "fecha_inicio": "", "fecha_fin": "",
        "rango_legible": "",
        "dias_transcurridos": 0, "dias_mes_total": 0, "avance_esperado_pct": 0.0,
    }

    try:
        ruta_visitas_cloud = f"{paths.RUTA_CARPETA_SALIDAS_CIF}/informe_visitas_procesado.csv"
        token = _obtener_token_azure()
        headers = {"Authorization": f"Bearer {token}"}
        drive_id = _obtener_default_drive_id(token)
        ruta_limpia = _limpiar_ruta_graph(ruta_visitas_cloud)
        
        url_codificada = urllib.parse.quote(ruta_limpia)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url_codificada}:/content"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            return rango
            
        df = pd.read_csv(
            io.BytesIO(response.content),
            sep=";", encoding="utf-8-sig", usecols=["FECHA_VISITA"],
        )
        df["F"] = pd.to_datetime(df["FECHA_VISITA"], dayfirst=True, errors="coerce")
        # Filtrar al mes/año del periodo activo
        df = df[(df["F"].dt.month == mes) & (df["F"].dt.year == anio)]
        if df.empty:
            return rango
        f_ini = df["F"].min().date()
        f_fin = df["F"].max().date()
        rango["fecha_inicio_dt"] = f_ini
        rango["fecha_fin_dt"]    = f_fin
    except Exception:
        return rango

    meses_es = {1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo",
                6:"junio", 7:"julio", 8:"agosto", 9:"septiembre",
                10:"octubre", 11:"noviembre", 12:"diciembre"}
    rango["fecha_inicio"] = f"{f_ini.day} de {meses_es.get(f_ini.month, '')}"
    rango["fecha_fin"]    = f"{f_fin.day} de {meses_es.get(f_fin.month, '')}"

    if f_ini.month == f_fin.month and f_ini.year == f_fin.year:
        rango["rango_legible"] = (
            f"{f_ini.day} al {f_fin.day} de {meses_es.get(f_fin.month, '')}"
        )
    else:
        rango["rango_legible"] = f"{rango['fecha_inicio']} al {rango['fecha_fin']}"

    dias_mes = calendar.monthrange(anio, mes)[1]
    dias_trans = f_fin.day  
    rango["dias_mes_total"] = dias_mes
    rango["dias_transcurridos"] = dias_trans
    rango["avance_esperado_pct"] = (dias_trans / dias_mes) if dias_mes > 0 else 0.0
    return rango


# ─────────────────────────────────────────────────────────────────────────────
# COLORES PARA EL EXCEL DE SALIDA
# ─────────────────────────────────────────────────────────────────────────────
COLOR_HEADER     = "1F3864"
COLOR_HEADER_FNT = "FFFFFF"
COLOR_OK         = "C6EFCE"
COLOR_WARN       = "FFEB9C"
COLOR_ERROR      = "FFC7CE"
COLOR_FILA_PAR   = "F2F2F2"
COLOR_MAESTRA_HDR= "2E4057"

FMT_PCT   = "0.00%"            
FMT_VENTA = '"$"#,##0'
FMT_NUM   = "0.####"            


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES (Adaptadas a Nube)
# ─────────────────────────────────────────────────────────────────────────────

def _semaforo(valor) -> str:
    if pd.isna(valor):
        return "—"
    if valor >= UMBRAL_OK:
        return "✅"
    if valor >= UMBRAL_WARNING:
        return "⚠️"
    return "❌"


def _color_celda(valor) -> str:
    """Color hex (sin #) según umbral."""
    if pd.isna(valor):
        return COLOR_FILA_PAR
    if valor >= UMBRAL_OK:
        return COLOR_OK
    if valor >= UMBRAL_WARNING:
        return COLOR_WARN
    return COLOR_ERROR


def _safe_div(num: pd.Series, den: pd.Series) -> pd.Series:
    """División segura: NaN donde denominador ≤ 0 o NaN."""
    return num.where(den > 0, other=np.nan) / den.where(den > 0, other=np.nan)


def _leer(ruta_sharepoint: str, nombre: str) -> pd.DataFrame:
    """Carga un Excel desde SharePoint utilizando Graph API con manejo descriptivo de error."""
    try:
        print(f"    ⏳ Leyendo [{nombre}] desde SharePoint ({ruta_sharepoint})...")
        token = _obtener_token_azure()
        headers = {"Authorization": f"Bearer {token}"}
        drive_id = _obtener_default_drive_id(token)
        ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
        
        url_codificada = urllib.parse.quote(ruta_limpia)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url_codificada}:/content"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            print(f"    ⚠️  [{nombre}] No se pudo encontrar el archivo en la ruta: {ruta_sharepoint}")
            print(f"        Este módulo no se incluirá en el cálculo.")
            return pd.DataFrame()
            
        df = pd.read_excel(io.BytesIO(response.content))
        df.columns = df.columns.str.strip()
        print(f"    ✓ [{nombre}] {len(df):,} filas cargadas desde la nube")
        return df
    except Exception as e:
        print(f"    ⚠️  [{nombre}] Error al cargar desde la nube: {e}")
        return pd.DataFrame()


def _normalizar_nombre(serie: pd.Series) -> pd.Series:
    """Strip + upper para evitar fallos de merge por espacios o casing."""
    return serie.astype(str).str.strip().str.upper()


def _w_avg(valores: pd.Series, pesos: pd.Series) -> float:
    """
    Promedio ponderado robusto: ignora pares donde el valor es NaN o el
    peso es NaN/0. Devuelve NaN si no hay pares válidos.
    """
    pesos = pesos.fillna(0)
    mask = valores.notna() & (pesos > 0)
    if not mask.any():
        return np.nan
    w = pesos[mask]
    v = valores[mask]
    s = w.sum()
    if s <= 0:
        return np.nan
    return float((v * w).sum() / s)


def _promedio_simple(*vals) -> float:
    """Promedio aritmético ignorando NaN."""
    valid = [v for v in vals if v is not None and not pd.isna(v)]
    return float(sum(valid) / len(valid)) if valid else np.nan


def _coalesce_periodo(df: pd.DataFrame, default_mes: int, default_anio: int) -> tuple[int, int]:
    """Detecta MES y AÑO predominante en un df, con fallback."""
    if "MES" in df.columns and df["MES"].notna().any():
        mes = int(pd.to_numeric(df["MES"], errors="coerce").dropna().mode().iloc[0])
    else:
        mes = default_mes
    if "AÑO" in df.columns and df["AÑO"].notna().any():
        anio = int(pd.to_numeric(df["AÑO"], errors="coerce").dropna().mode().iloc[0])
    else:
        anio = default_anio
    return mes, anio


# ─────────────────────────────────────────────────────────────────────────────
# RESOLUCIÓN DE ACRONIMOS (Base cupos como llave canónica)
# ─────────────────────────────────────────────────────────────────────────────
# Filler usado para armonización por palabras
_PALABRAS_FILLER = {"DE", "LA", "LAS", "LOS", "DEL", "Y"}


def _palabras(s: str) -> set:
    """Tokens significativos de un nombre (excluye fillers)."""
    return {w for w in str(s).upper().split() if w and w not in _PALABRAS_FILLER}


def _resolver_acr_supervisor(nombre_sup: str, nombres_sups_bc: set, idx: dict) -> str:
    """
    Mapea un nombre de supervisor (puede venir truncado) al ACRONIMO
    canónico de Base cupos. Usa armonización por subset de palabras.
    Devuelve "" si no hay match.
    """
    if not nombre_sup or pd.isna(nombre_sup):
        return ""
    s = str(nombre_sup).strip().upper()
    nombre_a_acr = idx["nombre_a_acronimo"]
    if s in nombre_a_acr and s in nombres_sups_bc:
        return nombre_a_acr[s]
    pal_v = _palabras(s)
    if not pal_v:
        return ""
    for nombre_pt in nombres_sups_bc:
        pal_pt = _palabras(nombre_pt)
        if pal_v.issubset(pal_pt) or pal_pt.issubset(pal_v):
            return nombre_a_acr.get(nombre_pt, "")
    return ""


def _enriquecer_acronimo_pt(
    df_calc: pd.DataFrame,
    df_pdv: pd.DataFrame,
    idx: dict,
    nombres_sups_bc: set,
) -> pd.DataFrame:
    """
    Enriquece un DataFrame de cálculo (CIF/NP/Precios/SOS por persona)
    con dos columnas:
        ACRONIMO     → ACRONIMO canónico de Base cupos para esta persona.
        ACRONIMO_SUP → ACRONIMO canónico del supervisor a cargo.

    df_calc: salida de cif_por_gestor / calcular_modulo_simple (gestor side).
             Tiene NOMBRE (y opcional SUPERVISOR_LIDER).
    df_pdv:  el DataFrame fuente que tiene ACRONIMO + CEDULA + NOMBRE.
    """
    if df_calc.empty:
        out = df_calc.copy()
        out["ACRONIMO"]     = ""
        out["ACRONIMO_SUP"] = ""
        return out

    # Recuperar ACRONIMO + CEDULA originales desde la fuente, una fila por NOMBRE
    cols_src = [c for c in ["NOMBRE", "CEDULA", "ACRONIMO"] if c in df_pdv.columns]
    src = df_pdv[cols_src].copy()
    src["NOMBRE"] = src["NOMBRE"].astype(str).str.strip().str.upper()
    if "CEDULA" in src.columns:
        src["CEDULA"] = (
            src["CEDULA"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        )
    if "ACRONIMO" in src.columns:
        src["ACRONIMO"] = src["ACRONIMO"].astype(str).str.strip().str.upper()
    src = src.dropna(subset=["NOMBRE"]).drop_duplicates(subset=["NOMBRE"], keep="first")

    out = df_calc.merge(src, on="NOMBRE", how="left", suffixes=("", "_src"))

    # Resolver ACRONIMO canónico (cascada de fallbacks):
    #   1) ACRONIMO del PT si existe y está en la maestra
    #   2) CEDULA del PT si está en la maestra
    #   3) NOMBRE del PT (último recurso, exact match contra Base cupos)
    if "ACRONIMO" in out.columns:
        cedula_serie = out["CEDULA"] if "CEDULA" in out.columns else None
        out["ACRONIMO"] = bcm.resolver_acronimo(
            out["ACRONIMO"].fillna(""),
            idx,
            "acronimo+cedula",
            df_pt_cedula=cedula_serie if cedula_serie is not None else None,
        )
    else:
        out["ACRONIMO"] = ""

    # Fallback por NOMBRE para filas que aún no resolvieron (típico de NP,
    # cuyo reporte no incluye ACRONIMO ni CEDULA).
    if "NOMBRE" in out.columns:
        nombre_a_acr = idx["nombre_a_acronimo"]
        falta = out["ACRONIMO"].fillna("").astype(str) == ""
        if falta.any():
            out.loc[falta, "ACRONIMO"] = (
                out.loc[falta, "NOMBRE"].astype(str).str.strip().str.upper()
                   .map(nombre_a_acr).fillna("")
            )

    # ACRONIMO_SUP: armonizar SUPERVISOR_LIDER contra supervisores Base cupos
    if "SUPERVISOR_LIDER" in out.columns:
        out["ACRONIMO_SUP"] = out["SUPERVISOR_LIDER"].apply(
            lambda s: _resolver_acr_supervisor(s, nombres_sups_bc, idx)
        )
    else:
        out["ACRONIMO_SUP"] = ""

    return out


def _enriquecer_acronimo_supervisor(
    df_calc: pd.DataFrame,
    idx: dict,
    nombres_sups_bc: set,
) -> pd.DataFrame:
    """
    Para los DataFrames "supervisor side" (cif_por_supervisor,
    calcular_modulo_simple._sup), donde NOMBRE es el nombre del supervisor.
    Resuelve ACRONIMO mediante armonización contra Base cupos.
    """
    if df_calc.empty:
        out = df_calc.copy()
        out["ACRONIMO"] = ""
        return out
    out = df_calc.copy()
    out["ACRONIMO"] = out["NOMBRE"].apply(
        lambda s: _resolver_acr_supervisor(s, nombres_sups_bc, idx)
    )
    return out


# ─────────────────────────────────────────────────────────────────────────────
# A.5) CARGA DE KPIS V3 (Sprint 10 - Adaptado a Nube)
# ─────────────────────────────────────────────────────────────────────────────
# Los ETLs ahora producen *_KPIS.xlsx con los cumplimientos por gestor
# ya pre-calculados (Sprints 7/8/9). Esta función reemplaza el cálculo
# que antes hacía calcular_cumplimientos a partir del detalle crudo.
#
# Formato de salida (consumido por ensamblar_detalle):
#   • df_cif_gest    → ACRONIMO, NOMBRE, SUPERVISOR_LIDER, MES, AÑO,
#                      CIF_COB_%, CIF_FREC_%, CIF_HRS_%, CIF_%,
#                      N_PDVS, VENTA_TOTAL
#   • df_np_gest     → ACRONIMO, NOMBRE, NP_%
#   • df_pr_gest     → ACRONIMO, NOMBRE, PRECIOS_%
#   • df_sos_gest    → ACRONIMO, NOMBRE, SOS_%
#
# La distinción "gestor vs supervisor" se preserva en `ensamblar_detalle`
# mediante el universo de Base cupos. Los KPIs V3 no separan ambos —
# cada persona en KPIS_CIF.xlsx se evalúa por su propio detalle.

def cargar_kpis_v3(
    idx_bc: dict,
    nombres_sups_bc: set,
    mes_filtro: int | None = None,
    anio_filtro: int | None = None,
) -> dict:
    """
    Lee los archivos *_KPIS.xlsx (CIF/SOS/NP/Precios/Exhibiciones) generados
    por los ETLs y los devuelve en el formato que espera ensamblar_detalle.
    Filtra opcionalmente por (mes_filtro, anio_filtro).

    Los ETLs actuales guardan estos archivos en disco LOCAL
    (paths.CIF_OUT_KPIS, paths.SOS_OUT_KPIS, etc. — carpeta SALIDA/<módulo>
    definida en paths.py), no en SharePoint. Por eso se leen primero de ahí.
    Esto funciona igual corriendo en GitHub Actions: los ETLs y este script
    corren como pasos del mismo job, sobre el mismo workspace del runner, así
    que el archivo local que dejó el ETL sigue disponible para este paso.
    Si el archivo local no existe, se intenta como respaldo la misma ruta
    en SharePoint (por si en el futuro los ETLs empiezan a subir estos KPIs
    a la nube con el mismo nombre).
    """
    def _read_kpi(ruta_local, nombre_kpi: str, ruta_cloud: str | None = None) -> pd.DataFrame:
        ruta_local = Path(ruta_local)
        if ruta_local.is_file():
            try:
                print(f"    ⏳ Leyendo KPI [{nombre_kpi}] ({ruta_local})...")
                df = pd.read_excel(ruta_local, engine="openpyxl")
                df.columns = df.columns.str.strip()
                print(f"    ✓ KPI [{nombre_kpi}] {len(df):,} filas cargadas (local)")
                return df
            except Exception as e:
                print(f"    ⚠️  KPI [{nombre_kpi}] Error al leer archivo local ({ruta_local}): {e}")

        if ruta_cloud:
            print(f"    ⏳ KPI [{nombre_kpi}] no está en disco local, probando SharePoint...")
            df = _leer(ruta_cloud, nombre_kpi)
            if not df.empty:
                return df

        print(f"    ⚠️  KPI [{nombre_kpi}] No se encontró (¿ya corriste el ETL de este módulo para el periodo?): {ruta_local}")
        return pd.DataFrame()

    def _filtrar_periodo(df, col_mes="MES", col_anio="AÑO"):
        if df.empty or col_mes not in df.columns or col_anio not in df.columns:
            return df
        df = df.copy()
        df[col_mes]  = pd.to_numeric(df[col_mes],  errors="coerce").fillna(0).astype(int)
        df[col_anio] = pd.to_numeric(df[col_anio], errors="coerce").fillna(0).astype(int)
        if mes_filtro is not None:
            df = df[df[col_mes] == int(mes_filtro)]
        if anio_filtro is not None:
            df = df[df[col_anio] == int(anio_filtro)]
        return df

    # ── CIF (V3): COBERTURA / INTENSIDAD / FRECUENCIA / TOTAL ────────────
    ruta_cif_kpis = paths.CIF_OUT_KPIS
    df_cif_kpi = _filtrar_periodo(_read_kpi(
        ruta_cif_kpis, "CIF",
        ruta_cloud=f"{paths.RUTA_CARPETA_SALIDAS_CIF}/{ruta_cif_kpis.name}",
    ))
    if not df_cif_kpi.empty:
        df_cif_kpi = df_cif_kpi.rename(columns={
            "CUMPLIMIENTO COBERTURA":  "CIF_COB_%",
            "CUMPLIMIENTO INTENSIDAD": "CIF_HRS_%",
            "CUMPLIMIENTO FRECUENCIA": "CIF_FREC_%",
            "TOTAL":                   "CIF_%",
        })
        df_cif_kpi["NOMBRE"] = df_cif_kpi["NOMBRE"].astype(str).str.strip().str.upper()
    df_cif_gest = df_cif_kpi.copy()

    # ── SOS ──────────────────────────────────────────────────────────────
    ruta_sos_kpis = paths.SOS_OUT_KPIS
    df_sos_kpi = _filtrar_periodo(_read_kpi(
        ruta_sos_kpis, "SOS",
        ruta_cloud=f"{paths.RUTA_CARPETA_SALIDAS_SOS}/{ruta_sos_kpis.name}",
    ))
    if not df_sos_kpi.empty:
        df_sos_kpi = df_sos_kpi.rename(columns={"CUMPLIMIENTO": "SOS_%"})
        df_sos_kpi["NOMBRE"] = df_sos_kpi["NOMBRE"].astype(str).str.strip().str.upper()
    df_sos_gest = df_sos_kpi.copy()

    # ── NP (col EJECUCION) ───────────────────────────────────────────────
    ruta_np_kpis = paths.NP_OUT_KPIS
    df_np_kpi = _filtrar_periodo(_read_kpi(
        ruta_np_kpis, "NP",
        ruta_cloud=f"{paths.RUTA_CARPETA_SALIDAS_NP}/{ruta_np_kpis.name}",
    ))
    if not df_np_kpi.empty:
        df_np_kpi = df_np_kpi.rename(columns={"EJECUCION": "NP_%"})
        df_np_kpi["NOMBRE"] = df_np_kpi["NOMBRE"].astype(str).str.strip().str.upper()
    df_np_gest = df_np_kpi.copy()

    # ── PRECIOS ──────────────────────────────────────────────────────────
    ruta_pr_kpis = paths.PR_OUT_KPIS
    df_pr_kpi = _filtrar_periodo(_read_kpi(
        ruta_pr_kpis, "PRECIOS",
        ruta_cloud=f"{paths.RUTA_CARPETA_SALIDAS_PRECIOS}/{ruta_pr_kpis.name}",
    ))
    if not df_pr_kpi.empty:
        df_pr_kpi = df_pr_kpi.rename(columns={"CUMPLIMIENTO": "PRECIOS_%"})
        df_pr_kpi["NOMBRE"] = df_pr_kpi["NOMBRE"].astype(str).str.strip().str.upper()
    df_pr_gest = df_pr_kpi.copy()

    # ── EXHIBICIONES PAGADAS (V3 — Sprint 9 / Sprint 17 captura) ─────────
    ruta_exp_kpis = paths.EXHIB_PAG_OUT_KPIS
    df_exp_kpi = _filtrar_periodo(_read_kpi(
        ruta_exp_kpis, "EXHIB_PAG",
        ruta_cloud=f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/{ruta_exp_kpis.name}",
    ))
    if not df_exp_kpi.empty:
        df_exp_kpi = df_exp_kpi.rename(columns={
            "CUMPLIMIENTO":        "EXHIB_PAG_%",
            "CUMPLIMIENTO_CAPTURA": "EXHIB_PAG_CAPTURA_%",
            "EMPLEADO":            "NOMBRE",
        })
        df_exp_kpi["NOMBRE"] = df_exp_kpi["NOMBRE"].astype(str).str.strip().str.upper()
    df_exp_gest = df_exp_kpi.copy()

    # ── EXHIBICIONES GRATIS (V3 — Sprint 9 / Sprint 17 cumplimiento) ────
    ruta_egr_kpis = paths.EXHIB_GRA_OUT_KPIS
    df_egr_kpi = _filtrar_periodo(_read_kpi(
        ruta_egr_kpis, "EXHIB_GRA",
        ruta_cloud=f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/{ruta_egr_kpis.name}",
    ), col_mes="Mes", col_anio="Año")
    if not df_egr_kpi.empty:
        df_egr_kpi = df_egr_kpi.rename(columns={
            "Mes":           "MES",
            "Año":           "AÑO",
            "Empleado":      "NOMBRE",
            "ALTO IMPACTO":  "EXHIB_GRATIS_ALTO",
            "MEDIO IMPACTO": "EXHIB_GRATIS_MEDIO",
            "TOTAL":         "EXHIB_GRATIS_TOTAL",
            "CUMP_ALTO":     "EXHIB_GRA_ALTO_%",
            "CUMP_MEDIO":    "EXHIB_GRA_MEDIO_%",
        })
        df_egr_kpi["NOMBRE"] = df_egr_kpi["NOMBRE"].astype(str).str.strip().str.upper()
    df_egr_gest = df_egr_kpi.copy()

    # ── Enriquecer cada uno con ACRONIMO usando Base cupos ────────────────
    nombre_a_acr = idx_bc.get("nombre_a_acronimo", {})

    def _add_acronimo(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            df["ACRONIMO"] = ""
            return df
        df = df.copy()
        df["ACRONIMO"] = df["NOMBRE"].map(nombre_a_acr).fillna("")
        falta = df["ACRONIMO"] == ""
        if falta.any():
            df.loc[falta, "ACRONIMO"] = df.loc[falta, "NOMBRE"].apply(
                lambda n: _resolver_acr_supervisor(n, nombres_sups_bc, idx_bc)
            )
        return df

    df_cif_gest = _add_acronimo(df_cif_gest)
    df_sos_gest = _add_acronimo(df_sos_gest)
    df_np_gest  = _add_acronimo(df_np_gest)
    df_pr_gest  = _add_acronimo(df_pr_gest)
    df_exp_gest = _add_acronimo(df_exp_gest)
    df_egr_gest = _add_acronimo(df_egr_gest)

    # ── ACRONIMO_SUP: armonizar SUPERVISOR_LIDER vs Base cupos ───────────
    def _add_acr_sup(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty or "SUPERVISOR_LIDER" not in df.columns:
            if not df.empty:
                df["ACRONIMO_SUP"] = ""
            return df
        df = df.copy()
        df["ACRONIMO_SUP"] = df["SUPERVISOR_LIDER"].apply(
            lambda s: _resolver_acr_supervisor(s, nombres_sups_bc, idx_bc)
        )
        return df

    df_cif_gest = _add_acr_sup(df_cif_gest)
    df_sos_gest = _add_acr_sup(df_sos_gest)
    df_np_gest  = _add_acr_sup(df_np_gest)
    df_pr_gest  = _add_acr_sup(df_pr_gest)
    
    acr_a_sup: dict = {}
    for src in (df_cif_gest, df_np_gest, df_pr_gest, df_sos_gest):
        if src is None or src.empty:
            continue
        if "ACRONIMO" not in src.columns or "ACRONIMO_SUP" not in src.columns:
            continue
        for _, r in src[["ACRONIMO", "ACRONIMO_SUP"]].iterrows():
            acr, asu = r["ACRONIMO"], r["ACRONIMO_SUP"]
            if acr and asu and acr not in acr_a_sup:
                acr_a_sup[acr] = asu

    def _add_acr_sup_via_acr(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        df = df.copy()
        df["ACRONIMO_SUP"] = df["ACRONIMO"].map(acr_a_sup).fillna("")
        return df

    df_exp_gest = _add_acr_sup_via_acr(df_exp_gest)
    df_egr_gest = _add_acr_sup_via_acr(df_egr_gest)

    return {
        "cif_gest": df_cif_gest,
        "sos_gest": df_sos_gest,
        "np_gest":  df_np_gest,
        "pr_gest":  df_pr_gest,
        "exp_gest": df_exp_gest,   
        "egr_gest": df_egr_gest,   
    }


# ─────────────────────────────────────────────────────────────────────────────
# A) GENERAR / VERIFICAR MAESTRA DE SUPERVISORES (Adaptado a Nube)
# ─────────────────────────────────────────────────────────────────────────────

def generar_maestra(df_np: pd.DataFrame, df_pr: pd.DataFrame, df_sos: pd.DataFrame) -> None:
    """
    Genera o actualiza MAESTRO_SUPERVISORES.xlsx desde Base cupos en la nube.
    """
    DIR_ALERTAS.mkdir(parents=True, exist_ok=True)

    df_bc   = bcm.cargar()
    sups_bc = bcm.supervisores(df_bc)
    nombres_bc = sorted(sups_bc["NOMBRE"].dropna().unique())

    existente = pd.DataFrame(columns=["NOMBRE_SUPERVISOR", "CORREO", "TELEGRAM_CHAT_ID"])

    # Ruta de la maestra en SharePoint — definida SIEMPRE antes del try para
    # que esté disponible más abajo incluso si la lectura falla.
    ruta_maestra_cloud = RUTA_MAESTRA_CLOUD

    # Lectura de la maestra desde SharePoint en lugar de disco local
    try:
        print(f"    ⏳ Leyendo maestra de supervisores desde SharePoint...")
        token = _obtener_token_azure()
        headers = {"Authorization": f"Bearer {token}"}
        drive_id = _obtener_default_drive_id(token)
        ruta_limpia = _limpiar_ruta_graph(ruta_maestra_cloud)
        
        url_codificada = urllib.parse.quote(ruta_limpia)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url_codificada}:/content"
        
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            existente = pd.read_excel(io.BytesIO(response.content))
            existente.columns = existente.columns.str.strip().str.upper()
            existente["NOMBRE_SUPERVISOR"] = (
                existente["NOMBRE_SUPERVISOR"].astype(str).str.strip().str.upper()
            )
            existente = existente[
                existente["NOMBRE_SUPERVISOR"].notna()
                & (existente["NOMBRE_SUPERVISOR"] != "")
                & (existente["NOMBRE_SUPERVISOR"] != "NAN")
                & (~existente["NOMBRE_SUPERVISOR"].str.contains(r"[⚠✓✗]", na=False))
                & (~existente["NOMBRE_SUPERVISOR"].str.contains(r"^COMPLETA", na=False))
            ].reset_index(drop=True)
    except Exception as e:
        print(f"    ⚠️  No pude leer la maestra existente desde la nube ({e}); recreando")
        existente = pd.DataFrame(columns=["NOMBRE_SUPERVISOR", "CORREO", "TELEGRAM_CHAT_ID"])

    nombres_existentes = set(existente["NOMBRE_SUPERVISOR"]) if not existente.empty else set()
    nombres_nuevos = [n for n in nombres_bc if n not in nombres_existentes]
    nombres_obsoletos = [
        n for n in nombres_existentes
        if n and n != "NAN" and n not in set(nombres_bc)
    ]

    try:
        crudo_n = len(existente)
    except Exception:
        crudo_n = len(existente)
    hay_basura = crudo_n > len(existente)

    if not nombres_nuevos and not hay_basura and not existente.empty:
        if nombres_obsoletos:
            print(
                f"  ⚠️  {len(nombres_obsoletos)} supervisor(es) en la maestra ya no "
                f"figuran como activos en Base cupos:"
            )
            for n in nombres_obsoletos[:5]:
                print(f"     · {n}")
            print(f"     (Se conservan en la maestra; revisa si fueron desactivados.)")
        print(f"  ✓ Maestra al día: {len(existente)} supervisores")
        return

    if not existente.empty:
        existente = existente[["NOMBRE_SUPERVISOR", "CORREO", "TELEGRAM_CHAT_ID"]].copy()
        nuevos_df = pd.DataFrame({
            "NOMBRE_SUPERVISOR": nombres_nuevos,
            "CORREO":            [""] * len(nombres_nuevos),
            "TELEGRAM_CHAT_ID":  [""] * len(nombres_nuevos),
        })
        df_maestra = pd.concat([existente, nuevos_df], ignore_index=True)
        df_maestra = df_maestra.sort_values("NOMBRE_SUPERVISOR").reset_index(drop=True)
        accion = (
            f"✏️  Maestra actualizada: {len(existente)} existentes + "
            f"{len(nombres_nuevos)} nuevos (correo/chat_id vacíos)"
        )
    else:
        df_maestra = pd.DataFrame({
            "NOMBRE_SUPERVISOR": nombres_bc,
            "CORREO":            [""] * len(nombres_bc),
            "TELEGRAM_CHAT_ID":  [""] * len(nombres_bc),
        })
        accion = f"📋 Maestra creada con {len(nombres_bc)} supervisores"

    print(f"\n  {accion}")
    if nombres_nuevos:
        print(f"     Nuevos:")
        for n in nombres_nuevos[:10]:
            print(f"       · {n}")
        if len(nombres_nuevos) > 10:
            print(f"       · ... ({len(nombres_nuevos) - 10} más)")

    wb = Workbook()
    ws = wb.active
    ws.title = "Supervisores"

    headers = ["NOMBRE_SUPERVISOR", "CORREO", "TELEGRAM_CHAT_ID"]
    anchos  = [40, 40, 20]

    hdr_font  = Font(name="Arial", bold=True, color=COLOR_HEADER_FNT, size=11)
    hdr_fill  = PatternFill("solid", fgColor=COLOR_MAESTRA_HDR)
    hdr_align = Alignment(horizontal="center", vertical="center")
    thin      = Side(style="thin", color="CCCCCC")
    borde     = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, (hdr, ancho) in enumerate(zip(headers, anchos), start=1):
        cell = ws.cell(row=1, column=col_idx, value=hdr)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = hdr_align
        cell.border    = borde
        ws.column_dimensions[get_column_letter(col_idx)].width = ancho
    ws.row_dimensions[1].height = 22

    fill_par = PatternFill("solid", fgColor=COLOR_FILA_PAR)
    for row_idx, row in enumerate(df_maestra.itertuples(index=False), start=2):
        fill = fill_par if row_idx % 2 == 0 else PatternFill()
        for col_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = fill
            cell.alignment = Alignment(vertical="center")
            cell.border    = borde

    ws.freeze_panes = "A2"
    # El nombre de una pestaña de Excel tiene un límite DURO de 31
    # caracteres — "Supervisores - Completa CORREO y TELEGRAM_CHAT_ID"
    # (50 caracteres) lo excede y produce un .xlsx inválido que Excel
    # reporta como dañado/corrupto al abrirlo. Se deja el título corto y
    # la instrucción se agrega como nota en una celda en vez de en el título.
    ws.cell(row=1, column=5,
            value="👉 Completa CORREO y TELEGRAM_CHAT_ID para los supervisores nuevos")

    # Guardar en la ruta local temporal y luego subir a SharePoint (o guardar directo si se usa BytesIO)
    output_io = io.BytesIO()
    wb.save(output_io)
    output_io.seek(0)
    
    # Subida a SharePoint usando Graph API
    try:
        token = _obtener_token_azure()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        drive_id = _obtener_default_drive_id(token)
        ruta_limpia = _limpiar_ruta_graph(ruta_maestra_cloud)
        
        url_codificada = urllib.parse.quote(ruta_limpia)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url_codificada}:/content"
        
        response = requests.put(url, headers=headers, data=output_io.getvalue())
        if response.status_code in [200, 201]:
            print(f"  ✅ Maestra guardada en SharePoint con {len(df_maestra)} supervisores: {ruta_maestra_cloud}")
        else:
            print(f"  ⚠️ Error al subir la maestra a SharePoint: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"  ⚠️ Excepción al subir la maestra a SharePoint: {e}")

    if nombres_nuevos:
        print(f"     → Completa CORREO y TELEGRAM_CHAT_ID para los {len(nombres_nuevos)} nuevos.\n")


# ─────────────────────────────────────────────────────────────────────────────
# B) CÁLCULO DE KPIS  →  movido a los ETLs (V3, Sprint 7/8/9).
#    Antes vivía aquí (calcular_cif_componentes_por_pdv, _cif_agregar_por_clave,
#    cif_por_gestor, cif_por_supervisor, calcular_modulo_simple) — ~310 líneas.
#    Hoy cada ETL escribe su <MOD>_KPIS.xlsx y este módulo solo los ensambla
#    via cargar_kpis_v3() (definido arriba en la sección A.5).
# ─────────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# C) ENSAMBLAJE FINAL — df_detalle y df_resumen
# ─────────────────────────────────────────────────────────────────────────────

def ensamblar_detalle(
    universo: pd.DataFrame,
    df_cif_gest: pd.DataFrame,
    df_cif_sup:  pd.DataFrame,
    df_np_gest:  pd.DataFrame, df_np_sup:  pd.DataFrame,
    df_pr_gest:  pd.DataFrame, df_pr_sup:  pd.DataFrame,
    df_sos_gest: pd.DataFrame, df_sos_sup: pd.DataFrame,
    df_dyp_gest: pd.DataFrame | None = None,
    df_dyp_sup:  pd.DataFrame | None = None,
    df_exp_gest: pd.DataFrame | None = None,
    df_egr_gest: pd.DataFrame | None = None,
    mes: int | None = None,
    anio: int | None = None,
) -> pd.DataFrame:
    """
    Una fila por persona del UNIVERSO (Base cupos), keyed por ACRONIMO.

    Columnas:
      ACRONIMO, NOMBRE, CEDULA, ROL, CANAL, ES_SUPERVISOR,
      ACRONIMO_SUP, NOMBRE_SUPERVISOR,
      N_PDVS, VENTA_TOTAL,
      CIF_COB_%, CIF_FREC_%, CIF_HRS_%, CIF_%,
      NP_%, PRECIOS_%, SOS_%,
      VENTA_%, IMPACTOS_%, MSL_%, PROD_NUEVOS_%,
      MES, AÑO

    Reglas de fusión:
      • Universo = `bcm.universo_personas(df_bc)` (200 personas activas).
      • Para cada DataFrame de cálculo (gestor/supervisor), merge por ACRONIMO.
      • `ACRONIMO_SUP` por persona se toma con prioridad: D&P → CIF →
        otros (NP/Precios/SOS), reflejando que D&P es la fuente más
        completa para transferencistas puros.
      • `NOMBRE_SUPERVISOR` se deriva del ACRONIMO_SUP cruzando con el
        universo.
    """
    # ── BASE: universo de personas (de Base cupos) ───────────────────────
    base = universo.copy()
    # Renombrar para usar nombres canónicos del PT en el output
    # (NOMBRE ya es UPPER por bcm.cargar)
    if mes is not None:
        base["MES"] = mes
    if anio is not None:
        base["AÑO"] = anio

    # Helper local: agregar/upsert columnas desde un df_calc cruzando por ACRONIMO
    def _merge_por_acr(target, df_calc, cols_traer):
        if df_calc is None or df_calc.empty:
            for c in cols_traer:
                if c not in target.columns:
                    target[c] = np.nan
            return target
        cols_disp = [c for c in cols_traer if c in df_calc.columns]
        if not cols_disp or "ACRONIMO" not in df_calc.columns:
            for c in cols_traer:
                if c not in target.columns:
                    target[c] = np.nan
            return target
        merged = (
            df_calc[df_calc["ACRONIMO"] != ""][["ACRONIMO"] + cols_disp]
                  .drop_duplicates(subset=["ACRONIMO"])
        )
        return target.merge(merged, on="ACRONIMO", how="left")

    # ── CIF (gestor + supervisor) ────────────────────────────────────────
    cif_g_ren = pd.DataFrame()
    if not df_cif_gest.empty:
        cif_g_ren = df_cif_gest.rename(columns={
            "COB":  "CIF_COB_%", "FREC": "CIF_FREC_%",
            "HRS":  "CIF_HRS_%", "CIF":  "CIF_%",
        })
    cif_s_ren = pd.DataFrame()
    if not df_cif_sup.empty:
        cif_s_ren = df_cif_sup.rename(columns={
            "COB":  "CIF_COB_%", "FREC": "CIF_FREC_%",
            "HRS":  "CIF_HRS_%", "CIF":  "CIF_%",
        })
    # Concatenar gestor + supervisor (cada persona aparece en uno u otro)
    cif_unif = pd.concat([cif_g_ren, cif_s_ren], ignore_index=True) \
                  if (not cif_g_ren.empty or not cif_s_ren.empty) else pd.DataFrame()
    cols_cif = ["CIF_COB_%", "CIF_FREC_%", "CIF_HRS_%", "CIF_%",
                "N_PDVS", "VENTA_TOTAL"]
    base = _merge_por_acr(base, cif_unif, cols_cif)

    # Si N_PDVS/VENTA_TOTAL no llegaron, asegurar columnas
    for c in ["N_PDVS", "VENTA_TOTAL"]:
        if c not in base.columns:
            base[c] = np.nan

    # ── NP, Precios, SOS — unificar gestor+supervisor de cada uno ────────
    for df_g, df_s, col in [
        (df_np_gest,  df_np_sup,  "NP_%"),
        (df_pr_gest,  df_pr_sup,  "PRECIOS_%"),
        (df_sos_gest, df_sos_sup, "SOS_%"),
    ]:
        unif = pd.concat(
            [df for df in (df_g, df_s) if df is not None and not df.empty],
            ignore_index=True,
        ) if any(df is not None and not df.empty for df in (df_g, df_s)) else pd.DataFrame()
        base = _merge_por_acr(base, unif, [col])

    # ── EXHIBICIONES (Sprint 12 / Sprint 17) ─────────────────────────────
    # KPIs que entran al CUMPL_GLOBAL: EXHIB_PAG_% (ejecución), EXHIB_PAG_CAPTURA_%,
    # EXHIB_GRA_ALTO_% (vs target 3), EXHIB_GRA_MEDIO_% (vs target 5).
    # Conteos absolutos informativos (no entran al global):
    # EXHIB_GRATIS_TOTAL / ALTO / MEDIO.
    if df_exp_gest is not None:
        base = _merge_por_acr(base, df_exp_gest,
                              ["EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%"])
    if df_egr_gest is not None:
        base = _merge_por_acr(
            base, df_egr_gest,
            ["EXHIB_GRATIS_TOTAL", "EXHIB_GRATIS_ALTO", "EXHIB_GRATIS_MEDIO",
             "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%"],
        )

    # ── D&P (gestor + supervisor) ────────────────────────────────────────
    cols_dyp = ["VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]
    dyp_unif = pd.concat(
        [df for df in (df_dyp_gest, df_dyp_sup) if df is not None and not df.empty],
        ignore_index=True,
    ) if (df_dyp_gest is not None and not df_dyp_gest.empty) or \
         (df_dyp_sup  is not None and not df_dyp_sup.empty) else pd.DataFrame()
    base = _merge_por_acr(base, dyp_unif, cols_dyp)

    # ── ACRONIMO_SUP por persona (prioridad: D&P > CIF > otros) ──────────
    # Construimos un mapa ACRONIMO → ACRONIMO_SUP a partir de cada fuente.
    sup_map = {}

    def _acumular(df, prioridad):
        if df is None or df.empty:
            return
        if "ACRONIMO" not in df.columns or "ACRONIMO_SUP" not in df.columns:
            return
        for _, r in df.iterrows():
            acr = r.get("ACRONIMO", "")
            asu = r.get("ACRONIMO_SUP", "")
            if not acr or not asu:
                continue
            existente = sup_map.get(acr)
            if existente is None or existente[0] > prioridad:
                sup_map[acr] = (prioridad, asu)

    # Prioridad menor = más prioritario
    _acumular(df_dyp_gest, prioridad=1)        # D&P primero (gestores)
    _acumular(df_cif_gest if not df_cif_gest.empty else None, prioridad=2)
    _acumular(df_np_gest  if not df_np_gest.empty  else None, prioridad=3)
    _acumular(df_pr_gest  if not df_pr_gest.empty  else None, prioridad=4)
    _acumular(df_sos_gest if not df_sos_gest.empty else None, prioridad=5)

    base["ACRONIMO_SUP"] = base["ACRONIMO"].map(
        lambda a: sup_map[a][1] if a in sup_map else ""
    )
    # NOMBRE_SUPERVISOR derivado del universo
    nombre_por_acr = dict(zip(universo["ACRONIMO"], universo["NOMBRE"]))
    base["NOMBRE_SUPERVISOR"] = base["ACRONIMO_SUP"].map(nombre_por_acr).fillna("")

    # ── Reordenar columnas finales ───────────────────────────────────────
    cols_finales = [
        "ACRONIMO", "NOMBRE", "CEDULA", "ROL", "CANAL", "CIUDAD_CUPO",
        "ES_SUPERVISOR", "ES_GDD", "ES_LIDER",         # Sprint 17.15
        "ACRONIMO_SUP", "NOMBRE_SUPERVISOR",
        "N_PDVS", "VENTA_TOTAL",
        "CIF_COB_%", "CIF_FREC_%", "CIF_HRS_%", "CIF_%",
        "NP_%", "PRECIOS_%", "SOS_%",
        # Exhibiciones (Sprint 12 / Sprint 17)
        "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
        "EXHIB_GRATIS_TOTAL", "EXHIB_GRATIS_ALTO", "EXHIB_GRATIS_MEDIO",
        "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
        "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%",
        "MES", "AÑO",
    ]
    for c in cols_finales:
        if c not in base.columns:
            base[c] = np.nan
    base = base[cols_finales]

    # Orden: supervisores primero por ACRONIMO; gestores agrupados por su sup.
    base = base.sort_values(
        ["ES_SUPERVISOR", "ACRONIMO_SUP", "NOMBRE"],
        ascending=[False, True, True],
    ).reset_index(drop=True)

    return base


def calcular_resumen(df_detalle: pd.DataFrame, universo: pd.DataFrame) -> pd.DataFrame:
    """
    Una fila por supervisor:
      ACRONIMO_SUP, SUPERVISOR_LIDER (= NOMBRE del supervisor),
      N_GESTORES, VENTA_EQUIPO,
      CIF_%, NP_%, PRECIOS_%, SOS_%,
      VENTA_%, IMPACTOS_%, MSL_%, PROD_NUEVOS_%,
      CUMPL_GLOBAL_%,
      MES, AÑO

    Los KPIs del supervisor se toman de la fila propia (ES_SUPERVISOR=True)
    en df_detalle. N_GESTORES y VENTA_EQUIPO se calculan agrupando por
    ACRONIMO_SUP las filas no-supervisor.
    """
    if df_detalle.empty:
        return pd.DataFrame()

    # ── Stats del equipo: agrupar gestores por ACRONIMO_SUP ──────────────
    no_sup = df_detalle[df_detalle["ES_SUPERVISOR"] == False].copy()
    if not no_sup.empty:
        equipo_stats = (
            no_sup[no_sup["ACRONIMO_SUP"].fillna("").astype(str) != ""]
                  .groupby("ACRONIMO_SUP", dropna=False)
                  .agg(
                      N_GESTORES   = ("ACRONIMO",    "nunique"),
                      VENTA_EQUIPO = ("VENTA_TOTAL", "sum"),
                  )
                  .reset_index()
                  .rename(columns={"ACRONIMO_SUP": "ACRONIMO"})
        )
    else:
        equipo_stats = pd.DataFrame(columns=["ACRONIMO", "N_GESTORES", "VENTA_EQUIPO"])

    # ── Filas SUPERVISOR de df_detalle (sus propios KPIs) ────────────────
    sups = df_detalle[df_detalle["ES_SUPERVISOR"] == True].copy()

    # Sprint 12 / Sprint 17: entran al CUMPL_GLOBAL_% (D11=Sí):
    #   EXHIB_PAG_% (ejecución), EXHIB_PAG_CAPTURA_% (captura del planning),
    #   EXHIB_GRA_ALTO_% (vs target 3), EXHIB_GRA_MEDIO_% (vs target 5).
    # Conteos informativos (no global): EXHIB_GRATIS_TOTAL/ALTO/MEDIO.
    cols_pct = ["CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
                "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
                "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
                "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]
    cols_info = ["EXHIB_GRATIS_TOTAL", "EXHIB_GRATIS_ALTO", "EXHIB_GRATIS_MEDIO"]
    cols_sup = ["ACRONIMO", "NOMBRE", "MES", "AÑO"] + cols_pct + cols_info
    sup_metrics = sups[[c for c in cols_sup if c in sups.columns]].copy()

    # ── KPI agregado por equipo (Sprint 10) ──────────────────────────────
    # Los KPIs V3 solo poblan al supervisor cuando él mismo tiene ruta
    # (típico en CIF). Para NP/Precios/SOS/D&P/EXHIB, el supervisor casi
    # nunca aparece como persona con KPI propio. Tomamos el promedio del
    # equipo de sus gestores como fallback razonable (suma para conteos
    # informativos como EXHIB_GRATIS_TOTAL).
    if not no_sup.empty:
        equipo_filtered = no_sup[no_sup["ACRONIMO_SUP"].fillna("").astype(str) != ""]
        team_kpis_pct = (
            equipo_filtered.groupby("ACRONIMO_SUP", dropna=False)[cols_pct]
                           .mean()
                           .reset_index()
        )
        cols_info_disp = [c for c in cols_info if c in no_sup.columns]
        if cols_info_disp:
            team_kpis_info = (
                equipo_filtered.groupby("ACRONIMO_SUP", dropna=False)[cols_info_disp]
                               .sum(min_count=1)
                               .reset_index()
            )
            team_kpis = team_kpis_pct.merge(team_kpis_info, on="ACRONIMO_SUP")
        else:
            team_kpis = team_kpis_pct
        team_kpis = team_kpis.rename(columns={"ACRONIMO_SUP": "ACRONIMO"})
    else:
        team_kpis = pd.DataFrame(columns=["ACRONIMO"] + cols_pct + cols_info)

    # Combinar: KPI propio del supervisor > promedio del equipo
    sup_metrics_combinado = sup_metrics.merge(
        team_kpis, on="ACRONIMO", how="outer", suffixes=("", "_eq")
    )
    # Para %s, prevalece KPI propio sobre el de equipo
    for c in cols_pct + cols_info:
        col_eq = f"{c}_eq"
        if col_eq in sup_metrics_combinado.columns:
            sup_metrics_combinado[c] = (
                sup_metrics_combinado[c].where(
                    sup_metrics_combinado[c].notna(),
                    sup_metrics_combinado[col_eq],
                )
            )
            sup_metrics_combinado = sup_metrics_combinado.drop(columns=[col_eq])
    sup_metrics = sup_metrics_combinado

    # Merge externo: para no perder supervisores sin equipo ni equipos cuyo
    # supervisor no esté en el universo.
    resumen = pd.merge(sup_metrics, equipo_stats, on="ACRONIMO", how="outer")

    # Para filas que vinieron del equipo_stats (sup sin métricas propias),
    # rellenar NOMBRE desde el universo.
    nombre_por_acr = dict(zip(universo["ACRONIMO"], universo["NOMBRE"]))
    if "NOMBRE" in resumen.columns:
        resumen["NOMBRE"] = resumen["NOMBRE"].fillna(
            resumen["ACRONIMO"].map(nombre_por_acr)
        )
    else:
        resumen["NOMBRE"] = resumen["ACRONIMO"].map(nombre_por_acr)

    # Defensivo: completar columnas faltantes
    for c in cols_pct + cols_info + ["N_GESTORES", "VENTA_EQUIPO", "MES", "AÑO"]:
        if c not in resumen.columns:
            resumen[c] = np.nan
    resumen["N_GESTORES"]   = resumen["N_GESTORES"].fillna(0).astype(int)
    resumen["VENTA_EQUIPO"] = resumen["VENTA_EQUIPO"].fillna(0.0)

    # Cumplimiento global (promedio simple de los KPIs % disponibles).
    # EXHIB_PAG_% SÍ entra (es un KPI), EXHIB_GRATIS_TOTAL NO (es conteo).
    resumen["CUMPL_GLOBAL_%"] = resumen[cols_pct].mean(axis=1, skipna=True)

    # Renombrar para output legible
    resumen = resumen.rename(columns={
        "ACRONIMO": "ACRONIMO_SUP",
        "NOMBRE":   "SUPERVISOR_LIDER",
    })

    cols_finales = [
        "ACRONIMO_SUP", "SUPERVISOR_LIDER", "N_GESTORES", "VENTA_EQUIPO",
        "CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
        "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
        "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
        "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%",
        "EXHIB_GRATIS_TOTAL", "EXHIB_GRATIS_ALTO", "EXHIB_GRATIS_MEDIO",
        "CUMPL_GLOBAL_%",
        "MES", "AÑO",
    ]
    for c in cols_finales:
        if c not in resumen.columns:
            resumen[c] = np.nan
    resumen = resumen[cols_finales]

    # Sprint 17.15 — agregar destinatarios GDD y Líder de Ejecución.
    resumen["_ROL_DEST"] = "SUPERVISOR"
    resumen_gdd_lid = _filas_gdd_y_lider(df_detalle, universo, cols_finales)
    if not resumen_gdd_lid.empty:
        resumen = pd.concat([resumen, resumen_gdd_lid], ignore_index=True)

    return resumen.sort_values("SUPERVISOR_LIDER").reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# Sprint 17.15 — destinatarios adicionales: GDD y Líder de Ejecución
# ─────────────────────────────────────────────────────────────────────────────

# Mapping ciudad → ACRONIMO del líder (canal DIRECTO y combinados).
LIDER_POR_CIUDAD: dict[str, str] = {
    "CALI": "LE001", "PEREIRA": "LE001", "PASTO": "LE001",
    "BUCARAMANGA": "LE001", "CUCUTA": "LE001",
    "BOGOTA": "LE002", "NEIVA": "LE002", "SOGAMOSO": "LE002",
    "MEDELLIN": "LE003", "BARRANQUILLA": "LE003",
    "CARTAGENA": "LE003", "VALLEDUPAR": "LE003", "MONTERIA": "LE003",
}


def _filas_gdd_y_lider(
    df_detalle: pd.DataFrame,
    universo: pd.DataFrame,
    cols_finales: list[str],
) -> pd.DataFrame:
    """
    Genera filas de "resumen" adicionales para:
      • GDDs (1 fila c/u con sus mediciones propias del df_detalle).
      • Líderes (1 fila c/u con sus mediciones propias + promedios de los
        supervisores asignados según LIDER_POR_CIUDAD).
    Esquema idéntico al resumen de supervisor.
    """
    cols_pct = ["CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
                "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
                "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
                "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]
    cols_info = ["EXHIB_GRATIS_TOTAL", "EXHIB_GRATIS_ALTO", "EXHIB_GRATIS_MEDIO"]

    if df_detalle.empty or universo.empty:
        return pd.DataFrame(columns=cols_finales)

    mes_v  = df_detalle["MES"].dropna().iloc[0]  if "MES" in df_detalle.columns and df_detalle["MES"].notna().any() else np.nan
    anio_v = df_detalle["AÑO"].dropna().iloc[0]  if "AÑO" in df_detalle.columns and df_detalle["AÑO"].notna().any() else np.nan

    # Mapa rápido ACRONIMO → fila del detalle (KPIs propios).
    det_idx = df_detalle.set_index("ACRONIMO", drop=False)

    def _fila_persona(acr: str, n_gestores: int = 0, venta_eq: float = 0.0) -> dict:
        """Construye una fila estilo resumen con los KPIs de la persona."""
        fila = {c: np.nan for c in cols_finales}
        per = (universo[universo["ACRONIMO"] == acr].iloc[0]
               if (universo["ACRONIMO"] == acr).any() else None)
        fila["ACRONIMO_SUP"]     = acr
        fila["SUPERVISOR_LIDER"] = per["NOMBRE"] if per is not None else acr
        fila["N_GESTORES"]       = int(n_gestores)
        fila["VENTA_EQUIPO"]     = float(venta_eq)
        fila["MES"]  = mes_v
        fila["AÑO"]  = anio_v
        if acr in det_idx.index:
            row = det_idx.loc[acr]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            for c in cols_pct + cols_info:
                if c in row.index and pd.notna(row[c]):
                    fila[c] = row[c]
        # CUMPL_GLOBAL_% = promedio de KPIs disponibles.
        vals = [fila[c] for c in cols_pct if pd.notna(fila.get(c))]
        fila["CUMPL_GLOBAL_%"] = float(np.mean(vals)) if vals else np.nan
        return fila

    filas: list[dict] = []

    # ── GDDs ──────────────────────────────────────────────────────────────
    gdds = universo[universo["ES_GDD"] == True]
    for _, p in gdds.iterrows():
        f = _fila_persona(p["ACRONIMO"])
        f["_ROL_DEST"] = "GDD"
        filas.append(f)

    # ── Líderes ───────────────────────────────────────────────────────────
    lideres = universo[universo["ES_LIDER"] == True]
    # Sprint 17.18 — solo supervisores cuyo CANAL incluya DIRECTO entran al
    # equipo de un líder (DIRECTO, DIRECTO-DROGUERIAS, DIRECTO-PROXIMITY,
    # DIRECTO-DROGUERIAS-PROXIMITY). DROGUERIAS, PROXIMITY, PROXIMITY-TAT,
    # PROXIMITY-DROGUERIAS quedan FUERA.
    sups_univ = universo[universo["ES_SUPERVISOR"] == True].copy()
    sups_univ["CIUDAD_U"] = sups_univ["CIUDAD_CUPO"].astype(str).str.upper().str.strip()
    sups_univ["LIDER_ACR"] = sups_univ["CIUDAD_U"].map(LIDER_POR_CIUDAD).fillna("")
    sups_univ["CANAL_U"]  = sups_univ["CANAL"].astype(str).str.upper()
    sups_univ = sups_univ[sups_univ["CANAL_U"].str.contains("DIRECTO", na=False)]

    for _, lid in lideres.iterrows():
        acr_lid = lid["ACRONIMO"]
        # Supervisores asignados a este líder (ya filtrados a los que tienen DIRECTO).
        sups_asig = sups_univ[sups_univ["LIDER_ACR"] == acr_lid]["ACRONIMO"].tolist()
        fila = _fila_persona(acr_lid)
        # Sprint 17.18 — para el LIDER, las métricas son el promedio del
        # EQUIPO de cada supervisor asignado (gestores + supervisor mismo),
        # no solo los KPIs propios del supervisor.
        if sups_asig:
            # Equipo completo: gestores cuyo ACRONIMO_SUP esté en sups_asig
            # + las filas propias de los supervisores asignados.
            mask_eq = (
                df_detalle["ACRONIMO_SUP"].isin(sups_asig) |
                df_detalle["ACRONIMO"].isin(sups_asig)
            )
            equipo_filas = df_detalle[mask_eq]
            if not equipo_filas.empty:
                for c in cols_pct:
                    if c in equipo_filas.columns:
                        v = pd.to_numeric(equipo_filas[c], errors="coerce")
                        if c in ("EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%"):
                            v = v.clip(upper=1.0)   # cap consistencia HTML
                        v = v.dropna()
                        if not v.empty:
                            fila[c] = float(v.mean())
                for c in cols_info:
                    if c in equipo_filas.columns:
                        v = equipo_filas[c].dropna()
                        if not v.empty:
                            fila[c] = float(v.sum())
                # N_GESTORES para LIDER = # supervisores a cargo (textos
                # "n supervisores a cargo").
                fila["N_GESTORES"] = int(len(sups_asig))
                vals = [fila[c] for c in cols_pct if pd.notna(fila.get(c))]
                fila["CUMPL_GLOBAL_%"] = float(np.mean(vals)) if vals else np.nan
        fila["_ROL_DEST"] = "LIDER"
        filas.append(fila)

    if not filas:
        return pd.DataFrame(columns=cols_finales + ["_ROL_DEST"])
    return pd.DataFrame(filas)[cols_finales + ["_ROL_DEST"]]


# ─────────────────────────────────────────────────────────────────────────────
# D) ESCRITURA DE EXCEL CON FORMATO
# ─────────────────────────────────────────────────────────────────────────────

def _aplicar_formato_tabla(
    ws,
    df: pd.DataFrame,
    cols_pct: list[str],
    cols_venta: list[str] | None = None,
    fila_inicio_datos: int = 2,
) -> None:
    """
    Aplica formato profesional a una hoja:
      • encabezado azul oscuro
      • filas alternas blanco/gris
      • columnas en `cols_pct` con formato 0.0% y color de semáforo
      • columnas en `cols_venta` con formato $#,##0
    """
    cols_venta = cols_venta or []
    thin   = Side(style="thin", color="CCCCCC")
    borde  = Border(left=thin, right=thin, top=thin, bottom=thin)

    hdr_font  = Font(name="Arial", bold=True, color=COLOR_HEADER_FNT, size=10)
    hdr_fill  = PatternFill("solid", fgColor=COLOR_HEADER)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Encabezados (en fila inmediatamente anterior a fila_inicio_datos)
    fila_hdr = fila_inicio_datos - 1
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=fila_hdr, column=col_idx, value=col_name)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = hdr_align
        cell.border    = borde
    ws.row_dimensions[fila_hdr].height = 30

    # Datos
    # Sprint 17.17 — cap a 100% en columnas Exh Gratis Alto/Medio para
    # consistencia con el HTML del cuerpo del correo.
    cols_cap_a_100 = {
        "Cump Exhibiciones Gratis Alto Impacto",
        "Cump Exhibiciones Gratis Medio Impacto",
        "CUMPLIMIENTO EXH. GRATIS ALTO",      # nombres legacy por si quedan
        "CUMPLIMIENTO EXH. GRATIS MEDIO",
        "EXHIB_GRA_ALTO_%",
        "EXHIB_GRA_MEDIO_%",
    }

    fill_par = PatternFill("solid", fgColor=COLOR_FILA_PAR)
    for offset, row in enumerate(df.itertuples(index=False)):
        row_idx = fila_inicio_datos + offset
        fill_base = fill_par if row_idx % 2 == 0 else PatternFill()
        for col_idx, (col_name, val) in enumerate(zip(df.columns, row), start=1):
            # openpyxl no puede serializar pd.NA ni pd.NaT, y NaN lo escribe
            # como cadena rara. Convertimos cualquier "missing" a None.
            if val is pd.NA:
                val = None
            else:
                try:
                    if pd.isna(val):
                        val = None
                except (TypeError, ValueError):
                    pass

            # Sprint 17.17 — cap a 100% en columnas seleccionadas.
            if col_name in cols_cap_a_100 and isinstance(val, (int, float)) and val is not None:
                val = min(float(val), 1.0)

            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center", horizontal="center")
            cell.border    = borde

            if col_name in cols_pct and isinstance(val, (int, float)) and not pd.isna(val):
                cell.number_format = FMT_PCT
                cell.fill = PatternFill("solid", fgColor=_color_celda(val))
            elif col_name in cols_venta and isinstance(val, (int, float)) and not pd.isna(val):
                cell.number_format = FMT_VENTA
                cell.fill = fill_base
            elif isinstance(val, (int, float)) and not pd.isna(val) and not isinstance(val, bool):
                # Sprint 17.17 — numéricos genéricos: máx 4 decimales.
                cell.number_format = FMT_NUM
                cell.fill = fill_base
            else:
                cell.fill = fill_base

    # Anchos de columna automáticos (mín 12, máx 40)
    for col_idx, col_name in enumerate(df.columns, start=1):
        max_len = max(len(str(col_name)), 12)
        for row in df[col_name].astype(str).values:
            max_len = max(max_len, len(str(row)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    # Freeze + auto_filter
    ws.freeze_panes = ws.cell(row=fila_inicio_datos, column=1).coordinate
    ws.auto_filter.ref = ws.dimensions


def guardar_detalle(df: pd.DataFrame, mes: int, anio: int) -> Path:
    DIR_ALERTAS.mkdir(parents=True, exist_ok=True)
    ruta = DIR_ALERTAS / f"DETALLE_CUMPLIMIENTO_{mes:02d}_{anio}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Detalle"

    cols_pct = [
        "CIF_COB_%", "CIF_FREC_%", "CIF_HRS_%", "CIF_%",
        "NP_%", "PRECIOS_%", "SOS_%",
        "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
        "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
        "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%",
    ]
    cols_venta = ["VENTA_TOTAL"]
    _aplicar_formato_tabla(ws, df, cols_pct, cols_venta)

    wb.save(ruta)
    print(f"  ✅ Detalle guardado: {ruta}")
    return ruta


def guardar_resumen(df: pd.DataFrame, mes: int, anio: int) -> Path:
    DIR_ALERTAS.mkdir(parents=True, exist_ok=True)
    ruta = DIR_ALERTAS / f"RESUMEN_CUMPLIMIENTO_{mes:02d}_{anio}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"

    cols_pct   = [
        "CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
        "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
        "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
        "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%",
        "CUMPL_GLOBAL_%",
    ]
    cols_venta = ["VENTA_EQUIPO"]
    _aplicar_formato_tabla(ws, df, cols_pct, cols_venta)

    wb.save(ruta)
    print(f"  ✅ Resumen guardado: {ruta}")
    return ruta


# ─────────────────────────────────────────────────────────────────────────────
# E) ADJUNTOS POR SUPERVISOR  (4 hojas operativas + 3 hojas de productos nuevos)
#    Hojas:
#      1. Resumen Equipo
#      2. CIF Detalle PDVs
#      3. NP Detalle PDVs
#      4. Precios y SOS Detalle PDVs
#      5. ListSant     ← productos nuevos (target del segmento)
#      6. DoyPackBaby
#      7. CremasBaby
# ─────────────────────────────────────────────────────────────────────────────

_MENSAJE_VACIO      = "Sin incumplimientos para este periodo"
_MENSAJE_NO_APLICA  = "No aplica para este rol"


def construir_equipo_lider(df_detalle: pd.DataFrame, acr_lider: str) -> pd.DataFrame:
    """
    Para un LIDER, devuelve un DataFrame con (a) su propia fila y (b) una
    fila SINTÉTICA por cada supervisor asignado (filtrados a canal DIRECTO),
    donde los KPIs son el PROMEDIO del equipo del supervisor (gestores +
    supervisor) y N_PDVS es la SUMA. Usado por la hoja "Resumen Equipo"
    del adjunto Excel y por el cuerpo HTML del correo del líder.
    """
    propio = df_detalle[df_detalle["ACRONIMO"] == acr_lider]
    sups = df_detalle[df_detalle["ES_SUPERVISOR"] == True].copy()
    sups["_CIUDAD"]  = sups.get("CIUDAD_CUPO", "").astype(str).str.upper().str.strip()
    sups["_LIDER"]   = sups["_CIUDAD"].map(LIDER_POR_CIUDAD).fillna("")
    sups["_CANAL_U"] = sups.get("CANAL", "").astype(str).str.upper()
    sups = sups[sups["_CANAL_U"].str.contains("DIRECTO", na=False)]
    sups_acr = sups[sups["_LIDER"] == acr_lider]["ACRONIMO"].tolist()

    cols_pct_eq = ["CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
                   "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
                   "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
                   "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]
    filas_sup = []
    for acr_s in sups_acr:
        mask = (df_detalle["ACRONIMO_SUP"] == acr_s) | (df_detalle["ACRONIMO"] == acr_s)
        eq_sup = df_detalle[mask]
        if eq_sup.empty:
            continue
        base_row = df_detalle[df_detalle["ACRONIMO"] == acr_s].iloc[0].copy()
        for c in cols_pct_eq:
            if c in eq_sup.columns:
                v = pd.to_numeric(eq_sup[c], errors="coerce")
                if c in ("EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%"):
                    v = v.clip(upper=1.0)
                v = v.dropna()
                base_row[c] = float(v.mean()) if not v.empty else np.nan
        if "N_PDVS" in eq_sup.columns:
            base_row["N_PDVS"] = int(
                pd.to_numeric(eq_sup["N_PDVS"], errors="coerce").fillna(0).sum()
            )
        filas_sup.append(base_row)

    partes = [propio] + ([pd.DataFrame(filas_sup)] if filas_sup else [])
    return pd.concat([p for p in partes if not p.empty], ignore_index=True)


def _hoja_resumen_equipo(ws, df_detalle: pd.DataFrame, acr_sup: str) -> None:
    """
    Hoja 1: filas a mostrar al destinatario. Sprint 17.15:
      - SUPERVISOR: sus gestores a cargo + él mismo (si tiene KPIs propios).
      - GDD:        él mismo.
      - LIDER:      él mismo + supervisores que le asigne LIDER_POR_CIUDAD.
    """
    # Detectar rol del destinatario.
    propio = df_detalle[df_detalle["ACRONIMO"] == acr_sup]
    if propio.empty:
        eq = pd.DataFrame(columns=df_detalle.columns)
    else:
        info = propio.iloc[0]
        es_gdd   = bool(info.get("ES_GDD",   False))
        es_lider = bool(info.get("ES_LIDER", False))
        if es_gdd:
            eq = propio.copy()
        elif es_lider:
            eq = construir_equipo_lider(df_detalle, acr_sup)
        else:
            # SUPERVISOR (caso histórico): gestores a cargo + supervisor.
            gestores = df_detalle[
                (df_detalle["ES_SUPERVISOR"] == False)
                & (df_detalle["ACRONIMO_SUP"] == acr_sup)
            ]
            eq = pd.concat([gestores, propio], ignore_index=True)

    cols_pct_calc = ["CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
                     "EXHIB_PAG_%", "EXHIB_PAG_CAPTURA_%",
                     "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
                     "VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]
    # Columnas que efectivamente se muestran (sin VENTA_TOTAL ni MSL_%)
    cols_visibles_origen = [
        "NOMBRE", "ROL", "N_PDVS",
        "CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
        "EXHIB_PAG_CAPTURA_%", "EXHIB_PAG_%",
        "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",
        "VENTA_%", "IMPACTOS_%", "PROD_NUEVOS_%",
        "CUMPL_GLOBAL_%",
    ]
    # LIDER de ejecución no es responsable de canal D&P; omitir esas columnas
    # del adjunto y del cálculo de CUMPL_GLOBAL.
    if not propio.empty and bool(propio.iloc[0].get("ES_LIDER", False)):
        cols_dyp = {"VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"}
        cols_pct_calc       = [c for c in cols_pct_calc       if c not in cols_dyp]
        cols_visibles_origen = [c for c in cols_visibles_origen if c not in cols_dyp]

    # Mapeo a los nombres del entregable
    RENOMBRES = {
        "NOMBRE":               "NOMBRE",
        "ROL":                  "ROL",
        "N_PDVS":               "POS ASIGNADOS",
        "CIF_%":                "CUMPLIMIENTO CIF",
        "NP_%":                 "CUMPLIMIENTO CAPTURA AGOTADOS",
        "PRECIOS_%":            "CUMPLIMIENTO CAPTURA PRECIOS",
        "SOS_%":                "CUMPLIMIENTO CAPTURA SOS",
        "EXHIB_PAG_CAPTURA_%":  "CUMPLIMIENTO CAPTURA EXH. PAGADAS",
        "EXHIB_PAG_%":          "CUMPLIMIENTO EJECUCIÓN EXH. PAGADAS",
        "EXHIB_GRA_ALTO_%":     "Cump Exhibiciones Gratis Alto Impacto",
        "EXHIB_GRA_MEDIO_%":    "Cump Exhibiciones Gratis Medio Impacto",
        "VENTA_%":              "CUMPLIMIENTO CUOTA D&P",
        "IMPACTOS_%":           "CUMPLIMIENTO IMPACTOS GENERALES D&P",
        "PROD_NUEVOS_%":        "CUMPLIMIENTO IMPACTOS SKU NUEVOS",
        "CUMPL_GLOBAL_%":       "CUMPLIMIENTO GLOBAL",
    }
    cols_finales = [RENOMBRES[c] for c in cols_visibles_origen]

    if eq.empty:
        df_out = pd.DataFrame(columns=cols_finales)
        df_out.loc[0] = [_MENSAJE_NO_APLICA] + [None] * (len(cols_finales) - 1)
    else:
        # Sprint 17.17 — capar Exh Gratis ALTO/MEDIO a 100% antes de promediar
        # (consistencia con el cuerpo HTML del correo). Sobrecumplimientos
        # extremos no inflan el CUMPL_GLOBAL.
        eq = eq.copy()
        for _c in ("EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%"):
            if _c in eq.columns:
                eq[_c] = pd.to_numeric(eq[_c], errors="coerce").clip(upper=1.0)
        # Calcular CUMPL_GLOBAL_% por gestor (promedio simple de los KPIs)
        eq["CUMPL_GLOBAL_%"] = eq[cols_pct_calc].mean(axis=1, skipna=True)
        for c in cols_visibles_origen:
            if c not in eq.columns:
                eq[c] = np.nan
        df_out = (
            eq[cols_visibles_origen]
              .sort_values("NOMBRE")
              .reset_index(drop=True)
              .rename(columns=RENOMBRES)
        )

    cols_pct_render_origen = ["CIF_%", "NP_%", "PRECIOS_%", "SOS_%",
                              "EXHIB_PAG_CAPTURA_%", "EXHIB_PAG_%",          # Sprint 17.17
                              "EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%",       # Sprint 17.17
                              "VENTA_%", "IMPACTOS_%", "PROD_NUEVOS_%",
                              "CUMPL_GLOBAL_%"]
    cols_pct_render = [RENOMBRES[c] for c in cols_pct_render_origen
                       if RENOMBRES[c] in df_out.columns]
    _aplicar_formato_tabla(
        ws, df_out,
        cols_pct=cols_pct_render,
        cols_venta=[],
    )


def _hoja_cif_detalle(ws, df_cif_pdv: pd.DataFrame, acr_sup: str) -> None:
    """Hoja 2: PDVs del equipo con incumplimiento en algún componente CIF."""
    if "ACRONIMO_SUP" not in df_cif_pdv.columns:
        pdvs = pd.DataFrame()
    else:
        pdvs = df_cif_pdv[
            (df_cif_pdv["ROL"] == ROL_GESTOR)
            & (df_cif_pdv["ACRONIMO_SUP"] == acr_sup)
        ].copy()

    # Sprint 10: V3 cambió los nombres COB/FREC/HRS → COBERTURA/INTENSIDAD/FRECUENCIA.
    # Mantenemos compat con los nombres legacy para no tocar el resto del flujo.
    if not pdvs.empty:
        rename_v3 = {}
        if "COBERTURA"  in pdvs.columns and "COB"  not in pdvs.columns: rename_v3["COBERTURA"]  = "COB"
        if "INTENSIDAD" in pdvs.columns and "HRS"  not in pdvs.columns: rename_v3["INTENSIDAD"] = "HRS"
        if "FRECUENCIA" in pdvs.columns and "FREC" not in pdvs.columns: rename_v3["FRECUENCIA"] = "FREC"
        if rename_v3:
            pdvs = pdvs.rename(columns=rename_v3)

    if not pdvs.empty:
        # Incumplimiento = COB == 0 OR FREC < 1 OR HRS < 1
        incumple = (
            pdvs["COB"].fillna(0).eq(0)
            | pdvs["FREC"].fillna(1).lt(1)
            | pdvs["HRS"].fillna(1).lt(1)
        )
        pdvs = pdvs[incumple].copy()

    cols_origen = ["NOMBRE_GESTOR", "NOMBRE_PDV", "VENTAS_PROMEDIO_MES",
                   "CANTIDAD_VISITAS", "VISITAS_REAL",
                   "COB", "FREC", "HRS",
                   "HORAS_MES_PLANEADO", "SUMA_TIEMPO_SERVICIO_HORAS_REAL"]

    RENOMBRES = {
        "NOMBRE_GESTOR":                    "NOMBRE COLABORADOR",
        "NOMBRE_PDV":                       "NOMBRE PUNTO DE VENTA",
        "VENTAS_PROMEDIO_MES":              "VENTAS PROMEDIO PDV",
        "CANTIDAD_VISITAS":                 "FREC PLAN",
        "VISITAS_REAL":                     "FREC REAL",
        "COB":                              "CUMPL COBERTURA",
        "FREC":                             "CUMPL FRECUENCIA",
        "HRS":                              "CUMPL INTENSIDAD",
        "HORAS_MES_PLANEADO":               "INTENSIDAD PLAN",
        "SUMA_TIEMPO_SERVICIO_HORAS_REAL":  "INTENSIDAD REAL",
    }
    cols_finales = [RENOMBRES[c] for c in cols_origen]

    if pdvs.empty:
        df_out = pd.DataFrame(columns=cols_finales)
        df_out.loc[0] = [_MENSAJE_VACIO] + [None] * (len(cols_finales) - 1)
    else:
        pdvs = pdvs.rename(columns={"NOMBRE": "NOMBRE_GESTOR"})
        df_out = (
            pdvs[cols_origen]
                .sort_values(["NOMBRE_GESTOR", "VENTAS_PROMEDIO_MES"],
                             ascending=[True, False])
                .reset_index(drop=True)
                .rename(columns=RENOMBRES)
        )

    _aplicar_formato_tabla(
        ws, df_out,
        cols_pct=[RENOMBRES["COB"], RENOMBRES["FREC"], RENOMBRES["HRS"]],
        cols_venta=[RENOMBRES["VENTAS_PROMEDIO_MES"]],
    )


def _hoja_np_detalle(ws, df_np: pd.DataFrame, df_cif_pdv: pd.DataFrame, acr_sup: str) -> None:
    """Hoja 3: PDVs con %_CUMPLIMIENTO_MES < 1 en NP, del equipo del supervisor.

    Nota: VENTAS_PROMEDIO_MES se eliminó del entregable; sólo se conserva
    internamente como criterio de orden secundario.
    """
    cols_origen = ["NOMBRE_GESTOR", "NOMBRE_PDV",
                   "PLANEADO_MES", "CAPTURA_MES", "%_CUMPLIMIENTO_MES"]
    RENOMBRES = {
        "NOMBRE_GESTOR":      "NOMBRE COLABORADOR",
        "NOMBRE_PDV":         "NOMBRE PUNTO DE VENTA",
        "PLANEADO_MES":       "POS PLANEADOS MEDICION",
        "CAPTURA_MES":        "POS MEDIDOS REAL",
        "%_CUMPLIMIENTO_MES": "CUMP CAPTURA",
    }
    cols_finales = [RENOMBRES[c] for c in cols_origen]

    def _vacio(mensaje):
        df_out = pd.DataFrame(columns=cols_finales)
        df_out.loc[0] = [mensaje] + [None] * (len(cols_finales) - 1)
        _aplicar_formato_tabla(
            ws, df_out,
            cols_pct=[RENOMBRES["%_CUMPLIMIENTO_MES"]],
            cols_venta=[],
        )

    if df_np.empty:
        _vacio(_MENSAJE_VACIO)
        return

    df = df_np.copy()
    df["NOMBRE"]             = _normalizar_nombre(df["NOMBRE"])
    df["%_CUMPLIMIENTO_MES"] = pd.to_numeric(df["%_CUMPLIMIENTO_MES"], errors="coerce")

    # Sprint 13.4: filtrar PDVs con cumplimiento < UMBRAL_OK_NP (default 1.0)
    umb_np = umbral_de("NP_%")
    if "ACRONIMO_SUP" not in df.columns:
        df = df.head(0)  # vacío
    else:
        df = df[
            (df["ACRONIMO_SUP"] == acr_sup)
            & (pd.to_numeric(df["PLANEADO_MES"], errors="coerce").fillna(0) >= 1)
            & (df["%_CUMPLIMIENTO_MES"].fillna(0) < umb_np)
        ].copy()

    if df.empty:
        _vacio(_MENSAJE_VACIO)
        return

    # Cruce con CIF para usar VENTAS_PROMEDIO_MES como criterio de orden
    ventas_pdv = (
        df_cif_pdv[["ID_PDV_INVOLVES", "VENTAS_PROMEDIO_MES"]]
        .drop_duplicates(subset=["ID_PDV_INVOLVES"])
    )
    df = df.merge(ventas_pdv, on="ID_PDV_INVOLVES", how="left")
    df = df.rename(columns={"NOMBRE": "NOMBRE_GESTOR"})

    df_out = (
        df.sort_values(["%_CUMPLIMIENTO_MES", "VENTAS_PROMEDIO_MES"],
                       ascending=[True, False])
          [cols_origen]
          .reset_index(drop=True)
          .rename(columns=RENOMBRES)
    )

    _aplicar_formato_tabla(
        ws, df_out,
        cols_pct=[RENOMBRES["%_CUMPLIMIENTO_MES"]],
        cols_venta=[],
    )


def _hoja_modulo_captura(
    ws,
    df: pd.DataFrame,
    acr_sup: str,
    umbral_ok: float = 1.0,
) -> None:
    """Helper genérico para hojas Precios/SOS — Detalle (Sprint 13.5).

    Filtra PDVs donde el equipo del supervisor tiene incumplimiento:
        CAPTURA_PLANEADA >= 1  Y  CAPTURA_EJECUTADA/CAPTURA_PLANEADA < umbral_ok
    Sprint 13.4: con umbral=1.0 entran TODOS los PDVs que no llegaron al 100%.
    """
    cols_origen = ["NOMBRE_GESTOR", "NOMBRE_PDV",
                   "CAPTURA_PLANEADA", "CAPTURA_EJECUTADA",
                   "CATEGORIAS_FALTANTES"]
    RENOMBRES = {
        "NOMBRE_GESTOR":        "NOMBRE COLABORADOR",
        "NOMBRE_PDV":           "NOMBRE PUNTO DE VENTA",
        "CAPTURA_PLANEADA":     "POS PLANEADOS MEDICION",
        "CAPTURA_EJECUTADA":    "POS MEDIDOS REAL",
        "CATEGORIAS_FALTANTES": "DETALLE CATEGORIAS FALTANTES",
    }
    cols_finales = [RENOMBRES[c] for c in cols_origen]

    if df is None or df.empty or "ACRONIMO_SUP" not in df.columns:
        df_out = pd.DataFrame(columns=cols_finales)
        df_out.loc[0] = [_MENSAJE_VACIO] + [None] * (len(cols_finales) - 1)
        _aplicar_formato_tabla(ws, df_out, cols_pct=[], cols_venta=[])
        return

    d = df.copy()
    d["NOMBRE"] = _normalizar_nombre(d["NOMBRE"])
    d["CAPTURA_PLANEADA"]  = pd.to_numeric(d["CAPTURA_PLANEADA"],  errors="coerce").fillna(0)
    d["CAPTURA_EJECUTADA"] = pd.to_numeric(d["CAPTURA_EJECUTADA"], errors="coerce").fillna(0)

    # Cumplimiento por fila = ejecutada/planeada
    cumpl = np.where(
        d["CAPTURA_PLANEADA"] > 0,
        d["CAPTURA_EJECUTADA"] / d["CAPTURA_PLANEADA"],
        np.nan,
    )
    d["_cumpl"] = cumpl

    d = d[
        (d["ACRONIMO_SUP"] == acr_sup)
        & (d["CAPTURA_PLANEADA"] >= 1)
        & (d["_cumpl"] < umbral_ok)   # incumplimiento según umbral por KPI
    ].copy()

    if d.empty:
        df_out = pd.DataFrame(columns=cols_finales)
        df_out.loc[0] = [_MENSAJE_VACIO] + [None] * (len(cols_finales) - 1)
    else:
        d = d.rename(columns={"NOMBRE": "NOMBRE_GESTOR"})
        if "CATEGORIAS_FALTANTES" not in d.columns:
            d["CATEGORIAS_FALTANTES"] = ""
        if "VENTAS_PROMEDIO_MES" not in d.columns:
            d["VENTAS_PROMEDIO_MES"] = np.nan
        df_out = (
            d.sort_values(["_cumpl", "VENTAS_PROMEDIO_MES"],
                           ascending=[True, False])
             [cols_origen]
             .reset_index(drop=True)
             .rename(columns=RENOMBRES)
        )

    _aplicar_formato_tabla(ws, df_out, cols_pct=[], cols_venta=[])


def _hoja_impactos_detalle(
    ws,
    df_detalle: pd.DataFrame,
    acr_sup: str,
    mes: int,
    anio: int,
) -> None:
    """Hoja "Impactos — Detalle" (Sprint 13.6).

    Lee Consolidado_Impactos.csv y muestra el detalle por gestor del equipo
    del supervisor con su cumplimiento de impactos (Clientes Impactados /
    Total Clientes Impactar). Solo aparece si el supervisor tiene gestores
    con datos D&P; sino, mensaje "No aplica para este rol".
    """
    cols_origen = ["NOMBRE_GESTOR", "CIUDAD",
                   "CLIENTES_IMPACTAR", "CLIENTES_IMPACTADOS", "IMPACTOS_%"]
    RENOMBRES = {
        "NOMBRE_GESTOR":       "NOMBRE COLABORADOR",
        "CIUDAD":              "CIUDAD",
        "CLIENTES_IMPACTAR":   "CLIENTES A IMPACTAR",
        "CLIENTES_IMPACTADOS": "CLIENTES IMPACTADOS",
        "IMPACTOS_%":          "CUMP IMPACTOS",
    }
    cols_finales = [RENOMBRES[c] for c in cols_origen]

    def _vacio(mensaje):
        df_out = pd.DataFrame(columns=cols_finales)
        df_out.loc[0] = [mensaje] + [None] * (len(cols_finales) - 1)
        _aplicar_formato_tabla(
            ws, df_out,
            cols_pct=[RENOMBRES["IMPACTOS_%"]],
            cols_venta=[],
        )

    if not paths.DYP_OUT_IMPACTOS.is_file():
        _vacio(_MENSAJE_NO_APLICA)
        return

    try:
        df_imp = pd.read_csv(
            str(paths.DYP_OUT_IMPACTOS),
            sep="|", decimal=",", encoding="utf-8",
            dtype={"MES": str, "AÑO": str},
            low_memory=False,
        )
    except Exception:
        _vacio(_MENSAJE_NO_APLICA)
        return

    # Filtrar al periodo activo
    from cumplimiento_dyp import MESES_INT_A_STR, parsear_nombre_asesor
    mes_str  = MESES_INT_A_STR.get(int(mes), "")
    anio_str = str(int(anio))
    df_imp["MES"] = df_imp["MES"].astype(str).str.strip().str.capitalize()
    df_imp["AÑO"] = df_imp["AÑO"].astype(str).str.strip()
    df_imp = df_imp[(df_imp["MES"] == mes_str) & (df_imp["AÑO"] == anio_str)].copy()
    if df_imp.empty:
        _vacio(_MENSAJE_NO_APLICA)
        return

    df_imp["NOMBRE_GESTOR"]       = df_imp["Asesor"].apply(parsear_nombre_asesor)
    df_imp["CIUDAD"]              = df_imp["Ciudad"].astype(str).str.strip()
    df_imp["CLIENTES_IMPACTAR"]   = pd.to_numeric(df_imp["Total Clientes Impactar"], errors="coerce").fillna(0)
    df_imp["CLIENTES_IMPACTADOS"] = pd.to_numeric(df_imp["Clientes Impactados"],     errors="coerce").fillna(0)

    # Agregar por (NOMBRE_GESTOR, CIUDAD)
    agg = (
        df_imp.groupby(["NOMBRE_GESTOR", "CIUDAD"], as_index=False)
              .agg(CLIENTES_IMPACTAR=("CLIENTES_IMPACTAR", "sum"),
                   CLIENTES_IMPACTADOS=("CLIENTES_IMPACTADOS", "sum"))
    )
    agg["IMPACTOS_%"] = np.where(
        agg["CLIENTES_IMPACTAR"] > 0,
        agg["CLIENTES_IMPACTADOS"] / agg["CLIENTES_IMPACTAR"],
        np.nan,
    )

    # Filtrar al equipo del supervisor (ACRONIMO_SUP=acr_sup): cruzar
    # NOMBRE_GESTOR con df_detalle para conseguir ACRONIMO_SUP.
    # Sprint 17.16 — armonización por subset de palabras: el CSV de impactos
    # trae nombres truncados ("CLEOFELINA TORRADO") que no matchean con el
    # universo ("CLEOFELINA TORRADO ALVAREZ") por igualdad literal.
    equipo = df_detalle[
        (df_detalle["ES_SUPERVISOR"] == False)
        & (df_detalle["ACRONIMO_SUP"] == acr_sup)
    ][["NOMBRE"]].copy()
    equipo["NOMBRE"] = equipo["NOMBRE"].astype(str).str.strip().str.upper()
    nombres_equipo = set(equipo["NOMBRE"])

    def _match_subset(nombre_csv: str) -> bool:
        pal_csv = _palabras(str(nombre_csv))
        if not pal_csv:
            return False
        for n_eq in nombres_equipo:
            pal_eq = _palabras(n_eq)
            if pal_csv.issubset(pal_eq) or pal_eq.issubset(pal_csv):
                return True
        return False

    agg = agg[agg["NOMBRE_GESTOR"].apply(_match_subset)].copy()

    # Filtrar solo incumplimientos (cumplimiento < UMBRAL_OK_IMPACTOS)
    umb_imp = umbral_de("IMPACTOS_%")
    agg = agg[agg["IMPACTOS_%"].fillna(0) < umb_imp].copy()

    if agg.empty:
        _vacio(_MENSAJE_VACIO)
        return

    df_out = (
        agg.sort_values(["IMPACTOS_%", "CLIENTES_IMPACTAR"],
                         ascending=[True, False])
            [cols_origen]
            .reset_index(drop=True)
            .rename(columns=RENOMBRES)
    )

    _aplicar_formato_tabla(
        ws, df_out,
        cols_pct=[RENOMBRES["IMPACTOS_%"]],
        cols_venta=[],
    )


def _hoja_precios_detalle(ws, df_pr: pd.DataFrame, acr_sup: str) -> None:
    """Hoja Precios — Detalle (Sprint 13.5)."""
    _hoja_modulo_captura(ws, df_pr, acr_sup, umbral_ok=umbral_de("PRECIOS_%"))


def _hoja_sos_detalle(ws, df_sos: pd.DataFrame, acr_sup: str) -> None:
    """Hoja SOS — Detalle (Sprint 13.5)."""
    _hoja_modulo_captura(ws, df_sos, acr_sup, umbral_ok=umbral_de("SOS_%"))


# Versión legacy mantenida por compatibilidad con código antiguo si alguien
# todavía la llama. Sprint 13.5 oficializó las funciones separadas.
def _hoja_precios_sos_detalle(
    ws, df_pr: pd.DataFrame, df_sos: pd.DataFrame, acr_sup: str,
) -> None:
    """[Deprecated Sprint 13.5] Use _hoja_precios_detalle + _hoja_sos_detalle."""
    _hoja_modulo_captura(ws, df_pr, acr_sup, umbral_ok=umbral_de("PRECIOS_%"))


# ─────────────────────────────────────────────────────────────────────────────
# E.5)  HOJAS DE PRODUCTOS NUEVOS POR SUPERVISOR
#       Una hoja por segmento (ListSant, DoyPackBaby, CremasBaby) — MSL queda
#       fuera por petición del cliente. La estructura es la misma que la de
#       `etl_impactos_segmentos.construir_bd`, restringiendo el universo de
#       PDVs a los del supervisor (intersección con los PDVs target del
#       segmento).
# ─────────────────────────────────────────────────────────────────────────────

# Segmentos a generar (excluye MustStock = MSL, según indicación del cliente)
# Sprint 14.5: 4 segmentos V2. Importamos la lista canónica del ETL para
# que el único punto de verdad sea etl_impactos_segmentos.SEGMENTOS_PRODNUEVOS.
try:
    from etl_impactos_segmentos import SEGMENTOS_PRODNUEVOS as _SEGMENTOS_NUEVOS
except Exception:
    _SEGMENTOS_NUEVOS = ["Listerine_Sandia", "DoyPack_J&J_Baby",
                          "Liserine_Kids", "Visine"]


def _precargar_segmentos_nuevos(
    mes: int,
    anio: int,
    idx_bc: dict,
    nombres_sups_bc: set,
) -> dict:
    """
    Carga UNA SOLA VEZ los insumos para las hojas de productos nuevos:
      • rutero (PDVs canónicos)
      • listas de referencia (segmentos → EANs + PDVs target)
      • ventas consolidadas (todas las del CSV)
      • universo de PDVs por supervisor (vendidos por su equipo en el periodo)

    Devuelve dict con claves: rutero, listas, ventas, mes_actual, pdvs_por_sup.
    Si falta cualquier insumo, devuelve dict vacío (las hojas quedarán con
    el mensaje "No aplica para este rol").
    """
    out = {
        "rutero": None, "listas": None, "ventas": None,
        "mes_actual": "", "pdvs_por_sup": {},
    }

    try:
        import etl_impactos_segmentos as eis
    except Exception as e:
        print(f"  ⚠️  No pude importar etl_impactos_segmentos: {e}")
        return out

    if not (paths.DYP_RUTERO_FILE.is_file()
            and paths.DYP_LISTAS_FILE.is_file()
            and paths.DYP_OUT_VENTAS.is_file()):
        print("  ⚠️  Faltan insumos D&P (rutero / listas / consolidado de ventas)")
        print("       Las hojas de productos nuevos quedarán como 'No aplica'.")
        return out

    try:
        rutero = eis.cargar_rutero(paths.DYP_RUTERO_FILE)
        listas = eis.cargar_listas(paths.DYP_LISTAS_FILE)
        # V2 (Sprint 14.3): cargar_ventas_desde_csv devuelve (ventas, desc_map)
        ventas_result = eis.cargar_ventas_desde_csv(paths.DYP_OUT_VENTAS)
        ventas = ventas_result[0] if isinstance(ventas_result, tuple) else ventas_result
    except Exception as e:
        print(f"  ⚠️  Error cargando insumos de segmentos: {e}")
        return out

    if ventas.empty:
        print("  ⚠️  Consolidado de ventas vacío — hojas de productos nuevos no se llenarán.")
        return out

    # Período más reciente — el mismo criterio que usa etl_impactos_segmentos
    ventas["_sort"] = ventas.apply(
        lambda r: eis.periodo_sort_key(f"{r['mes']}-{r['año']}"), axis=1
    )
    fila_max   = ventas.loc[ventas["_sort"].idxmax()]
    mes_actual = f"{fila_max['mes']}-{fila_max['año']}"
    ventas.drop(columns=["_sort"], inplace=True)

    # Universo de PDVs del supervisor — PDVs donde el equipo vendió en el mes
    # activo (según el campo Supervisor del CSV de ventas). Resolvemos cada
    # nombre crudo a su ACRONIMO de Base cupos.
    df_periodo = cumplimiento_dyp._cargar_ventas_periodo(mes, anio)
    pdvs_por_sup: dict[str, set] = {}
    if not df_periodo.empty:
        df_v = df_periodo[df_periodo["cant_total"] > 0].copy()
        df_v["acr_sup"] = df_v["supervisor"].apply(
            lambda s: _resolver_acr_supervisor(s, nombres_sups_bc, idx_bc)
        )
        pdvs_por_sup = (
            df_v[df_v["acr_sup"] != ""]
              .groupby("acr_sup")["cod_cliente"]
              .apply(lambda s: set(s.dropna().astype(str).str.strip()))
              .to_dict()
        )

    out.update({
        "rutero": rutero,
        "listas": listas,
        "ventas": ventas,
        "mes_actual": mes_actual,
        "pdvs_por_sup": pdvs_por_sup,
    })
    return out


def _hoja_segmento_nuevo(
    ws,
    nombre_segmento: str,
    seg_data: dict,
    acr_sup: str,
) -> None:
    """
    Hoja por segmento de producto nuevo (ListSant / DoyPackBaby / CremasBaby).

    Universo: intersección de los PDVs target del segmento con los PDVs
    donde el equipo del supervisor vendió en el periodo. Si esa intersección
    es vacía → "No aplica para este rol".
    """
    # Guard: ¿se pudieron precargar los insumos?
    if not seg_data or seg_data.get("rutero") is None:
        ws.cell(row=1, column=1, value=_MENSAJE_NO_APLICA)
        ws.column_dimensions["A"].width = 50
        return

    rutero      = seg_data["rutero"]
    listas      = seg_data["listas"]
    ventas      = seg_data["ventas"]
    mes_actual  = seg_data["mes_actual"]
    pdvs_sup    = seg_data["pdvs_por_sup"].get(acr_sup, set())

    pdvs_target_seg = listas.get(nombre_segmento, {}).get("pdvs") or set()
    eans_seg        = listas.get(nombre_segmento, {}).get("eans") or set()

    # Universo del supervisor para este segmento
    pdvs_target_para_sup = pdvs_target_seg & pdvs_sup if pdvs_sup else set()

    if not pdvs_target_para_sup:
        ws.cell(row=1, column=1, value=_MENSAJE_NO_APLICA)
        ws.column_dimensions["A"].width = 50
        return

    # Reusamos construir_bd del ETL — misma estructura/columnas que la salida
    # global de impacto_segmentos.xlsx.
    import etl_impactos_segmentos as eis
    df_seg = eis.construir_bd(
        rutero        = rutero,
        ventas        = ventas,
        eans_segmento = eans_seg,
        pdvs_target   = pdvs_target_para_sup,
        mes_actual    = mes_actual,
    )

    if df_seg.empty:
        ws.cell(row=1, column=1, value=_MENSAJE_NO_APLICA)
        ws.column_dimensions["A"].width = 50
        return

    # Si TODOS los PDVs target del supervisor tuvieron impacto → no hay nada
    # que reportar como incumplimiento.
    if "tiene_impacto" in df_seg.columns and (df_seg["tiene_impacto"] == "Sí").all():
        ws.cell(row=1, column=1, value=_MENSAJE_VACIO)
        ws.column_dimensions["A"].width = 50
        return

    # Volcado a la hoja con el mismo formato visual que el ETL global
    # (encabezado azul, filas verde/rojo según impacto).
    for col_idx, col_name in enumerate(df_seg.columns, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)
    for row_idx, row in enumerate(df_seg.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row, start=1):
            try:
                if pd.isna(val):
                    val = None
            except (TypeError, ValueError):
                pass
            ws.cell(row=row_idx, column=col_idx, value=val)
    eis.aplicar_formato(ws, df_seg)


# ─────────────────────────────────────────────────────────────────────────────
# HOJAS ADICIONALES SPRINT 17.10 — Exhibiciones (no capturadas / fuera de regla)
# ─────────────────────────────────────────────────────────────────────────────

_CACHE_EXH_PLANNING: dict = {}
_CACHE_BASE_CUPOS: dict = {}


def _hoja_exh_pagadas_no_capturadas(ws, acr_sup: str, mes: int, anio: int,
                                     df_cif_pdv: pd.DataFrame | None = None) -> None:
    """
    Lista los PDVs del PLANNING de exhibiciones pagadas asignados al equipo
    del supervisor que NO respondieron la encuesta para el periodo.
    Granularidad: (PDV, gestor, marca, tipo). Filtrado al equipo del supervisor.

    Los archivos de planning se localizan y leen directamente desde
    SharePoint (Graph API), replicando en la nube la convención de nombres
    de `periodo_resolver.exh_planning_master` / `exh_base_planning`.

    Esta hoja se genera una vez POR SUPERVISOR (llamada desde
    `generar_adjuntos_por_supervisor`), pero el planning del periodo es el
    mismo para todos — se cachea en `_CACHE_EXH_PLANNING` para no volver a
    descargar los mismos 2 archivos de SharePoint una vez por cada
    supervisor (con 30+ supervisores, eso son 60+ descargas evitables).
    """
    cache_key = (mes, anio)
    if cache_key in _CACHE_EXH_PLANNING:
        cached = _CACHE_EXH_PLANNING[cache_key]
        if isinstance(cached, Exception):
            ws.cell(row=1, column=1, value=f"No disponible: {cached}")
            ws.column_dimensions["A"].width = 80
            return
        df_plan, df_enc = cached[0].copy(), cached[1].copy()
    else:
        import periodo_resolver as pr_mod
        spec = pr_mod.resolver(mes, anio)
        try:
            carpeta_exhib = _resolver_exhib_data_dir_cloud()
            ruta_plan_cloud = _buscar_archivo_cloud(
                carpeta_exhib,
                rf"^PLANNING DE {re.escape(spec.mes_str_upper)} {spec.anio}\.xlsx$",
                contexto=f"Exh/Planning maestro {spec.etiqueta}",
            )
            ruta_enc_cloud = _buscar_archivo_cloud(
                carpeta_exhib,
                rf"^Base Exhibiciones Planning {re.escape(spec.mes_str)} {spec.anio}\.xlsx$",
                contexto=f"Exh/Base Planning {spec.etiqueta}",
            )
        except (FileNotFoundError, RuntimeError) as e:
            _CACHE_EXH_PLANNING[cache_key] = e
            ws.cell(row=1, column=1, value=f"No disponible: {e}")
            ws.column_dimensions["A"].width = 80
            return

        df_plan = _leer_excel_cloud(ruta_plan_cloud, "Exh Planning Maestro")
        df_enc  = _leer_excel_cloud(ruta_enc_cloud,  "Exh Base Planning")
        _CACHE_EXH_PLANNING[cache_key] = (df_plan, df_enc)
        df_plan, df_enc = df_plan.copy(), df_enc.copy()

    col_pdv_p = next((c for c in df_plan.columns if "punto de venta" in str(c).lower()), None)
    col_emp_p = next((c for c in df_plan.columns if "empleado"       in str(c).lower()), None)
    if not (col_pdv_p and col_emp_p):
        ws.cell(row=1, column=1, value="Planning sin columnas reconocidas.")
        return

    df_plan["PDV_N"]      = df_plan[col_pdv_p].astype(str).str.strip().str.upper()
    df_plan["EMPLEADO_N"] = df_plan[col_emp_p].astype(str).str.strip().str.upper()
    col_pdv_e = "PDV" if "PDV" in df_enc.columns else next(
        (c for c in df_enc.columns if str(c).strip().upper() == "PDV"), None)
    pdvs_capturados = (
        set(df_enc[col_pdv_e].astype(str).str.strip().str.upper().unique())
        if col_pdv_e else set()
    )

    # Mapa EMPLEADO_N → ACRONIMO_SUP via Base cupos.
    # Se usa bcm.cargar() (ya migrado a SharePoint) en vez de leer el Excel
    # local directamente. Se cachea a nivel de módulo porque esta hoja se
    # genera una vez por supervisor y bcm.cargar() es una descarga + proceso
    # completo (incluye la auditoría) — sin caché se repetía 30+ veces.
    if "bc" not in _CACHE_BASE_CUPOS:
        _CACHE_BASE_CUPOS["bc"] = bcm.cargar()
    bc = _CACHE_BASE_CUPOS["bc"].copy()
    bc["NOMBRE_N"] = (
        bc["NOMBRE"].astype(str).str.strip().str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )
    # Si Base cupos no tiene ACRONIMO_SUP, derivamos del df_detalle global
    # (queda como tarea simple: filtramos por empleados cuyo nombre aparezca
    # en el equipo del supervisor según el resumen del cumplimiento).
    bc["CANAL"] = bc["CANAL"].astype(str).str.strip().str.upper()

    # Empleados del equipo desde df_cif_pdv (en memoria, tiene ACRONIMO_SUP).
    if df_cif_pdv is not None and not df_cif_pdv.empty and "ACRONIMO_SUP" in df_cif_pdv.columns:
        empleados_eq = set(
            df_cif_pdv.loc[df_cif_pdv["ACRONIMO_SUP"] == acr_sup, "NOMBRE"]
                .astype(str).str.strip().str.upper()
                .str.replace(r"\s+", " ", regex=True).unique()
        )
    else:
        empleados_eq = set()

    if not empleados_eq:
        ws.cell(row=1, column=1, value=_MENSAJE_NO_APLICA)
        ws.column_dimensions["A"].width = 50
        return

    # Sprint 17.16 — armonización por subset de palabras: el PLANNING puede
    # traer nombres truncados que no matchean por igualdad literal.
    pal_eq_set = [(_palabras(n), n) for n in empleados_eq]

    def _match_emp(nombre_csv: str) -> bool:
        pal_csv = _palabras(str(nombre_csv))
        if not pal_csv:
            return False
        for pal_eq, _n in pal_eq_set:
            if pal_csv.issubset(pal_eq) or pal_eq.issubset(pal_csv):
                return True
        return False

    plan_eq = df_plan[df_plan["EMPLEADO_N"].apply(_match_emp)].copy()
    no_cap = plan_eq[~plan_eq["PDV_N"].isin(pdvs_capturados)].copy()
    if no_cap.empty:
        ws.cell(row=1, column=1, value="Sin PDVs pendientes de captura — todos respondieron.")
        ws.column_dimensions["A"].width = 70
        return

    cols_out = [
        col_emp_p, col_pdv_p,
        "*Tipo - Pagadas" if "*Tipo - Pagadas" in no_cap.columns else None,
        "*MARCA - Pagadas" if "*MARCA - Pagadas" in no_cap.columns else None,
        "*Fecha de expiración" if "*Fecha de expiración" in no_cap.columns else None,
    ]
    cols_out = [c for c in cols_out if c]
    detalle = no_cap[cols_out].drop_duplicates().reset_index(drop=True)
    detalle.columns = ["EMPLEADO", "PDV", "TIPO", "MARCA", "FECHA EXP"][:len(cols_out)]
    detalle = detalle.sort_values(["EMPLEADO", "PDV"]).reset_index(drop=True)

    for c_idx, c_name in enumerate(detalle.columns, start=1):
        ws.cell(row=1, column=c_idx, value=c_name)
    for r_idx, row in enumerate(detalle.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            if pd.isna(val): val = None
            ws.cell(row=r_idx, column=c_idx, value=val)


_CACHE_EXH_GRATIS_FUERA_REGLA: dict = {}


def _hoja_exh_gratis_fuera_regla(ws, acr_sup: str, mes: int, anio: int,
                                  df_cif_pdv: pd.DataFrame | None = None) -> None:
    """
    Lista las exhibiciones gratis del equipo que NO cumplieron la regla de
    "≥2 semanas distintas" cuando frecuencia de referencia > 1. Persistido
    por el ETL en SharePoint, en
    RUTA_CARPETA_SALIDAS_EXHIB/Exh_Gratis_Fuera_de_Regla.xlsx.

    Se cachea a nivel de módulo (`_CACHE_EXH_GRATIS_FUERA_REGLA`) porque
    esta hoja se genera una vez por supervisor pero el archivo fuente es el
    mismo para todos — sin caché se re-descargaba desde SharePoint una vez
    por cada supervisor.
    """
    if "df" not in _CACHE_EXH_GRATIS_FUERA_REGLA:
        ruta_cloud = f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/Exh_Gratis_Fuera_de_Regla.xlsx"
        _CACHE_EXH_GRATIS_FUERA_REGLA["df"] = _leer(ruta_cloud, "Exh Gratis Fuera de Regla")
    df = _CACHE_EXH_GRATIS_FUERA_REGLA["df"]
    if df.empty:
        ws.cell(row=1, column=1, value="No hay archivo de fuera de regla (no se generó este periodo).")
        ws.column_dimensions["A"].width = 70
        return
    col_mes  = "Mes del año" if "Mes del año" in df.columns else "Mes"
    col_anio = "Año" if "Año" in df.columns else next((c for c in df.columns if "año" in str(c).lower()), None)
    if col_mes in df.columns and col_anio:
        df = df[(pd.to_numeric(df[col_mes], errors="coerce") == mes) &
                (pd.to_numeric(df[col_anio], errors="coerce") == anio)]
    if df.empty:
        ws.cell(row=1, column=1, value="Sin exhibiciones fuera de regla en este periodo.")
        ws.column_dimensions["A"].width = 60
        return

    # Empleados del equipo desde df_cif_pdv (memoria con ACRONIMO_SUP).
    if df_cif_pdv is not None and not df_cif_pdv.empty and "ACRONIMO_SUP" in df_cif_pdv.columns:
        empleados_eq = set(
            df_cif_pdv.loc[df_cif_pdv["ACRONIMO_SUP"] == acr_sup, "NOMBRE"]
                .astype(str).str.strip().str.upper()
                .str.replace(r"\s+", " ", regex=True).unique()
        )
    else:
        empleados_eq = set()

    if "Empleado" in df.columns:
        # Sprint 17.16 — armonización por subset de palabras.
        pal_eq_set = [(_palabras(n), n) for n in empleados_eq]

        def _match_emp(nombre_csv: str) -> bool:
            pal_csv = _palabras(str(nombre_csv))
            if not pal_csv:
                return False
            for pal_eq, _n in pal_eq_set:
                if pal_csv.issubset(pal_eq) or pal_eq.issubset(pal_csv):
                    return True
            return False

        df = df[df["Empleado"].apply(_match_emp)]
    if df.empty:
        ws.cell(row=1, column=1, value="Sin exhibiciones fuera de regla para el equipo.")
        ws.column_dimensions["A"].width = 60
        return

    cols_show = [
        "Empleado", "ID del PDV", "Categoría", "Marca",
        "Tipo Exhibición", "Cantidad", "_semanas_distintas", "_frec_ref",
    ]
    cols_show = [c for c in cols_show if c in df.columns]
    out = df[cols_show].rename(columns={
        "_semanas_distintas": "SEMANAS REGISTRADAS",
        "_frec_ref":          "FRECUENCIA REQUERIDA",
    }).sort_values(["Empleado", "ID del PDV"]).reset_index(drop=True)

    for c_idx, c_name in enumerate(out.columns, start=1):
        ws.cell(row=1, column=c_idx, value=c_name)
    for r_idx, row in enumerate(out.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            if pd.isna(val): val = None
            ws.cell(row=r_idx, column=c_idx, value=val)


def generar_adjuntos_por_supervisor(
    df_detalle: pd.DataFrame,
    df_cif_pdv: pd.DataFrame,
    df_np:      pd.DataFrame,
    df_pr:      pd.DataFrame,
    df_sos:     pd.DataFrame,
    mes: int,
    anio: int,
    seg_data: dict | None = None,
) -> dict:
    """
    Genera un Excel por supervisor con 7 hojas:
      1. Resumen Equipo
      2. CIF Detalle PDVs
      3. NP Detalle PDVs
      4. Precios y SOS Detalle PDVs
      5. ListSant     (productos nuevos — segmento Sandia)
      6. DoyPackBaby  (productos nuevos — DoyPack Baby)
      7. CremasBaby   (productos nuevos — Cremas Baby)

    `seg_data` es el dict producido por `_precargar_segmentos_nuevos`. Si
    se pasa None o queda vacío, las 3 hojas de productos nuevos se llenan
    con el mensaje "No aplica para este rol".

    Devuelve un dict {NOMBRE_SUPERVISOR: ruta_archivo}.
    """
    # Los adjuntos definitivos viven en la NUBE:
    #   Equipo Información/BI/INVOLVES/SALIDAS/ALERTAS/ADJUNTOS
    # Localmente sólo se deja una copia temporal, necesaria para que
    # `alertas_email` pueda adjuntar el archivo al correo.
    dir_adj = Path(tempfile.gettempdir()) / "alertas_adjuntos"
    dir_adj.mkdir(parents=True, exist_ok=True)

    seg_data = seg_data or {}

    # Sprint 17.15 — destinatarios: SUPERVISOR + GDD + LIDER.
    # Cada uno recibe un adjunto propio.
    if df_detalle.empty:
        supervisores_iter = []
    else:
        sups_df = (
            df_detalle[
                (df_detalle.get("ES_SUPERVISOR", False) == True) |
                (df_detalle.get("ES_GDD",         False) == True) |
                (df_detalle.get("ES_LIDER",       False) == True)
            ][["ACRONIMO", "NOMBRE"]]
                     .drop_duplicates(subset=["ACRONIMO"])
                     .sort_values("NOMBRE")
        )
        supervisores_iter = list(sups_df.itertuples(index=False, name=None))

    rutas = {}
    for acr_sup, nombre_sup in supervisores_iter:
        if not acr_sup or not nombre_sup:
            continue
        nombre_archivo = f"Detalle_{nombre_sup.replace(' ', '_')}_{mes:02d}_{anio}.xlsx"
        ruta = dir_adj / nombre_archivo

        wb = Workbook()
        # Hoja 1 — Resumen del equipo
        ws1 = wb.active
        ws1.title = "Resumen Equipo"
        _hoja_resumen_equipo(ws1, df_detalle, acr_sup)

        # Sprint 13.5: nombres alineados con el resumen del correo
        # Hoja 2 — CIF
        ws2 = wb.create_sheet("CIF — Detalle")
        _hoja_cif_detalle(ws2, df_cif_pdv, acr_sup)

        # Hoja 3 — No Presencia
        ws3 = wb.create_sheet("No Presencia — Detalle")
        _hoja_np_detalle(ws3, df_np, df_cif_pdv, acr_sup)

        # Hojas 4 y 5 — Precios y SOS separados (D18=Sí)
        ws4 = wb.create_sheet("Precios — Detalle")
        _hoja_precios_detalle(ws4, df_pr, acr_sup)
        ws5 = wb.create_sheet("SOS — Detalle")
        _hoja_sos_detalle(ws5, df_sos, acr_sup)

        # Hoja 6 — Impactos D&P (Sprint 13.6)
        ws6 = wb.create_sheet("Impactos — Detalle")
        _hoja_impactos_detalle(ws6, df_detalle, acr_sup, mes, anio)

        # Sprint 17.10 — hojas nuevas de Exhibiciones detalle
        ws_excap = wb.create_sheet("Exh Pagadas — No Capturadas")
        _hoja_exh_pagadas_no_capturadas(ws_excap, acr_sup, mes, anio, df_cif_pdv)

        ws_exgfr = wb.create_sheet("Exh Gratis — Fuera de Regla")
        _hoja_exh_gratis_fuera_regla(ws_exgfr, acr_sup, mes, anio, df_cif_pdv)

        # Hojas finales — Productos nuevos (uno por segmento)
        for nombre_segmento in _SEGMENTOS_NUEVOS:
            ws_seg = wb.create_sheet(nombre_segmento)
            _hoja_segmento_nuevo(ws_seg, nombre_segmento, seg_data, acr_sup)

        # Copia temporal local (para adjuntar al correo)
        wb.save(ruta)

        # Subida a la nube — destino oficial de los adjuntos
        buffer = io.BytesIO()
        wb.save(buffer)
        try:
            _subir_bytes_cloud(
                buffer.getvalue(),
                f"{RUTA_CARPETA_ADJUNTOS_CLOUD}/{nombre_archivo}",
                f"adjunto {nombre_archivo}",
            )
        except Exception as e:
            _log.warning("Fallo al subir adjunto %s a la nube: %s", nombre_archivo, e)

        # Indexamos el dict por NOMBRE del supervisor (lo que usa alertas_email
        # para resolver el destinatario).
        rutas[nombre_sup.upper()] = str(ruta)

    print(f"  ✅ {len(rutas)} adjuntos generados (9 hojas c/u)")
    return rutas


if __name__ == "__main__":
    main()