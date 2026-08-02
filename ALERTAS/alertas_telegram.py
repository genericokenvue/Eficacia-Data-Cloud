"""
alertas_telegram.py
───────────────────
Fase B del sistema de alertas de Eficacia.

Envía un mensaje push por Telegram a cada supervisor con el resumen
de cumplimiento de su equipo, desglosado por PERFIL DE PDV
(Directo+Droguerías vs Proximity) para los KPIs operativos, y con los
KPIs comerciales D&P en bloque aparte.

Estructura del mensaje (por petición del cliente)
─────────────────────────────────────────────────
    📊 *Cumplimiento Abril 2026*
    Fecha de corte: 04/05/2026

    Equipo: *NOMBRE SUPERVISOR*
    👥 Gestores a cargo: 7

    📋 *Indicadores Directo*
    `CIF          ` 87.3% ⚠️
    `NO PRESENCIA ` 92.0% ✅
    `PRECIOS      ` 78.6% ⚠️
    `SOS          ` 65.1% ❌

    📋 *Indicadores Proximity*
    `CIF          ` 91.0% ✅
    `NO PRESENCIA ` 88.5% ⚠️
    ...

    💼 *Indicador D&P Ind*
    `CUMPLIMIENTO CUOTA           ` 102.3% ✅
    `CUMPLIMIENTO IMPACTOS        ` 97.5% ✅
    `CUMPLIMIENTO IMP. SKU NUEVOS ` 81.0% ⚠️

    📌 *Adicionales D&P Ind*
    `% IMPACTOS MSL ` 74.0%

    🎯 *CUMPLIMIENTO GENERAL EQUIPO: 89.4% ⚠️*

    _El detalle completo fue enviado a tu correo._

Reglas de armado
────────────────
  • Una sección de perfil (Directo / Proximity) se OMITE entera si el
    supervisor no tiene PDVs en ese perfil.
  • La sección "Indicador D&P Ind" se OMITE si el supervisor no tiene
    datos D&P en el período.
  • La sección "Adicionales D&P Ind" se omite si MSL es N/D.
  • CUMPLIMIENTO GENERAL EQUIPO = promedio simple de los KPIs visibles
    (operativos por perfil + 3 KPIs D&P). MSL queda fuera por petición
    del cliente.

Modo de ejecución
─────────────────
    # Standalone (re-lee los archivos en disco; no puede desglosar por
    # perfil porque eso requiere los DataFrames crudos en memoria)
    python alertas_telegram.py

    # Como módulo desde run_alertas.py (recibe DataFrames en memoria)
    from alertas_telegram import enviar_resumen_telegram
    enviar_resumen_telegram(
        df_resumen, df_maestro, mes, anio,
        df_cif_pdv=..., df_np=..., df_pr=..., df_sos=...,
    )
"""

import io
import os
import re
import time
import urllib.parse
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────
# Cargar config.env al importar para que UMBRAL_OK/WARNING/TELEGRAM_TOKEN
# queden disponibles en os.environ ANTES de leer los floats de abajo.
from config_loader import cargar_config
cargar_config()

from alertas_logger import get_logger
_log = get_logger("alertas_telegram")

# alertas_logger añadió SCRIPTS/ a sys.path
import paths
BASE         = paths.BASE
DIR_ALERTAS  = paths.ALERTAS_DIR
RUTA_MAESTRA_CLOUD = f"{paths._SALIDAS_ROOT}/ALERTAS/maestro_supervisores.xlsx"


# ─────────────────────────────────────────────────────────────────────────────
# CONEXIÓN A LA NUBE (SharePoint vía Microsoft Graph API)
# ─────────────────────────────────────────────────────────────────────────────
_DRIVE_ID_CACHE: dict = {}


def _obtener_default_drive_id(token: str) -> str:
    if "drive_id" in _DRIVE_ID_CACHE:
        return _DRIVE_ID_CACHE["drive_id"]
    headers = {"Authorization": f"Bearer {token}"}
    nombre_sitio = getattr(paths, "SHAREPOINT_SITE_NAME", "eficacia")
    res_site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites?search={urllib.parse.quote(nombre_sitio)}",
        headers=headers,
    )
    if res_site.status_code == 200:
        sites = res_site.json().get("value", [])
        if sites:
            site = next(
                (s for s in sites
                 if nombre_sitio.lower() in (s.get("name", "").lower(), s.get("displayName", "").lower())),
                sites[0],
            )
            res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site.get('id')}/drive", headers=headers)
            if res_drive.status_code == 200:
                _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
                return _DRIVE_ID_CACHE["drive_id"]
    res_root = requests.get("https://graph.microsoft.com/v1.0/sites/root", headers=headers)
    if res_root.status_code == 200:
        site_id = res_root.json().get("id")
        res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
        if res_drive.status_code == 200:
            _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
            return _DRIVE_ID_CACHE["drive_id"]
    raise Exception(f"No se pudo obtener el Drive ID del sitio SharePoint '{nombre_sitio}': {res_site.text}")


def _limpiar_ruta_graph(ruta_sharepoint: str) -> str:
    ruta_sharepoint = ruta_sharepoint.replace("\\", "/")
    for prefijo in ["Documentos compartidos/", "Shared Documents/"]:
        if ruta_sharepoint.startswith(prefijo):
            ruta_sharepoint = ruta_sharepoint.replace(prefijo, "", 1)
    return ruta_sharepoint.lstrip("/")


def _listar_hijos_cloud(ruta_carpeta: str) -> list:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_codificada = urllib.parse.quote(_limpiar_ruta_graph(ruta_carpeta))
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{ruta_codificada}:/children"
    items = []
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Error al listar carpeta '{ruta_carpeta}': {response.status_code} - {response.text}")
        data = response.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _leer_excel_cloud(ruta_sharepoint: str, descripcion: str, **kwargs) -> pd.DataFrame:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    for r in (ruta_limpia, f"Documentos compartidos/{ruta_limpia}"):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(r)}:/content"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return pd.read_excel(io.BytesIO(response.content), **kwargs)
    raise FileNotFoundError(f"No se pudo leer {descripcion} desde SharePoint ({ruta_sharepoint}).")


def _ultimo_archivo_cloud(ruta_carpeta: str, patron_regex: str) -> str | None:
    """Devuelve la ruta SharePoint del último archivo (orden alfabético/nombre,
    igual que sorted(...)[-1] con glob local) que matchee patron_regex, o None."""
    try:
        hijos = _listar_hijos_cloud(ruta_carpeta)
    except Exception as e:
        _log.warning(f"No se pudo listar '{ruta_carpeta}' en SharePoint: {e}")
        return None
    nombres = sorted(
        h["name"] for h in hijos
        if h.get("file") and re.match(patron_regex, h.get("name", "")) and not h["name"].startswith("~$")
    )
    if not nombres:
        return None
    return f"{ruta_carpeta}/{nombres[-1]}"


def _resolver_exhib_data_dir_cloud() -> str:
    """Equivalente en la nube de paths._resolver_exhib_data_dir(): toma la
    subcarpeta más reciente tipo '01. ...' dentro de BASES/EXHIBICIONES."""
    raiz = paths.RUTA_CARPETA_BASES_EXHIB
    try:
        hijos = _listar_hijos_cloud(raiz)
    except Exception:
        return raiz
    candidatos = [h for h in hijos if h.get("folder") and re.match(r"^\d{2}\.\s", h.get("name", ""))]
    if not candidatos:
        return raiz
    candidatos.sort(key=lambda h: h.get("lastModifiedDateTime", ""), reverse=True)
    return f"{raiz}/{candidatos[0]['name']}"


def _cargar_token() -> str:
    """Devuelve TELEGRAM_TOKEN desde el entorno (config.env ya fue volcado)."""
    token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    # Defensa: rechazar marcadores y placeholders evidentes
    if not token or token.startswith("<") or "REEMPLAZAR" in token.upper():
        raise ValueError(
            "TELEGRAM_TOKEN no configurado.\n"
            "Edita ALERTAS/config.env y pon: TELEGRAM_TOKEN=NN:AA... (sin < >)"
        )
    return token


UMBRAL_OK      = float(os.environ.get("UMBRAL_OK",      "0.90"))
UMBRAL_WARNING = float(os.environ.get("UMBRAL_WARNING",  "0.70"))

# Segundos de espera entre mensajes (evita rate-limit de Telegram: 30 msg/s)
DELAY_ENTRE_MENSAJES = 0.5

# Pesos del CIF — deben coincidir con calcular_cumplimientos.py
# Pesos del CIF — fórmula V3 confirmada por el cliente:
#     TOTAL = 0.5·COBERTURA + 0.2·INTENSIDAD + 0.3·FRECUENCIA
# (V3 invirtió los pesos respecto a la fórmula original 0.5/0.2/0.3 que daba
#  más peso a HRS que a FREC.)
PESO_COB  = 0.50
PESO_INT  = 0.20    # antes PESO_HRS=0.30 — cambió por V3
PESO_FREC = 0.30    # antes PESO_FREC=0.20 — cambió por V3


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES DE FORMATO
# ─────────────────────────────────────────────────────────────────────────────

def _umbral_ok_de(col_o_kpi: str | None) -> float:
    """Resuelve el UMBRAL_OK específico (Sprint 13.4)."""
    if not col_o_kpi:
        return UMBRAL_OK
    try:
        from calcular_cumplimientos import umbral_de
        return umbral_de(col_o_kpi)
    except Exception:
        return UMBRAL_OK


def _semaforo(valor, col_o_kpi: str | None = None) -> str:
    if pd.isna(valor):
        return "—"
    umb_ok = _umbral_ok_de(col_o_kpi)
    if valor >= umb_ok:
        return "✅"
    if valor >= UMBRAL_WARNING:
        return "⚠️"
    return "❌"


def _pct(valor) -> str:
    """Formatea un float 0-1 como '87.3%'. NaN → 'N/D'."""
    if pd.isna(valor):
        return "N/D"
    return f"{valor * 100:.1f}%"


# ─────────────────────────────────────────────────────────────────────────────
# PERFILES DE PDV — Directo+Droguerías vs Proximity
# ─────────────────────────────────────────────────────────────────────────────

def _cargar_perfiles_pdv() -> dict[str, str]:
    """
    Lee BASE PUNTOS DE VENTA y devuelve:
        { id_pdv (str) -> 'DIR' | 'PROX' }

    Reglas:
        • DIRECTO    → 'DIR'
        • DROGUERIAS → 'DIR'
        • PROXIMITY  → 'PROX'
        • NaN / otro → no se incluye

    Si el archivo no existe, devuelve dict vacío y avisa por log.
    """
    ruta_cloud = f"{paths.RUTA_CARPETA_BASES_CIF}/Base_Ventas.xlsx"
    try:
        df = _leer_excel_cloud(
            ruta_cloud, "Base_Ventas",
            sheet_name="Informe_Punto_de_Venda",
            usecols=["ID", "Perfil del PDV"],
            dtype={"ID": str},
        )
    except Exception as e:
        _log.warning(
            f"Base_Ventas no se pudo leer desde SharePoint ({ruta_cloud}): {e} — "
            "el desglose por perfil no se podrá hacer."
        )
        return {}

    df["ID"] = df["ID"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df["Perfil del PDV"] = df["Perfil del PDV"].astype(str).str.strip().str.upper()

    mapa: dict[str, str] = {}
    for _, r in df.iterrows():
        perfil = r["Perfil del PDV"]
        if perfil in ("DIRECTO", "DROGUERIAS"):
            mapa[r["ID"]] = "DIR"
        elif perfil == "PROXIMITY":
            mapa[r["ID"]] = "PROX"
    return mapa


def _ids_supervisor_perfil(
    df_cif_pdv: pd.DataFrame,
    perfiles: dict[str, str],
    acr_sup: str,
    grupo_perfil: str,
) -> set[str]:
    """
    PDVs (ID_PDV_INVOLVES) del equipo del supervisor que pertenecen al
    grupo_perfil indicado ('DIR' o 'PROX'). Se usan los PDVs vistos por
    el equipo (gestor + sup propio) en el período activo del CIF.
    """
    if df_cif_pdv is None or df_cif_pdv.empty:
        return set()
    if "ACRONIMO_SUP" not in df_cif_pdv.columns:
        return set()
    if "ID_PDV_INVOLVES" not in df_cif_pdv.columns:
        return set()

    sub = df_cif_pdv[df_cif_pdv["ACRONIMO_SUP"] == acr_sup]
    if sub.empty:
        return set()

    ids = (
        sub["ID_PDV_INVOLVES"]
        .dropna().astype(str).str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .unique()
    )
    return {i for i in ids if perfiles.get(i) == grupo_perfil}


# ─────────────────────────────────────────────────────────────────────────────
# RECÁLCULO DE KPIs OPERATIVOS POR (SUPERVISOR, PERFIL)
# Estos cálculos son ad-hoc para el mensaje de Telegram (no afectan al
# adjunto Excel ni al correo, que siguen mostrando el global del supervisor).
# ─────────────────────────────────────────────────────────────────────────────

def _serie_id(df: pd.DataFrame, col: str = "ID_PDV_INVOLVES") -> pd.Series:
    """Normaliza una columna de IDs a string sin .0 trailing."""
    return (
        df[col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    )


def _kpi_cif_perfil(df_cif_pdv: pd.DataFrame, ids_pdv: set[str], acr_sup: str) -> float:
    """
    CIF agregado para los PDVs del supervisor en un perfil dado.
    Promedio ponderado por VENTAS_PROMEDIO_MES de los componentes
    COBERTURA / INTENSIDAD / FRECUENCIA con la fórmula V3:
        TOTAL = 0.5·COB + 0.2·INT + 0.3·FREC

    Acepta nombres legacy (COB/HRS/FREC) y nombres V3 (COBERTURA/INTENSIDAD/
    FRECUENCIA) para compatibilidad durante la transición.
    """
    if df_cif_pdv is None or df_cif_pdv.empty or not ids_pdv:
        return float("nan")
    if "ACRONIMO_SUP" not in df_cif_pdv.columns:
        return float("nan")

    d = df_cif_pdv[df_cif_pdv["ACRONIMO_SUP"] == acr_sup].copy()
    d["_id"] = _serie_id(d)
    d = d[d["_id"].isin(ids_pdv)]
    if d.empty:
        return float("nan")

    pesos = pd.to_numeric(d["VENTAS_PROMEDIO_MES"], errors="coerce").fillna(0)
    if pesos.sum() <= 0:
        return float("nan")

    # Tolerar V3 (COBERTURA/INTENSIDAD/FRECUENCIA) y legacy (COB/HRS/FREC).
    col_cob  = "COBERTURA"  if "COBERTURA"  in d.columns else "COB"
    col_int  = "INTENSIDAD" if "INTENSIDAD" in d.columns else "HRS"
    col_frec = "FRECUENCIA" if "FRECUENCIA" in d.columns else "FREC"

    cob  = pd.to_numeric(d[col_cob],  errors="coerce").fillna(0)
    intens = pd.to_numeric(d[col_int],  errors="coerce").fillna(0)
    frec = pd.to_numeric(d[col_frec], errors="coerce").fillna(0)

    s = pesos.sum()
    w_cob  = (cob    * pesos).sum() / s
    w_int  = (intens * pesos).sum() / s
    w_frec = (frec   * pesos).sum() / s
    return PESO_COB * w_cob + PESO_INT * w_int + PESO_FREC * w_frec


def _kpi_simple_perfil(
    df: pd.DataFrame,
    ids_pdv: set[str],
    acr_sup: str,
    col_planeado: str,
    col_ejecutado: str,
) -> float:
    """
    KPI estilo NP/PRECIOS/SOS: sum(ejecutado) / sum(planeado) sobre las
    filas del supervisor cuyos PDVs están en `ids_pdv`. Sólo cuenta filas
    con planeado >= 1 (consistente con calcular_cumplimientos).
    """
    if df is None or df.empty or not ids_pdv:
        return float("nan")
    if "ACRONIMO_SUP" not in df.columns or "ID_PDV_INVOLVES" not in df.columns:
        return float("nan")

    d = df[df["ACRONIMO_SUP"] == acr_sup].copy()
    if d.empty:
        return float("nan")
    d["_id"] = _serie_id(d)
    d = d[d["_id"].isin(ids_pdv)]
    if d.empty:
        return float("nan")

    plan = pd.to_numeric(d[col_planeado],  errors="coerce").fillna(0)
    eje  = pd.to_numeric(d[col_ejecutado], errors="coerce").fillna(0)
    mask = plan >= 1
    if not mask.any():
        return float("nan")

    s_plan = plan[mask].sum()
    if s_plan <= 0:
        return float("nan")
    return eje[mask].sum() / s_plan


_EXH_PAG_CACHE: pd.DataFrame | None = None
_EXH_GRA_CACHE: pd.DataFrame | None = None


def _cargar_exhibiciones_pagadas() -> pd.DataFrame:
    """Carga el reporte agrupado de exhibiciones pagadas (con ID_PDV_INVOLVES,
    CANTIDAD_PLANEADA, CANTIDAD_EJECUTADA, *EMPLEADO). Cache global."""
    global _EXH_PAG_CACHE
    if _EXH_PAG_CACHE is not None:
        return _EXH_PAG_CACHE

    candidatos = _ultimo_archivo_cloud(paths.RUTA_CARPETA_SALIDAS_EXHIB, r"^Resultado_exhibiciones_pagadas.*\.xlsx$")
    if not candidatos:
        _EXH_PAG_CACHE = pd.DataFrame()
        return _EXH_PAG_CACHE

    df = _leer_excel_cloud(candidatos, "Resultado exhibiciones pagadas")
    df["ID_PDV_INVOLVES"] = (
        df.get("ID_PDV_INVOLVES", "").astype(str).str.strip()
          .str.replace(r"\.0$", "", regex=True)
    )
    col_empleado = "*EMPLEADO" if "*EMPLEADO" in df.columns else "EMPLEADO"
    df["EMPLEADO_N"] = df.get(col_empleado, "").astype(str).str.strip().str.upper()
    df["CANTIDAD_PLANEADA"]  = pd.to_numeric(df.get("CANTIDAD_PLANEADA"),  errors="coerce").fillna(0)
    df["CANTIDAD_EJECUTADA"] = pd.to_numeric(df.get("CANTIDAD_EJECUTADA"), errors="coerce").fillna(0)
    _EXH_PAG_CACHE = df
    return _EXH_PAG_CACHE


def _cargar_exhibiciones_gratis() -> pd.DataFrame:
    """Carga el reporte de exhibiciones gratis con ID PDV + Empleado + Cantidad."""
    global _EXH_GRA_CACHE
    if _EXH_GRA_CACHE is not None:
        return _EXH_GRA_CACHE

    ruta_cloud = f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/Resultado exhibiciones gratis.xlsx"
    try:
        df = _leer_excel_cloud(ruta_cloud, "Resultado exhibiciones gratis")
    except Exception:
        _EXH_GRA_CACHE = pd.DataFrame()
        return _EXH_GRA_CACHE

    df.columns = [str(c).strip() for c in df.columns]
    df["ID PDV"]  = df.get("ID PDV", "").astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df["Empleado_N"] = df.get("Empleado", "").astype(str).str.strip().str.upper()
    df["Cantidad"]   = pd.to_numeric(df.get("Cantidad"), errors="coerce").fillna(0)
    df["Categoría"]  = df.get("Categoría", "").astype(str).str.strip().str.upper()
    # Solo categoría Gratis (alineado con etl_exhibiciones_gratis V3)
    df = df[df["Categoría"] == "GRATIS"].copy()
    _EXH_GRA_CACHE = df
    return _EXH_GRA_CACHE


def _kpi_exh_pagadas_perfil(
    ids_perfil: set[str],
    empleados_equipo: set[str],
) -> float:
    """EXH PAGADAS por canal: presencia 0/1 sobre PDVs del perfil del equipo
    del supervisor. Sum(ejecutado_REC) / Sum(planeado_REC)."""
    df = _cargar_exhibiciones_pagadas()
    if df.empty or not ids_perfil:
        return float("nan")
    d = df[
        df["ID_PDV_INVOLVES"].isin(ids_perfil)
        & df["EMPLEADO_N"].isin(empleados_equipo)
    ].copy()
    if d.empty:
        return float("nan")
    d["P_REC"] = (d["CANTIDAD_PLANEADA"]  > 0).astype(int)
    d["E_REC"] = (d["CANTIDAD_EJECUTADA"] > 0).astype(int)
    plan = d["P_REC"].sum()
    if plan <= 0:
        return float("nan")
    return d["E_REC"].sum() / plan


def _kpi_exh_gratis_promedio_persona(
    ids_perfil: set[str],
    empleados_equipo: set[str],
) -> tuple[float, int]:
    """EXH GRATIS prom/persona por canal:
        Σ exhibiciones_gratis_canal / # gestores con ≥1 PDV en ese canal.
    Devuelve (promedio, n_gestores_canal).
    """
    df = _cargar_exhibiciones_gratis()
    if df.empty or not ids_perfil:
        return float("nan"), 0
    d = df[
        df["ID PDV"].isin(ids_perfil)
        & df["Empleado_N"].isin(empleados_equipo)
    ].copy()
    if d.empty:
        return float("nan"), 0
    total_exh = d["Cantidad"].sum()
    # # gestores del equipo con ≥1 PDV en este canal
    n_gestores_canal = d["Empleado_N"].nunique()
    if n_gestores_canal <= 0:
        return float("nan"), 0
    return float(total_exh / n_gestores_canal), int(n_gestores_canal)


# ── Sprint 17 — captura del PLANNING + conteos por nivel ────────────────────

_EXH_PLANNING_CACHE: pd.DataFrame | None = None
_EXH_ENCUESTAS_CACHE: pd.DataFrame | None = None
TARGET_EXH_GRA_ALTO  = 3
TARGET_EXH_GRA_MEDIO = 5


def _cargar_planning_master() -> pd.DataFrame:
    """Carga PLANNING DE {Mes} {Año}.xlsx (master de qué se planeó)."""
    global _EXH_PLANNING_CACHE
    if _EXH_PLANNING_CACHE is not None:
        return _EXH_PLANNING_CACHE
    carpeta = _resolver_exhib_data_dir_cloud()
    ruta_cloud = _ultimo_archivo_cloud(carpeta, r"^PLANNING DE .*\.xlsx$")
    if not ruta_cloud:
        _EXH_PLANNING_CACHE = pd.DataFrame()
        return _EXH_PLANNING_CACHE
    df = _leer_excel_cloud(ruta_cloud, "Planning Maestro")
    col_pdv = next((c for c in df.columns if "punto de venta" in str(c).lower()), None)
    col_emp = next((c for c in df.columns if "empleado" in str(c).lower()), None)
    if not (col_pdv and col_emp):
        _EXH_PLANNING_CACHE = pd.DataFrame()
        return _EXH_PLANNING_CACHE
    out = pd.DataFrame({
        "PDV_N":      df[col_pdv].astype(str).str.strip().str.upper(),
        "EMPLEADO_N": df[col_emp].astype(str).str.strip().str.upper(),
    })
    _EXH_PLANNING_CACHE = out.drop_duplicates()
    return _EXH_PLANNING_CACHE


def _cargar_encuestas_planning() -> set[str]:
    """Devuelve set de PDV_N (upper) que respondieron la encuesta del planning."""
    global _EXH_ENCUESTAS_CACHE
    if _EXH_ENCUESTAS_CACHE is not None:
        return set(_EXH_ENCUESTAS_CACHE["PDV_N"].unique())
    carpeta = _resolver_exhib_data_dir_cloud()
    ruta_cloud = _ultimo_archivo_cloud(carpeta, r"^Base Exhibiciones Planning .*\.xlsx$")
    if not ruta_cloud:
        _EXH_ENCUESTAS_CACHE = pd.DataFrame(columns=["PDV_N"])
        return set()
    df = _leer_excel_cloud(ruta_cloud, "Base Exhibiciones Planning")
    col_pdv = "PDV" if "PDV" in df.columns else next(
        (c for c in df.columns if str(c).strip().upper() == "PDV"), None)
    if not col_pdv:
        _EXH_ENCUESTAS_CACHE = pd.DataFrame(columns=["PDV_N"])
        return set()
    _EXH_ENCUESTAS_CACHE = pd.DataFrame({
        "PDV_N": df[col_pdv].astype(str).str.strip().str.upper()
    }).drop_duplicates()
    return set(_EXH_ENCUESTAS_CACHE["PDV_N"].unique())


def _kpi_exh_pag_captura_perfil(
    ids_perfil: set[str],
    empleados_equipo: set[str],
) -> float:
    """
    Captura del PLANNING por canal del equipo del supervisor:
      Numerador: PDVs del PLANNING (en el perfil + asignados al equipo)
                 que respondieron la encuesta.
      Denominador: PDVs del PLANNING del perfil asignados al equipo.

    Como `ids_perfil` está en ID_PDV_INVOLVES y el PLANNING usa nombre del PDV,
    cruzamos vía PUNTO DE VENTA del archivo de exhibiciones pagadas.
    """
    plan = _cargar_planning_master()
    if plan.empty:
        return float("nan")
    capturados = _cargar_encuestas_planning()

    # Filtrar plan a los empleados del equipo
    plan_eq = plan[plan["EMPLEADO_N"].isin(empleados_equipo)].copy()
    if plan_eq.empty:
        return float("nan")

    # Cruzar a IDs del perfil vía exh_pagadas (que tiene PDV ↔ ID_PDV_INVOLVES)
    df_pag = _cargar_exhibiciones_pagadas()
    if df_pag.empty:
        return float("nan")
    col_pdv = "*PUNTO DE VENTA" if "*PUNTO DE VENTA" in df_pag.columns else "PUNTO DE VENTA"
    if col_pdv not in df_pag.columns:
        return float("nan")
    mapa = (
        df_pag[[col_pdv, "ID_PDV_INVOLVES"]]
        .assign(PDV_N=lambda d: d[col_pdv].astype(str).str.strip().str.upper())
        .drop_duplicates("PDV_N")
        .set_index("PDV_N")["ID_PDV_INVOLVES"].to_dict()
    )
    plan_eq["ID_PDV_INVOLVES"] = plan_eq["PDV_N"].map(mapa)
    plan_eq = plan_eq[plan_eq["ID_PDV_INVOLVES"].isin(ids_perfil)]
    if plan_eq.empty:
        return float("nan")

    plan_eq["CAPTURADO"] = plan_eq["PDV_N"].isin(capturados).astype(int)
    denom = len(plan_eq)
    return float(plan_eq["CAPTURADO"].sum() / denom) if denom else float("nan")


def _kpi_exh_gratis_por_nivel_perfil(
    ids_perfil: set[str],
    universo_canal: set[str],
) -> dict:
    """
    Sprint 17.8 — granularidad por mercaderista (no por suma del equipo).

    Cuenta cuántos mercaderistas del canal cumplen su target individual:
      • ALTO  cumple si tiene ≥3 exhibiciones ALTO IMPACTO en el canal.
      • MEDIO cumple si tiene ≥5 exhibiciones MEDIO IMPACTO en el canal.

    `universo_canal`: NOMBRE (upper) de los gestores del equipo con ≥1 PDV
    en ese canal según CIF (calculado en _kpis_por_perfil con df_cif_pdv).

    Devuelve:
        n_gestores_canal,
        n_cumplen_alto, n_cumplen_medio,
        cump_alto_frac (n_cumplen_alto / n_gestores_canal),
        cump_medio_frac (idem para medio),
        alto, medio, total (sumas brutas, informativas).
    """
    out = {
        "alto": 0, "medio": 0, "total": 0,
        "n_gestores_canal": 0,
        "n_cumplen_alto": 0, "n_cumplen_medio": 0,
        "cump_alto_frac": float("nan"),
        "cump_medio_frac": float("nan"),
    }
    if not ids_perfil or not universo_canal:
        return out
    n_gest_canal = len(universo_canal)
    universo = universo_canal

    df = _cargar_exhibiciones_gratis()
    if df.empty:
        out["n_gestores_canal"] = n_gest_canal
        return out
    d = df[
        df["ID PDV"].isin(ids_perfil)
        & df["Empleado_N"].isin(universo)
    ].copy()

    alto_total = medio_total = 0.0
    n_cumplen_alto = n_cumplen_medio = 0
    if not d.empty:
        nivel = d["Nivel Impacto"].astype(str).str.strip().str.upper()
        d["_ALTO"]  = d["Cantidad"].where(nivel == "ALTO IMPACTO",  0)
        d["_MEDIO"] = d["Cantidad"].where(nivel == "MEDIO IMPACTO", 0)
        por_gestor = d.groupby("Empleado_N", as_index=False).agg(
            ALTO=("_ALTO", "sum"),
            MEDIO=("_MEDIO", "sum"),
        )
        alto_total  = float(por_gestor["ALTO"].sum())
        medio_total = float(por_gestor["MEDIO"].sum())
        n_cumplen_alto  = int((por_gestor["ALTO"]  >= TARGET_EXH_GRA_ALTO).sum())
        n_cumplen_medio = int((por_gestor["MEDIO"] >= TARGET_EXH_GRA_MEDIO).sum())

    return {
        "alto": alto_total, "medio": medio_total,
        "total": alto_total + medio_total,
        "n_gestores_canal": n_gest_canal,
        "n_cumplen_alto":  n_cumplen_alto,
        "n_cumplen_medio": n_cumplen_medio,
        "cump_alto_frac":  n_cumplen_alto  / n_gest_canal,
        "cump_medio_frac": n_cumplen_medio / n_gest_canal,
    }


def _empleados_del_equipo(df_cif_pdv: pd.DataFrame | None, acr_sup: str) -> set[str]:
    """Devuelve el set de NOMBRE (upper) de los gestores cuyo ACRONIMO_SUP=acr_sup."""
    if df_cif_pdv is None or df_cif_pdv.empty or "ACRONIMO_SUP" not in df_cif_pdv.columns:
        return set()
    sub = df_cif_pdv[df_cif_pdv["ACRONIMO_SUP"] == acr_sup]
    if sub.empty or "NOMBRE" not in sub.columns:
        return set()
    return set(sub["NOMBRE"].astype(str).str.strip().str.upper().unique())


def _empleados_del_canal(
    df_cif_pdv: pd.DataFrame | None,
    acr_sup: str,
    ids_perfil: set[str],
) -> set[str]:
    """
    Sprint 17.8 — gestores del equipo que tienen ≥1 PDV en el canal indicado.
    Universo derivado de CIF (lo más confiable: refleja el plan asignado).
    """
    if df_cif_pdv is None or df_cif_pdv.empty:
        return set()
    if not ids_perfil:
        return set()
    sub = df_cif_pdv[df_cif_pdv["ACRONIMO_SUP"] == acr_sup].copy()
    if sub.empty:
        return set()
    sub["ID_PDV_S"] = (
        sub["ID_PDV_INVOLVES"].astype(str).str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    sub = sub[sub["ID_PDV_S"].isin(ids_perfil)]
    if sub.empty:
        return set()
    return set(sub["NOMBRE"].astype(str).str.strip().str.upper().unique())


def _kpis_por_perfil(
    perfiles: dict[str, str],
    df_cif_pdv: pd.DataFrame | None,
    df_np: pd.DataFrame | None,
    df_pr: pd.DataFrame | None,
    df_sos: pd.DataFrame | None,
    acr_sup: str,
) -> dict[str, dict[str, float]]:
    """
    Calcula los KPIs operativos para cada grupo de perfil del supervisor.
    Sprint 14.6: incluye también EXH_PAG (%) y EXH_GRATIS prom/persona por canal.

    Devuelve:
      {
        'DIR':  {'CIF', 'NP', 'PR', 'SOS', 'EXH_PAG', 'EXH_GRA_PROM',
                  'EXH_GRA_N_GESTORES', 'N_PDVS'},
        'PROX': {...},
      }
    Si un grupo no tiene PDVs (`N_PDVS=0`) la sección se omitirá del mensaje.
    """
    empleados_equipo = _empleados_del_equipo(df_cif_pdv, acr_sup)
    out = {}
    for grupo in ("DIR", "PROX"):
        ids = _ids_supervisor_perfil(df_cif_pdv, perfiles, acr_sup, grupo)
        if not ids:
            out[grupo] = {"CIF": float("nan"), "NP": float("nan"),
                          "PR": float("nan"), "SOS": float("nan"),
                          "EXH_PAG": float("nan"),
                          "EXH_PAG_CAP": float("nan"),
                          "EXH_GRA": {"alto":0,"medio":0,"total":0,
                                       "n_gestores_canal":0,
                                       "n_cumplen_alto":0,"n_cumplen_medio":0,
                                       "cump_alto_frac":float("nan"),
                                       "cump_medio_frac":float("nan")},
                          "N_PDVS": 0}
            continue
        exh_pag = _kpi_exh_pagadas_perfil(ids, empleados_equipo)
        exh_pag_cap = _kpi_exh_pag_captura_perfil(ids, empleados_equipo)
        universo_canal = _empleados_del_canal(df_cif_pdv, acr_sup, ids)
        exh_gra_d   = _kpi_exh_gratis_por_nivel_perfil(ids, universo_canal)
        out[grupo] = {
            "CIF":   _kpi_cif_perfil(df_cif_pdv, ids, acr_sup),
            "NP":    _kpi_simple_perfil(df_np,  ids, acr_sup, "PLANEADO_MES",     "CAPTURA_MES"),
            "PR":    _kpi_simple_perfil(df_pr,  ids, acr_sup, "CAPTURA_PLANEADA", "CAPTURA_EJECUTADA"),
            "SOS":   _kpi_simple_perfil(df_sos, ids, acr_sup, "CAPTURA_PLANEADA", "CAPTURA_EJECUTADA"),
            "EXH_PAG":     exh_pag,
            "EXH_PAG_CAP": exh_pag_cap,
            "EXH_GRA":     exh_gra_d,
            "N_PDVS": len(ids),
        }
    return out


# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCCIÓN DEL MENSAJE
# ─────────────────────────────────────────────────────────────────────────────

_MESES_ES = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
    5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
    9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre",
}


def _bloque_perfil(titulo: str, kpis: dict[str, float]) -> tuple[str | None, list[float]]:
    """Renderiza bloque "Indicadores X" (Directo o Proximity).

    Sprint 13.4: cada KPI usa su UMBRAL_OK específico (NP/Pr/SOS=1.00).
    Sprint 14.6: ahora incluye EXHIBICIONES PAGADAS (%) y EXHIBICIONES
    GRATIS PROMEDIO/PERSONA por canal.

    Devuelve (texto, valores_para_global). EXH_PAG entra al global;
    EXH_GRA_PROM NO (es conteo).
    """
    if not kpis or kpis.get("N_PDVS", 0) <= 0:
        return None, []

    lineas = [
        f"`CIF                              ` {_pct(kpis['CIF'])} {_semaforo(kpis['CIF'], 'CIF')}",
        f"`NO PRESENCIA                     ` {_pct(kpis['NP'])} {_semaforo(kpis['NP'], 'NP')}",
        f"`PRECIOS                          ` {_pct(kpis['PR'])} {_semaforo(kpis['PR'], 'PRECIOS')}",
        f"`SOS                              ` {_pct(kpis['SOS'])} {_semaforo(kpis['SOS'], 'SOS')}",
    ]
    valores_global = [kpis.get(k) for k in ("CIF", "NP", "PR", "SOS")]

    # Sprint 17 — Exhibiciones Pagadas: CAPTURA + EJECUCIÓN (ambos al global).
    exh_pag_cap = kpis.get("EXH_PAG_CAP", float("nan"))
    lineas.append(
        f"`CAPTURA EXH. PAGADAS             ` {_pct(exh_pag_cap)} {_semaforo(exh_pag_cap, 'EXHIB_PAG_CAPTURA')}"
    )
    valores_global.append(exh_pag_cap)

    exh_pag = kpis.get("EXH_PAG", float("nan"))
    lineas.append(
        f"`EJECUCIÓN EXH. PAGADAS           ` {_pct(exh_pag)} {_semaforo(exh_pag, 'EXHIB_PAG')}"
    )
    valores_global.append(exh_pag)

    # Sprint 17.8 — Exh Gratis: cuántos mercaderistas cumplen target individual.
    gra = kpis.get("EXH_GRA") or {}
    n_canal = int(gra.get("n_gestores_canal", 0))
    if n_canal > 0:
        n_a = int(gra.get("n_cumplen_alto", 0))
        n_m = int(gra.get("n_cumplen_medio", 0))
        cump_a = gra.get("cump_alto_frac",  float("nan"))
        cump_m = gra.get("cump_medio_frac", float("nan"))
        lineas.append(
            f"`EXH. GRATIS ALTO   (≥3/persona)  ` Cumplen {n_a} de {n_canal} {_semaforo(cump_a, 'EXHIB_GRA_ALTO')}"
        )
        lineas.append(
            f"`EXH. GRATIS MEDIO  (≥5/persona)  ` Cumplen {n_m} de {n_canal} {_semaforo(cump_m, 'EXHIB_GRA_MEDIO')}"
        )
        valores_global.extend([cump_a, cump_m])
    else:
        lineas.append("`EXH. GRATIS                      ` —")

    return f"📋 *{titulo}*\n" + "\n".join(lineas), valores_global


def _bloque_personal(row: pd.Series, rol_dest: str, n_gestores: int) -> tuple[str | None, list[float]]:
    """
    Sprint 17.15 — bloque "Cumplimiento personal" para destinatarios sin
    desglose por canal (GDD y Líder). Muestra los KPIs de la propia fila
    del resumen (CIF/NP/PR/SOS, Captura Exh Pagadas, Exh Gratis promedio
    de alto+medio capeado a 100%).

    Para LIDER el bloque agrega "n supervisores a cargo" en el titulo.
    """
    cif    = row.get("CIF_%")
    np_    = row.get("NP_%")
    pr     = row.get("PRECIOS_%")
    sos    = row.get("SOS_%")
    cap    = row.get("EXHIB_PAG_CAPTURA_%")
    alto   = row.get("EXHIB_GRA_ALTO_%")
    medio  = row.get("EXHIB_GRA_MEDIO_%")

    # Exh Gratis: promedio capeado a 100% por componente.
    def _cap(v): return min(float(v), 1.0) if not pd.isna(v) else float("nan")
    if not pd.isna(alto) or not pd.isna(medio):
        a_c = _cap(alto)  if not pd.isna(alto)  else 0.0
        m_c = _cap(medio) if not pd.isna(medio) else 0.0
        exh_gra_prom = (a_c + m_c) / 2
    else:
        exh_gra_prom = float("nan")

    if all(pd.isna(v) for v in (cif, np_, pr, sos, cap, alto, medio)):
        return None, []

    if rol_dest == "LIDER":
        n_sups = max(n_gestores, 0)
        titulo = f"Cumplimiento personal + {n_sups} supervisor(es) a cargo"
    else:
        titulo = "Cumplimiento personal"

    lineas = []
    valores_global = []
    for label, val, kpi in [
        ("CIF                              ", cif,    "CIF"),
        ("NO PRESENCIA                     ", np_,    "NP"),
        ("PRECIOS                          ", pr,     "PRECIOS"),
        ("SOS                              ", sos,    "SOS"),
        ("CAPTURA EXH. PAGADAS             ", cap,    "EXHIB_PAG_CAPTURA"),
        ("EXH. GRATIS (prom A+M)           ", exh_gra_prom, "EXHIB_GRATIS_PROM"),
    ]:
        lineas.append(f"`{label}` {_pct(val)} {_semaforo(val, kpi)}")
        if not pd.isna(val):
            valores_global.append(val)
    return f"📋 *{titulo}*\n" + "\n".join(lineas), valores_global


def _bloque_dyp(row: pd.Series) -> tuple[str | None, list[float]]:
    """
    Renderiza el bloque "Indicador D&P Ind" con CUOTA, IMPACTOS y SKU NUEVOS.
    Devuelve (texto_bloque_o_None, lista_kpis_para_global).
    """
    cuota   = row.get("VENTA_%")
    imp_gen = row.get("IMPACTOS_%")
    sku_nv  = row.get("PROD_NUEVOS_%")
    valores = [cuota, imp_gen, sku_nv]

    if all(pd.isna(v) for v in valores):
        return None, []

    lineas = [
        f"`CUMPLIMIENTO CUOTA           ` {_pct(cuota)} {_semaforo(cuota, 'VENTA')}",
        f"`CUMPLIMIENTO IMPACTOS        ` {_pct(imp_gen)} {_semaforo(imp_gen, 'IMPACTOS')}",
        f"`CUMPLIMIENTO IMP. SKU NUEVOS ` {_pct(sku_nv)} {_semaforo(sku_nv, 'PROD_NUEVOS')}",
    ]
    return "💼 *Indicador D&P Ind*\n" + "\n".join(lineas), valores


def _bloque_msl(row: pd.Series) -> str | None:
    """Renderiza la sección "Adicionales D&P Ind" con MSL (informativo)."""
    msl = row.get("MSL_%")
    if pd.isna(msl):
        return None
    return (
        "📌 *Adicionales D&P Ind*\n"
        f"`% IMPACTOS MSL ` {_pct(msl)}"
    )


def _bloque_exhibiciones(row: pd.Series) -> tuple[str | None, list[float]]:
    """
    Renderiza el bloque "Exhibiciones" para un gestor individual:
      • CAPTURA EXH. PAGADAS    — % captura del PLANNING (entra al global, Sprint 17)
      • EJECUCIÓN EXH. PAGADAS  — % cumplimiento ejecución (entra al global)
      • EXH. GRATIS ALTO/MEDIO  — conteo vs target 3/5 (entra al global, Sprint 17)
      • EXH. GRATIS TOTAL       — conteo informativo

    Devuelve (texto_bloque_o_None, lista_kpis_para_global).
    """
    exh_pag      = row.get("EXHIB_PAG_%")
    exh_pag_cap  = row.get("EXHIB_PAG_CAPTURA_%")
    gratis_total = row.get("EXHIB_GRATIS_TOTAL")
    gratis_alto  = row.get("EXHIB_GRATIS_ALTO")
    gratis_medio = row.get("EXHIB_GRATIS_MEDIO")
    cump_alto    = row.get("EXHIB_GRA_ALTO_%")
    cump_medio   = row.get("EXHIB_GRA_MEDIO_%")

    if all(pd.isna(v) for v in (exh_pag, exh_pag_cap, gratis_total)):
        return None, []

    lineas: list[str] = []
    valores_global: list[float] = []

    if not pd.isna(exh_pag_cap):
        lineas.append(
            f"`CAPTURA EXH. PAGADAS ` {_pct(exh_pag_cap)} {_semaforo(exh_pag_cap, 'EXHIB_PAG_CAPTURA')}"
        )
        valores_global.append(exh_pag_cap)

    if not pd.isna(exh_pag):
        lineas.append(
            f"`EJECUCIÓN EXH. PAG.  ` {_pct(exh_pag)} {_semaforo(exh_pag, 'EXHIB_PAG')}"
        )
        valores_global.append(exh_pag)

    if not pd.isna(gratis_alto) or not pd.isna(gratis_medio):
        try:
            a = int(gratis_alto) if not pd.isna(gratis_alto) else 0
            m = int(gratis_medio) if not pd.isna(gratis_medio) else 0
            t = a + m
            lineas.append(
                f"`EXH. GRATIS ALTO     ` {a}/3 ({_pct(cump_alto)}) {_semaforo(cump_alto, 'EXHIB_GRA_ALTO')}"
            )
            lineas.append(
                f"`EXH. GRATIS MEDIO    ` {m}/5 ({_pct(cump_medio)}) {_semaforo(cump_medio, 'EXHIB_GRA_MEDIO')}"
            )
            lineas.append(
                f"`EXH. GRATIS TOTAL    ` {t}"
            )
            if not pd.isna(cump_alto):  valores_global.append(cump_alto)
            if not pd.isna(cump_medio): valores_global.append(cump_medio)
        except (TypeError, ValueError):
            pass

    if not lineas:
        return None, []

    return "🎁 *Exhibiciones*\n" + "\n".join(lineas), valores_global


def _construir_mensaje(
    row: pd.Series,
    mes: int,
    anio: int,
    df_cif_pdv: pd.DataFrame | None,
    df_np: pd.DataFrame | None,
    df_pr: pd.DataFrame | None,
    df_sos: pd.DataFrame | None,
    perfiles: dict[str, str],
    fecha_corte: str,
    rango_periodo: dict | None = None,
) -> str:
    """
    Construye el mensaje en Markdown legacy (parse_mode="Markdown").

    Estructura:
      Encabezado (mes + año + fecha de corte)
      Equipo + gestores a cargo
      Indicadores Directo  (omitido si sin PDVs)
      Indicadores Proximity (omitido si sin PDVs)
      Indicador D&P Ind (omitido si todo NaN)
      Adicionales D&P Ind (omitido si MSL NaN)
      CUMPLIMIENTO GENERAL EQUIPO (promedio simple de visibles, sin MSL)
    """
    mes_nombre = _MESES_ES.get(mes, str(mes))
    supervisor = row.get("SUPERVISOR_LIDER", row.get("NOMBRE_SUPERVISOR", "—"))
    n_gestores = int(row.get("N_GESTORES", 0))
    # df_resumen usa ACRONIMO_SUP; aceptamos ambos por compatibilidad
    acr_sup    = row.get("ACRONIMO_SUP", row.get("ACRONIMO", ""))
    rol_dest   = str(row.get("_ROL_DEST", "SUPERVISOR"))

    # Sprint 17.15 — Si destinatario es GDD o LIDER, no aplica el desglose
    # por canal (no tiene equipo de gestores). Mostramos bloque "Cumplimiento
    # personal" con los KPIs de su propia fila en df_resumen.
    if rol_dest in ("GDD", "LIDER"):
        bloque_personal, vals_pers = _bloque_personal(row, rol_dest, n_gestores)
        bloque_dir = bloque_prox = None
        vals_dir = vals_prox = []
        if rol_dest == "LIDER":
            # LIDER de ejecución no es responsable de canal D&P; sin cuota/impactos/MSL/prod nuevos
            bloque_dyp_text, vals_dyp = None, []
            bloque_msl_text = None
        else:
            bloque_dyp_text, vals_dyp = _bloque_dyp(row)
            bloque_msl_text           = _bloque_msl(row)
    else:
        bloque_personal = None
        vals_pers = []
        # ── KPIs operativos por perfil ───────────────────────────────────────
        if acr_sup:
            kpis_perfil = _kpis_por_perfil(
                perfiles, df_cif_pdv, df_np, df_pr, df_sos, acr_sup,
            )
        else:
            kpis_perfil = {"DIR": {"N_PDVS": 0}, "PROX": {"N_PDVS": 0}}

        bloque_dir,  vals_dir  = _bloque_perfil("Indicadores Directo",   kpis_perfil["DIR"])
        bloque_prox, vals_prox = _bloque_perfil("Indicadores Proximity", kpis_perfil["PROX"])

        # ── KPIs comerciales D&P ─────────────────────────────────────────────
        bloque_dyp_text, vals_dyp = _bloque_dyp(row)
        bloque_msl_text           = _bloque_msl(row)

    # ── CUMPLIMIENTO GENERAL EQUIPO (promedio simple de KPIs visibles) ───
    valores_global: list[float] = []
    for vals in (vals_dir, vals_prox, vals_pers):
        for v in vals:
            if v is not None and not pd.isna(v):
                valores_global.append(v)
    if bloque_dyp_text is not None:
        for v in vals_dyp:
            if v is not None and not pd.isna(v):
                valores_global.append(v)
    # MSL y EXH GRATIS PROM NO entran al global (informativos)

    if valores_global:
        global_pct = sum(valores_global) / len(valores_global)
    else:
        global_pct = float("nan")

    # ── Armado del mensaje (Sprint 14.6: sin bloque exhibiciones aparte) ─
    bloques = [b for b in (bloque_personal, bloque_dir, bloque_prox,
                            bloque_dyp_text, bloque_msl_text) if b]

    # Sprint 13.1 + 13.7: rango de fechas + avance esperado A+B
    if rango_periodo and rango_periodo.get("rango_legible"):
        encabezado_periodo = (
            f"📊 *Cumplimiento {mes_nombre} {anio}*\n"
            f"Corte: {rango_periodo['rango_legible']}\n"
        )
        dias_t = rango_periodo.get("dias_transcurridos", 0)
        dias_m = rango_periodo.get("dias_mes_total", 0)
        avance = rango_periodo.get("avance_esperado_pct", 0)
        if dias_m > 0:
            nota_avance = (
                f"\n_📖 Snapshot al día {dias_t} de {dias_m} — "
                f"avance esperado del mes: {int(avance*100)}%_\n"
            )
        else:
            nota_avance = ""
    else:
        encabezado_periodo = (
            f"📊 *Cumplimiento {mes_nombre} {anio}*\n"
            f"Fecha de corte: {fecha_corte}\n"
        )
        nota_avance = ""

    msg = (
        encabezado_periodo +
        f"\n"
        f"Equipo: *{supervisor}*\n"
        f"👥 Gestores a cargo: {n_gestores}\n"
        f"\n"
        + "\n\n".join(bloques) +
        f"\n\n"
        f"🎯 *CUMPLIMIENTO GENERAL EQUIPO: {_pct(global_pct)} {_semaforo(global_pct, 'GLOBAL')}*\n"
        + nota_avance +
        f"\n"
        f"_El detalle completo fue enviado a tu correo._"
    )
    return msg


# ─────────────────────────────────────────────────────────────────────────────
# ENVÍO HTTP CON REINTENTOS
# ─────────────────────────────────────────────────────────────────────────────

# Reintentos para errores transitorios (ConnectionReset / 5xx / timeout).
# Backoff exponencial corto: 1s, 2s, 4s. Da tiempo a que se libere la
# conexión sin disparar el rate limit de Telegram (30 msg/s).
_MAX_REINTENTOS    = 3
_BACKOFF_INICIAL_S = 1.0


def _enviar_mensaje(token: str, chat_id: str, texto: str) -> tuple[bool, str]:
    """
    Envía un mensaje a Telegram via Bot API. Reintenta automáticamente en
    errores transitorios (ConnectionReset, Timeout, HTTP 5xx) con backoff
    exponencial. Errores 4xx (chat_id inválido, token mal, etc.) NO se
    reintentan — fallan inmediatamente.
    Devuelve (éxito: bool, detalle: str).
    """
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id"    : chat_id,
        "text"       : texto,
        "parse_mode" : "Markdown",
    }

    espera = _BACKOFF_INICIAL_S
    ultimo_error = "sin intento"

    for intento in range(1, _MAX_REINTENTOS + 1):
        try:
            resp = requests.post(url, json=payload, timeout=15)
            try:
                data = resp.json()
            except ValueError:
                data = {}

            if resp.status_code == 200 and data.get("ok"):
                return True, "OK" if intento == 1 else f"OK (reintento {intento})"

            # 4xx → error definitivo (no reintentar)
            if 400 <= resp.status_code < 500:
                detalle = data.get("description", f"HTTP {resp.status_code}")
                return False, detalle

            # 5xx → transitorio, reintentar
            ultimo_error = data.get("description", f"HTTP {resp.status_code}")

        except requests.exceptions.Timeout:
            ultimo_error = "Timeout (15s)"
        except requests.exceptions.ConnectionError as e:
            ultimo_error = f"ConnReset/error de conexión: {type(e).__name__}"
        except Exception as e:
            ultimo_error = f"Excepción inesperada: {e}"

        # Si quedan reintentos, esperar y repetir
        if intento < _MAX_REINTENTOS:
            time.sleep(espera)
            espera *= 2

    return False, f"{ultimo_error} (tras {_MAX_REINTENTOS} intentos)"


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL EXPORTADA
# ─────────────────────────────────────────────────────────────────────────────

def enviar_resumen_telegram(
    df_resumen: pd.DataFrame,
    df_maestro: pd.DataFrame,
    mes: int,
    anio: int,
    modo_prueba: bool = False,
    df_cif_pdv: pd.DataFrame | None = None,
    df_np:      pd.DataFrame | None = None,
    df_pr:      pd.DataFrame | None = None,
    df_sos:     pd.DataFrame | None = None,
    rango_periodo: dict | None = None,
) -> dict:
    """
    Envía el resumen de cumplimiento a cada supervisor vía Telegram.

    Parámetros
    ──────────
    df_resumen   : DataFrame de calcular_cumplimientos.main()['df_resumen'].
                   Debe contener al menos: SUPERVISOR_LIDER, ACRONIMO,
                   N_GESTORES, VENTA_%, IMPACTOS_%, MSL_%, PROD_NUEVOS_%.
    df_maestro   : DataFrame leído de MAESTRO_SUPERVISORES.xlsx
    mes, anio    : Periodo del reporte
    modo_prueba  : Si True, imprime los mensajes sin enviar (dry-run)
    df_cif_pdv,
    df_np, df_pr,
    df_sos       : DataFrames crudos enriquecidos con ACRONIMO_SUP. Si se
                   proveen, los KPIs operativos se desglosan por perfil de
                   PDV (Directo+Droguerías vs Proximity). Si no, el
                   mensaje cae al modo legacy (KPIs globales del resumen).

    Retorna
    ───────
    {
      'enviados':    [supervisores con envío exitoso],
      'fallidos':    [(supervisor, motivo)],
      'sin_chat_id': [supervisores sin chat_id en la maestra],
    }
    """
    print("\n" + "═" * 55)
    print("  ALERTAS TELEGRAM — Eficacia")
    print("═" * 55)

    if modo_prueba:
        print("  ⚠️  MODO PRUEBA — los mensajes NO se enviarán\n")

    # ── Cargar token ──────────────────────────────────────────────────────
    if not modo_prueba:
        try:
            token = _cargar_token()
        except ValueError as e:
            print(f"  ❌ {e}")
            return {"enviados": [], "fallidos": [], "sin_chat_id": []}
    else:
        token = "MODO_PRUEBA"

    # ── Mapa de perfiles de PDV (para desglose Directo/Proximity) ────────
    perfiles = _cargar_perfiles_pdv()
    if perfiles:
        print(f"  Perfiles PDV cargados: {len(perfiles):,} IDs mapeados")
    else:
        print("  ⚠️  Sin mapa de perfiles — el mensaje no podrá desglosar por perfil.")

    if any(d is None or d.empty for d in (df_cif_pdv, df_np, df_pr, df_sos)):
        print("  ⚠️  No se recibieron los DataFrames crudos — las secciones por")
        print("      perfil quedarán vacías (omitidas).")

    # Fecha de corte = ayer (día anterior al envío), formato DD/MM/YYYY
    fecha_corte = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
    print(f"  Fecha de corte usada en mensajes: {fecha_corte}")

    # ── Normalizar maestra ────────────────────────────────────────────────
    df_m = df_maestro.copy()
    df_m.columns = df_m.columns.str.strip().str.upper()
    df_m["NOMBRE_SUPERVISOR"] = df_m["NOMBRE_SUPERVISOR"].astype(str).str.strip().str.upper()
    # Excel suele leer el chat_id como float (1497788540.0). Quitamos el
    # trailing ".0" para que la Bot API lo acepte como entero válido.
    df_m["TELEGRAM_CHAT_ID"] = (
        df_m["TELEGRAM_CHAT_ID"].astype(str).str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    # ── Cruzar resumen con maestra ────────────────────────────────────────
    df_r = df_resumen.copy()
    df_r["SUPERVISOR_LIDER"] = df_r["SUPERVISOR_LIDER"].astype(str).str.strip().str.upper()

    df_merged = df_r.merge(
        df_m[["NOMBRE_SUPERVISOR", "TELEGRAM_CHAT_ID"]],
        left_on ="SUPERVISOR_LIDER",
        right_on="NOMBRE_SUPERVISOR",
        how="left",
    )

    # ── Clasificar por disponibilidad de chat_id ──────────────────────────
    sin_chat   = df_merged[df_merged["TELEGRAM_CHAT_ID"].isin(["", "nan", "NAN"])]["SUPERVISOR_LIDER"].tolist()
    con_chat   = df_merged[~df_merged["TELEGRAM_CHAT_ID"].isin(["", "nan", "NAN"])]

    if sin_chat:
        print(f"  ⚠️  {len(sin_chat)} supervisor(es) sin TELEGRAM_CHAT_ID en la maestra:")
        for s in sin_chat:
            print(f"     · {s}")
        print()

    # ── Enviar mensajes ───────────────────────────────────────────────────
    enviados   = []
    fallidos   = []

    print(f"  Enviando a {len(con_chat)} supervisor(es):\n")

    for _, row in con_chat.iterrows():
        supervisor = row["SUPERVISOR_LIDER"]
        chat_id    = str(row["TELEGRAM_CHAT_ID"]).strip()
        mensaje    = _construir_mensaje(
            row, mes, anio,
            df_cif_pdv=df_cif_pdv, df_np=df_np, df_pr=df_pr, df_sos=df_sos,
            perfiles=perfiles, fecha_corte=fecha_corte,
            rango_periodo=rango_periodo,
        )

        if modo_prueba:
            print(f"  ── [{supervisor}] — PREVIEW ──")
            print(mensaje)
            print()
            enviados.append(supervisor)
            continue

        ok, detalle = _enviar_mensaje(token, chat_id, mensaje)

        if ok:
            print(f"  ✅ {supervisor}")
            enviados.append(supervisor)
        else:
            print(f"  ❌ {supervisor} — {detalle}")
            fallidos.append((supervisor, detalle))

        time.sleep(DELAY_ENTRE_MENSAJES)

    # ── Resumen ───────────────────────────────────────────────────────────
    print(f"\n  Enviados: {len(enviados)}")
    if fallidos:
        print(f"  Fallidos: {len(fallidos)}")
        for sup, motivo in fallidos:
            print(f"    · {sup}: {motivo}")
    if sin_chat:
        print(f"  Sin chat_id: {len(sin_chat)}")
    _log.info(
        f"Telegram — enviados={len(enviados)} fallidos={len(fallidos)} "
        f"sin_chat_id={len(sin_chat)} prueba={modo_prueba}"
    )
    for sup, motivo in fallidos:
        _log.error(f"Telegram fallo: {sup} → {motivo}")

    print("═" * 55 + "\n")

    return {
        "enviados"    : enviados,
        "fallidos"    : fallidos,
        "sin_chat_id" : sin_chat,
    }


# ─────────────────────────────────────────────────────────────────────────────
# EJECUCIÓN STANDALONE
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Envío de alertas Telegram — Eficacia")
    parser.add_argument("--prueba", action="store_true",
                        help="Modo prueba: imprime mensajes sin enviar")
    args = parser.parse_args()

    # Sprint 17.6 — antes el standalone leía RESUMEN_CUMPLIMIENTO_*.xlsx desde
    # disco y NO podía desglosar por perfil (canales DIR/PROX) porque los DF
    # crudos df_cif_pdv/df_np/df_pr/df_sos sólo vivían en memoria. Resultado:
    # los mensajes salían sin bloques de canal.
    # Fix: invocar calcular_cumplimientos.main() para obtener TODO en memoria.
    import calcular_cumplimientos as cc
    out = cc.main()
    df_resumen    = out["df_resumen"]
    mes           = out["mes"]
    anio          = out["anio"]
    df_cif_pdv    = out.get("df_cif_pdv")
    df_np         = out.get("df_np")
    df_pr         = out.get("df_pr")
    df_sos        = out.get("df_sos")
    rango_periodo = out.get("rango_periodo")

    try:
        df_maestro = _leer_excel_cloud(RUTA_MAESTRA_CLOUD, "Maestra de supervisores")
    except Exception as e:
        print(f"❌ MAESTRO_SUPERVISORES.xlsx no encontrado en SharePoint: {e}")
        exit(1)
    if df_maestro.empty:
        print("❌ MAESTRO_SUPERVISORES.xlsx vacío.")
        exit(1)

    enviar_resumen_telegram(
        df_resumen, df_maestro, mes, anio,
        modo_prueba=args.prueba,
        df_cif_pdv=df_cif_pdv, df_np=df_np, df_pr=df_pr, df_sos=df_sos,
        rango_periodo=rango_periodo,
    )