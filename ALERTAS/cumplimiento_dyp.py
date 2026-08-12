"""
cumplimiento_dyp.py
───────────────────
Cálculo de los 4 KPIs comerciales D&P para el sistema de alertas:

  • VENTA_%        = sum(Venta) / sum(Cuota) por gestor (y por supervisor con
                     suma del equipo a cargo). Sobrecumplimiento permitido.
  • IMPACTOS_%     = sum(Clientes Impactados) / sum(Total Clientes a Impactar).
  • MSL_%          = % de PDVs del gestor que vendieron al menos 1 EAN MSL en
                     el mes (universo = todos los PDVs activos del gestor en
                     el mes).
  • PROD_NUEVOS_%  = promedio simple de los 3 segmentos:
                     ListSant_%, DoyPackBaby_%, CremasBaby_%.
                     Cada uno es % de PDVs target del segmento que vendieron
                     al menos 1 EAN del segmento en el mes.

Fuentes
───────
  • SALIDA/D&P/Consolidado_Impactos.csv  → VENTA_%, IMPACTOS_%
  • SALIDA/D&P/Consolidado_Ventas.csv    → MSL_%, PROD_NUEVOS_%
  • BASES/D&P/Listas/listas_referencia.xlsx
       (hojas MSL, ListSant, DoyPackBaby, CremasBaby)

Identidad de personas
─────────────────────
  • Asesor del CSV de impactos viene como "1001-NOMBRE APELLIDO".
    Se parsea para extraer el nombre y se compara en UPPER+strip con
    `NOMBRE` del PT (CIF/NP/Precios/SOS).
  • Vendedor del CSV de ventas viene ya limpio ("LEIDY LOPERA").
  • Supervisor en ambos CSVs cruza directo por nombre normalizado con
    SUPERVISOR_LIDER del PT.
  • Algunos transferencistas que están en estos CSVs pueden no estar en el
    PT — para ellos los KPIs operativos (CIF/NP/Precios/SOS) saldrán NaN,
    lo cual es correcto.

Manejo de archivos faltantes
────────────────────────────
Si los CSVs / listas no existen, las funciones devuelven DataFrames vacíos
(no levantan excepción). El ensamblaje de `calcular_cumplimientos.main()`
queda con los KPIs en NaN, idéntico al comportamiento de NP/Precios/SOS
cuando un archivo no está.

Casos especiales
────────────────
  • Cuota=0 y Venta=0 → VENTA_% = NaN (no aplica).
  • Cuota=0 y Venta>0 → VENTA_% = 2.0 (placeholder de sobrecumplimiento).
  • Total Clientes Impactar=0 → IMPACTOS_% = NaN.
  • Vendedor sin PDVs en el mes → MSL_%, PROD_NUEVOS_% = NaN.
"""

from __future__ import annotations

import io
import os
import re
import sys
import unicodedata
import tempfile
import urllib.parse
import requests
from pathlib import Path

import numpy as np
import pandas as pd

# Asegurar que SCRIPTS/ esté en sys.path ANTES de importar paths.
# (Cuando se importa via alertas_logger ya está, pero al usar este módulo
# de forma standalone — para tests o utilidades — necesitamos el fallback.)
# ALERTAS/ vive DENTRO de SCRIPTS/ (SCRIPTS/ALERTAS/cumplimiento_dyp.py), no
# es su hermana. `Path(__file__).resolve().parent.parent` YA apunta a
# SCRIPTS/ — no hay que agregarle "/SCRIPTS" de nuevo (eso generaba
# .../SCRIPTS/SCRIPTS, que no existe en ningún entorno).
_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

import paths


MESES_INT_A_STR = {
    1: "Enero",  2: "Febrero",  3: "Marzo",     4: "Abril",
    5: "Mayo",   6: "Junio",    7: "Julio",     8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# Si Cuota=0 con Venta>0 → marcador de sobrecumplimiento (200%).
SOBRECUMPLIMIENTO_PLACEHOLDER = 2.0


# ─────────────────────────────────────────────────────────────────────────────
# CONEXIÓN A LA NUBE (SharePoint vía Microsoft Graph API)
# ─────────────────────────────────────────────────────────────────────────────
# Los CSV/Excel de D&P ya no viven en disco local (paths.DYP_OUT_IMPACTOS,
# paths.DYP_OUT_VENTAS, paths.DYP_LISTAS_FILE) — se leen directamente de
# SharePoint, con la misma convención Graph API que el resto del pipeline.
#
# IMPORTANTE: etl_ventas.py sube los consolidados de Ventas/Impactos con el
# periodo en el nombre del archivo (ej. "Consolidado_Ventas_JULIO_2026.csv"),
# no como un único archivo acumulado — por eso estas rutas son funciones de
# (mes, año), no constantes fijas.
def _ruta_impactos_cloud(mes: int, anio: int) -> str:
    mes_up = MESES_INT_A_STR.get(int(mes), "").upper()
    return f"{paths.RUTA_CARPETA_SALIDAS_DYP}/Consolidado_Impactos_{mes_up}_{int(anio)}.csv"


def _ruta_ventas_cloud(mes: int, anio: int) -> str:
    mes_up = MESES_INT_A_STR.get(int(mes), "").upper()
    return f"{paths.RUTA_CARPETA_SALIDAS_DYP}/Consolidado_Ventas_{mes_up}_{int(anio)}.csv"


RUTA_LISTAS_CLOUD = f"{paths.RUTA_CARPETA_BASES_DYP}/Listas/MSL & Listas Target Catman.xlsx"

_DRIVE_ID_CACHE: dict = {}
_CACHE_DYP_BYTES: dict = {}   # ruta_cloud -> bytes | None (cache por proceso)


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


def _descargar_bytes_cloud(ruta_sharepoint: str, descripcion: str) -> bytes:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    response = None
    for r in (ruta_limpia, f"Documentos compartidos/{ruta_limpia}"):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(r)}:/content"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.content
    status = response.status_code if response is not None else "N/A"
    text = response.text if response is not None else "Sin respuesta"
    raise FileNotFoundError(f"No se pudo descargar {descripcion} desde SharePoint ({ruta_sharepoint}): {status} - {text}")


def _bytes_cloud_cached(ruta_cloud: str, descripcion: str) -> bytes | None:
    """Descarga (y cachea por proceso) el contenido crudo de un archivo en
    SharePoint. Devuelve None si no existe/falla — el llamador decide el
    mensaje/fallback, igual que antes con `Path.is_file()`."""
    if ruta_cloud in _CACHE_DYP_BYTES:
        return _CACHE_DYP_BYTES[ruta_cloud]
    try:
        contenido = _descargar_bytes_cloud(ruta_cloud, descripcion)
    except Exception:
        contenido = None
    _CACHE_DYP_BYTES[ruta_cloud] = contenido
    return contenido


def _leer_csv_cloud(ruta_cloud: str, descripcion: str, **kwargs) -> pd.DataFrame | None:
    """Lee un CSV desde SharePoint (bytes cacheados por proceso). None si no existe."""
    contenido = _bytes_cloud_cached(ruta_cloud, descripcion)
    if contenido is None:
        return None
    return pd.read_csv(io.BytesIO(contenido), **kwargs)


def _ruta_listas_temp_cloud() -> str | None:
    """
    `etl_impactos_segmentos.cargar_listas()` espera una RUTA de archivo en
    disco (no bytes). Descargamos el Excel de listas una sola vez por
    proceso a un temporal y devolvemos esa ruta; None si no existe en la nube.
    """
    cache_key = "_listas_tmp_path"
    if cache_key in _CACHE_DYP_BYTES:
        return _CACHE_DYP_BYTES[cache_key]
    contenido = _bytes_cloud_cached(RUTA_LISTAS_CLOUD, "Listas D&P (MSL/Prod Nuevos)")
    if contenido is None:
        _CACHE_DYP_BYTES[cache_key] = None
        return None
    tmp_path = Path(tempfile.gettempdir()) / "listas_dyp_temp.xlsx"
    tmp_path.write_bytes(contenido)
    _CACHE_DYP_BYTES[cache_key] = str(tmp_path)
    return str(tmp_path)


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────────────────

def armonizar_nombres_a_pt(
    df: pd.DataFrame,
    col_nombre: str,
    nombres_pt: set | list,
) -> pd.DataFrame:
    """
    Reemplaza los valores de `df[col_nombre]` por su match canónico de
    `nombres_pt` cuando se identifica una correspondencia por palabras
    (subset de tokens). Resuelve el problema de nombres truncados:

        D&P  → "CRISTHIAN ENCISO"
        PT   → "CRISTHIAN CAMILO ENCISO VARGAS"
        Match: True (ambas palabras del corto están en el largo)

    Sin match → se conserva el valor original (caso típico: transferencistas
    que sólo existen en D&P, no en el PT).

    Notas:
      • Ignora "fillers" como DE/LA/LOS para no inflar coincidencias.
      • Si un nombre D&P matchea con varios candidatos PT, toma el primero
        (suele indicar caso patológico — los nombres en el PT son únicos).
    """
    PALABRAS_FILLER = {"DE", "LA", "LAS", "LOS", "DEL", "Y"}

    def _palabras(s: str) -> set[str]:
        if not s or pd.isna(s):
            return set()
        return {w for w in str(s).upper().split() if w and w not in PALABRAS_FILLER}

    pt_index = []
    for nombre in nombres_pt:
        pal = _palabras(nombre)
        if pal:
            pt_index.append((nombre, pal))

    def _buscar(valor) -> str:
        pal_v = _palabras(valor)
        if not pal_v:
            return valor
        for nombre_pt, pal_pt in pt_index:
            # Match si el conjunto más corto está totalmente contenido en el
            # otro. Empate (mismo set) también cuenta.
            if pal_v.issubset(pal_pt) or pal_pt.issubset(pal_v):
                return nombre_pt
        return valor

    df = df.copy()
    df[col_nombre] = df[col_nombre].apply(_buscar)
    return df


def _norm_col(c) -> str:
    """
    Normaliza un nombre de columna para poder compararlo sin depender de
    mayúsculas, tildes, espacios o de que la Ñ haya llegado mal codificada.

    El "AÃ±o" que aparece al inspeccionar el CSV es la Ñ de "Año" en UTF-8
    leída como latin-1 (mojibake). Se contempla explícitamente para que la
    resolución funcione aunque el archivo cambie de codificación.
    """
    s = str(c).strip()
    s = s.replace("Ã±", "ñ").replace("Ã‘", "Ñ")     # repara mojibake común
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return " ".join(s.upper().split())


def _buscar_col(df, *alias, contiene: list[str] | None = None):
    """
    Devuelve el nombre REAL de la primera columna de `df` que coincida con
    alguno de los `alias` (comparando normalizado), o que contenga todas las
    palabras de `contiene`. None si no encuentra ninguna.

    FIX: este módulo pedía las columnas como "MES" y "AÑO" en mayúsculas,
    pero el Consolidado de Impactos las trae como "Mes" y "Año". Como el
    chequeo era `if "MES" not in df.columns`, la función se rendía y devolvía
    DataFrames vacíos → VENTA_%, IMPACTOS_%, MSL_% y PROD_NUEVOS_% salían en
    blanco para TODOS los gestores ("D&P — gestores con datos: 0"), sin que
    nada fallara de forma visible.
    """
    norm = {_norm_col(c): c for c in df.columns}
    for a in alias:
        real = norm.get(_norm_col(a))
        if real is not None:
            return real
    if contiene:
        req = [_norm_col(p) for p in contiene]
        for n, real in norm.items():
            if all(p in n for p in req):
                return real
    return None


def _num_es(serie) -> pd.Series:
    """
    Convierte a número tolerando el formato del Consolidado de Impactos, que
    usa COMA COMO SEPARADOR DE MILES ("111,000,000" = 111 millones).

    FIX: el CSV se lee con `decimal=","`, así que `pd.to_numeric` sobre esos
    valores devolvía NaN en 67 de 68 filas — la Cuota quedaba vacía y por
    tanto VENTA_% salía NaN para todos, aunque IMPACTOS_% sí funcionara (sus
    columnas son enteros pequeños sin separador). Se eliminan las comas y,
    si el valor quedara con más de un punto, también se tratan como miles.
    """
    s = serie.astype(str).str.strip()
    s = s.str.replace(",", "", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def parsear_nombre_asesor(s) -> str:
    """
    Extrae el nombre limpio de un asesor.
        "1001-DEISY MUNZA"   → "DEISY MUNZA"
        "1001 - DEISY MUNZA" → "DEISY MUNZA"
        "DEISY MUNZA"        → "DEISY MUNZA"
    Devuelve UPPER+strip; cadena vacía si no aplica.
    """
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    m = re.match(r"^\d+\s*[-]?\s*(.+)$", s)
    if m and m.group(1).strip():
        return m.group(1).strip()
    return s


def _ratio(num, den) -> float:
    """
    Cumplimiento = num / den.
      • den > 0: num/den (sin cap; sobrecumplimiento permitido).
      • den = 0, num = 0: NaN (no aplica).
      • den = 0, num > 0: SOBRECUMPLIMIENTO_PLACEHOLDER (=2.0).
    """
    try:
        num = float(num); den = float(den)
    except (TypeError, ValueError):
        return np.nan
    if pd.isna(num) or pd.isna(den):
        return np.nan
    if den > 0:
        return num / den
    return np.nan if num == 0 else SOBRECUMPLIMIENTO_PLACEHOLDER


def parsear_codigo_asesor(s) -> str:
    """
    Extrae el código numérico de un asesor en formato "1001-NOMBRE":
        "1001-DEISY MUNZA"   → "1001"
        "1001 - DEISY MUNZA" → "1001"
        "DEISY MUNZA"        → ""
    """
    if pd.isna(s):
        return ""
    s = str(s).strip()
    m = re.match(r"^(\d+)\s*[-]?", s)
    return m.group(1) if m else ""


def _ean_clean(v) -> str:
    """Convierte el EAN a entero-string sin notación científica ni .0 trailing."""
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s in ("", "nan"):
        return ""
    try:
        return str(int(float(s)))
    except Exception:
        return s


# ─────────────────────────────────────────────────────────────────────────────
# RESOLUCIÓN DE ACRONIMOS  (cruce D&P → Base cupos)
# ─────────────────────────────────────────────────────────────────────────────

def _enriquecer_con_acronimo(
    gestor: pd.DataFrame,
    supervisor: pd.DataFrame,
    mes: int,
    anio: int,
    base_cupos_idx: dict,
    nombres_supervisores_bc: set | list | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Añade columnas `ACRONIMO` (llave canónica de la persona) y, en gestor,
    `ACRONIMO_SUP` (llave del supervisor a cargo).

    Estrategia
    ──────────
    • Para `gestor`: re-leemos el CSV de impactos y de ventas para obtener
      el COD_ASESOR_ECOM (que existe sólo en el origen, no en los
      DataFrames agregados). Cruzamos ese código con `codigo_a_acronimo`
      de Base cupos para resolver el ACRONIMO. Cuando un asesor sólo
      apareció en ventas (no impactos), también lo cubrimos.
    • Para `gestor.ACRONIMO_SUP`: armonizamos el campo `Supervisor` del
      CSV (truncado) contra los nombres canónicos de supervisores de
      Base cupos, y mapeamos a su ACRONIMO.
    • Para `supervisor`: armonizamos `SUPERVISOR_LIDER` y mapeamos a ACR.

    Las filas sin match (ej. códigos huérfanos `999`, `1002`, `1003`) se
    descartan, según la decisión del usuario.
    """
    cod_a_acr     = base_cupos_idx["codigo_a_acronimo"]
    nombre_a_acr  = base_cupos_idx["nombre_a_acronimo"]

    mes_str  = MESES_INT_A_STR.get(int(mes), "")
    anio_str = str(int(anio))

    # ── Construir mapeo NOMBRE_PERSONA_DyP → (ACRONIMO, NOMBRE_SUPERVISOR_DyP) ─
    # Reusamos ambos CSVs (impactos + ventas) para cubrir casos donde la
    # persona aparece en uno pero no en el otro.
    pairs = []   # (nombre_dyp, acronimo, nombre_supervisor_dyp)

    if (df_i := _leer_csv_cloud(
        _ruta_impactos_cloud(mes, anio), "Consolidado_Impactos.csv",
        sep="|", decimal=",", encoding="utf-8",
        dtype={"MES": str, "AÑO": str}, low_memory=False,
    )) is not None and all(_buscar_col(df_i, a, b) is not None
                           for a, b in [("MES", "Mes"), ("AÑO", "Año"),
                                        ("Asesor", "Asesor"), ("Supervisor", "Supervisor")]):
        _cm = _buscar_col(df_i, "MES", "Mes")
        _ca = _buscar_col(df_i, "AÑO", "Año", "ANIO", "ANO")
        _cs = _buscar_col(df_i, "Asesor")
        _cv = _buscar_col(df_i, "Supervisor")
        df_i["MES"] = df_i[_cm].fillna("").astype(str).str.strip().str.capitalize()
        df_i["AÑO"] = df_i[_ca].fillna("").astype(str).str.strip()
        df_i = df_i[(df_i["MES"] == mes_str) & (df_i["AÑO"] == anio_str)].copy()
        df_i["_nombre"] = df_i[_cs].apply(parsear_nombre_asesor)
        df_i["_codigo"] = df_i[_cs].apply(parsear_codigo_asesor)
        df_i["_acr"]    = df_i["_codigo"].map(cod_a_acr).fillna("")
        df_i["_sup"]    = df_i[_cv].astype(str).str.strip().str.upper()
        for _, r in df_i.iterrows():
            pairs.append((r["_nombre"], r["_acr"], r["_sup"]))

    # FIX: usecols con nombres exactos lanzaba ValueError ante cualquier
    # cambio de capitalización/tilde y el except dejaba df_v = None, perdiendo
    # todo el mapeo de nombres desde ventas sin ningún aviso.
    df_v = _leer_csv_cloud(
        _ruta_ventas_cloud(mes, anio), "Consolidado_Ventas.csv",
        sep="|", decimal=",", encoding="utf-8",
        dtype=str, low_memory=False,
    )
    if df_v is not None:
        _vm  = _buscar_col(df_v, "MES", "Mes")
        _va  = _buscar_col(df_v, "AÑO", "Año", "ANIO", "ANO")
        _vc  = _buscar_col(df_v, "Cod. Vendedor", "Cod Vendedor", contiene=["COD", "VENDEDOR"])
        _vv  = _buscar_col(df_v, "Vendedor")
        _vs  = _buscar_col(df_v, "Supervisor")
        if None in (_vm, _va, _vc, _vv, _vs):
            print(f"  ⚠️  Consolidado_Ventas sin columnas para el mapeo de nombres; se omite.")
            df_v = None
    if df_v is not None:
        df_v["MES"] = df_v[_vm].fillna("").astype(str).str.strip().str.capitalize()
        df_v["AÑO"] = df_v[_va].fillna("").astype(str).str.strip()
        df_v = df_v[(df_v["MES"] == mes_str) & (df_v["AÑO"] == anio_str)].copy()
        df_v["_nombre"] = df_v[_vv].astype(str).str.strip().str.upper()
        df_v["_codigo"] = (
            pd.to_numeric(df_v[_vc], errors="coerce")
              .astype("Int64").astype(str).replace("<NA>", "")
        )
        df_v["_acr"]    = df_v["_codigo"].map(cod_a_acr).fillna("")
        df_v["_sup"]    = df_v[_vs].astype(str).str.strip().str.upper()
        v_unicas = df_v[["_nombre", "_acr", "_sup"]].drop_duplicates()
        for _, r in v_unicas.iterrows():
            pairs.append((r["_nombre"], r["_acr"], r["_sup"]))

    # Construir lookups: nombre_dyp → ACRONIMO  y  nombre_dyp → ACRONIMO_SUP
    map_acr = {}
    map_sup_raw = {}
    for nombre_dyp, acr, sup_raw in pairs:
        if not nombre_dyp:
            continue
        if acr and nombre_dyp not in map_acr:
            map_acr[nombre_dyp] = acr
        if sup_raw and nombre_dyp not in map_sup_raw:
            map_sup_raw[nombre_dyp] = sup_raw

    # Resolver ACRONIMO_SUP: armonizar nombre supervisor crudo vs Base cupos
    # y mapear a su ACRONIMO.
    nombres_sup_bc = set(nombres_supervisores_bc or [])

    def _palabras(s):
        FILLER = {"DE", "LA", "LAS", "LOS", "DEL", "Y"}
        return {w for w in str(s).upper().split() if w and w not in FILLER}

    pt_sup_idx = [(n, _palabras(n)) for n in nombres_sup_bc if n]

    def _resolver_acr_sup(sup_raw: str) -> str:
        if not sup_raw:
            return ""
        # 1) match exacto
        if sup_raw in nombre_a_acr and sup_raw in nombres_sup_bc:
            return nombre_a_acr[sup_raw]
        # 2) armonización por palabras
        pal_v = _palabras(sup_raw)
        if not pal_v:
            return ""
        for nombre_pt, pal_pt in pt_sup_idx:
            if pal_v.issubset(pal_pt) or pal_pt.issubset(pal_v):
                return nombre_a_acr.get(nombre_pt, "")
        return ""

    # ── Enriquecer gestor ─────────────────────────────────────────────────
    if not gestor.empty:
        gestor = gestor.copy()
        gestor["ACRONIMO"]     = gestor["NOMBRE"].map(map_acr).fillna("")
        gestor["_sup_raw"]     = gestor["NOMBRE"].map(map_sup_raw).fillna("")
        gestor["ACRONIMO_SUP"] = gestor["_sup_raw"].apply(_resolver_acr_sup)
        gestor = gestor.drop(columns=["_sup_raw"])
        # Descartar gestores cuyo COD_ASESOR_ECOM no tenía maestra
        gestor = gestor[gestor["ACRONIMO"] != ""].copy()

    # ── Enriquecer supervisor ─────────────────────────────────────────────
    if not supervisor.empty:
        supervisor = supervisor.copy()
        supervisor["ACRONIMO"] = supervisor["SUPERVISOR_LIDER"].apply(_resolver_acr_sup)
        # Si no se resolvió por armonización, intentar match exacto contra nombre_a_acr
        falta = supervisor["ACRONIMO"] == ""
        supervisor.loc[falta, "ACRONIMO"] = (
            supervisor.loc[falta, "SUPERVISOR_LIDER"].map(nombre_a_acr).fillna("")
        )
        supervisor = supervisor[supervisor["ACRONIMO"] != ""].copy()

    return gestor, supervisor


# ─────────────────────────────────────────────────────────────────────────────
# 1) VENTA_% e IMPACTOS_%  (desde Consolidado_Impactos.csv)
# ─────────────────────────────────────────────────────────────────────────────

def cumplimientos_venta_impactos(mes: int, anio: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Devuelve (df_gestor, df_supervisor) con columnas:
        gestor:    NOMBRE, SUPERVISOR_LIDER, CUOTA, VENTA, VENTA_%,
                   CLIENTES_IMPACTAR, CLIENTES_IMPACTADOS, IMPACTOS_%
        supervisor:SUPERVISOR_LIDER, CUOTA, VENTA, VENTA_%,
                   CLIENTES_IMPACTAR, CLIENTES_IMPACTADOS, IMPACTOS_%

    Si el CSV no existe o no hay filas para (mes, anio), devuelve frames vacíos.
    """
    cols_g = ["NOMBRE", "SUPERVISOR_LIDER",
              "CUOTA", "VENTA", "VENTA_%",
              "CLIENTES_IMPACTAR", "CLIENTES_IMPACTADOS", "IMPACTOS_%"]
    cols_s = ["SUPERVISOR_LIDER",
              "CUOTA", "VENTA", "VENTA_%",
              "CLIENTES_IMPACTAR", "CLIENTES_IMPACTADOS", "IMPACTOS_%"]

    mes_str  = MESES_INT_A_STR.get(int(mes), "")
    anio_str = str(int(anio))

    ruta_impactos = _ruta_impactos_cloud(mes, anio)
    df = _leer_csv_cloud(
        ruta_impactos, "Consolidado_Impactos.csv",
        sep="|", decimal=",", encoding="utf-8",
        dtype={"MES": str, "AÑO": str}, low_memory=False,
    )
    if df is None:
        print(f"  ⚠️  Falta Consolidado_Impactos en SharePoint ({ruta_impactos}) — VENTA_% e IMPACTOS_% serán NaN")
        return pd.DataFrame(columns=cols_g), pd.DataFrame(columns=cols_s)
    # FIX: se pedían "MES"/"AÑO" en mayúsculas y "Total Clientes Impactar",
    # pero el CSV real trae "Mes", "Año" y "Total Clientes a Impactar" (con
    # la "a"). El chequeo estricto hacía que la función retornara vacío y
    # VENTA_%/IMPACTOS_% quedaran en NaN para todos. Ahora se resuelven por
    # nombre normalizado, tolerando mayúsculas, tildes y mojibake.
    col_mes  = _buscar_col(df, "MES", "Mes")
    col_anio = _buscar_col(df, "AÑO", "Año", "ANIO", "ANO")
    col_ases = _buscar_col(df, "Asesor")
    col_sup  = _buscar_col(df, "Supervisor")
    col_cuo  = _buscar_col(df, "Cuota")
    col_ven  = _buscar_col(df, "Venta")
    col_plan = _buscar_col(df, "Total Clientes a Impactar", "Total Clientes Impactar",
                           contiene=["CLIENTES", "IMPACTAR"])
    col_real = _buscar_col(df, "Clientes Impactados", contiene=["CLIENTES", "IMPACTADOS"])

    faltan = [n for n, c in [("Mes", col_mes), ("Año", col_anio), ("Asesor", col_ases),
                             ("Supervisor", col_sup), ("Cuota", col_cuo), ("Venta", col_ven),
                             ("Clientes a Impactar", col_plan),
                             ("Clientes Impactados", col_real)] if c is None]
    if faltan:
        print(f"  ⚠️  {ruta_impactos} sin columnas esperadas: {faltan} — VENTA_% e IMPACTOS_% serán NaN")
        print(f"      Columnas encontradas: {list(df.columns)}")
        return pd.DataFrame(columns=cols_g), pd.DataFrame(columns=cols_s)

    df["MES"] = df[col_mes].fillna("").astype(str).str.strip().str.capitalize()
    df["AÑO"] = df[col_anio].fillna("").astype(str).str.strip()
    df = df[(df["MES"] == mes_str) & (df["AÑO"] == anio_str)].copy()

    if df.empty:
        print(f"  ⚠️  Consolidado_Impactos sin filas para {mes_str}-{anio_str}")
        return pd.DataFrame(columns=cols_g), pd.DataFrame(columns=cols_s)

    df["NOMBRE"]            = df[col_ases].apply(parsear_nombre_asesor)
    df["SUPERVISOR_LIDER"]  = df[col_sup].astype(str).str.strip().str.upper()
    df["Cuota"]             = _num_es(df[col_cuo])
    df["Venta"]             = _num_es(df[col_ven])
    df["Total Clientes Impactar"] = _num_es(df[col_plan])
    df["Clientes Impactados"]     = _num_es(df[col_real])

    # Excluir filas con NOMBRE vacío (defensivo)
    df = df[df["NOMBRE"].astype(str).str.len() > 0].copy()

    # ── Por gestor ────────────────────────────────────────────────────────
    gestor = (
        df.groupby(["NOMBRE", "SUPERVISOR_LIDER"], as_index=False)
          .agg(CUOTA              =("Cuota",                    "sum"),
               VENTA              =("Venta",                    "sum"),
               CLIENTES_IMPACTAR  =("Total Clientes Impactar",  "sum"),
               CLIENTES_IMPACTADOS=("Clientes Impactados",      "sum"))
    )
    gestor["VENTA_%"]    = gestor.apply(lambda r: _ratio(r["VENTA"], r["CUOTA"]), axis=1)
    gestor["IMPACTOS_%"] = gestor.apply(
        lambda r: _ratio(r["CLIENTES_IMPACTADOS"], r["CLIENTES_IMPACTAR"]), axis=1
    )
    gestor = gestor[cols_g]

    # ── Por supervisor (sum sobre el equipo) ──────────────────────────────
    supervisor = (
        df.groupby("SUPERVISOR_LIDER", as_index=False)
          .agg(CUOTA              =("Cuota",                    "sum"),
               VENTA              =("Venta",                    "sum"),
               CLIENTES_IMPACTAR  =("Total Clientes Impactar",  "sum"),
               CLIENTES_IMPACTADOS=("Clientes Impactados",      "sum"))
    )
    supervisor["VENTA_%"]    = supervisor.apply(lambda r: _ratio(r["VENTA"], r["CUOTA"]), axis=1)
    supervisor["IMPACTOS_%"] = supervisor.apply(
        lambda r: _ratio(r["CLIENTES_IMPACTADOS"], r["CLIENTES_IMPACTAR"]), axis=1
    )
    supervisor = supervisor[cols_s]

    return gestor, supervisor


# ─────────────────────────────────────────────────────────────────────────────
# 2) MSL_% y PROD_NUEVOS_%  (desde Consolidado_Ventas.csv + listas)
# ─────────────────────────────────────────────────────────────────────────────

def _cargar_ventas_periodo(mes: int, anio: int) -> pd.DataFrame:
    """
    Lee SALIDA/D&P/Consolidado_Ventas.csv filtrado al periodo dado.
    Devuelve un df normalizado con columnas:
        cod_cliente, ean, vendedor, supervisor, cant_total
    Sólo conserva filas con cantidad > 0.
    """
    mes_str  = MESES_INT_A_STR.get(int(mes), "")
    anio_str = str(int(anio))

    ruta_ventas = _ruta_ventas_cloud(mes, anio)
    # FIX: `usecols` con nombres exactos hace que pandas lance ValueError si
    # una sola columna cambia de capitalización o tilde (p.ej. "Año" vs "AÑO"),
    # y el except devolvía un DataFrame vacío → MSL_% y PROD_NUEVOS_% en NaN
    # para todos. Se lee sin usecols y se resuelven los nombres reales; el
    # archivo es grande, así que se limitan las columnas DESPUÉS de resolver.
    df = _leer_csv_cloud(
        ruta_ventas, "Consolidado_Ventas.csv",
        sep="|", decimal=",", encoding="utf-8",
        dtype=str, low_memory=False,
    )
    if df is None:
        return pd.DataFrame()

    col_mes  = _buscar_col(df, "MES", "Mes")
    col_anio = _buscar_col(df, "AÑO", "Año", "ANIO", "ANO")
    col_cli  = _buscar_col(df, "Cod. Cliente", "Cod Cliente", contiene=["COD", "CLIENTE"])
    col_ean  = _buscar_col(df, "Cod. EAN Producto", "Cod EAN Producto", contiene=["EAN"])
    col_vend = _buscar_col(df, "Vendedor")
    col_sup  = _buscar_col(df, "Supervisor")
    col_cant = _buscar_col(df, "Cantidades Totales", contiene=["CANTIDADES", "TOTALES"])

    faltan = [n for n, c in [("Mes", col_mes), ("Año", col_anio), ("Cod. Cliente", col_cli),
                             ("Cod. EAN Producto", col_ean), ("Vendedor", col_vend),
                             ("Supervisor", col_sup), ("Cantidades Totales", col_cant)]
              if c is None]
    if faltan:
        print(f"  ⚠️  {ruta_ventas} sin columnas esperadas: {faltan}")
        print(f"      Columnas encontradas: {list(df.columns)}")
        return pd.DataFrame()

    df = df.rename(columns={
        col_cli: "Cod. Cliente", col_ean: "Cod. EAN Producto",
        col_vend: "Vendedor", col_sup: "Supervisor",
        col_cant: "Cantidades Totales",
    })
    df["MES"] = df[col_mes].fillna("").astype(str).str.strip().str.capitalize()
    df["AÑO"] = df[col_anio].fillna("").astype(str).str.strip()
    df = df[(df["MES"] == mes_str) & (df["AÑO"] == anio_str)].copy()
    if df.empty:
        return df

    df["Cantidades Totales"] = _num_es(df["Cantidades Totales"])
    df["Cod. Cliente"]      = df["Cod. Cliente"].fillna("").astype(str).str.strip()
    df["Cod. EAN Producto"] = df["Cod. EAN Producto"].map(_ean_clean)
    df["Vendedor"]          = df["Vendedor"].fillna("").astype(str).str.strip().str.upper()
    df["Supervisor"]        = df["Supervisor"].fillna("").astype(str).str.strip().str.upper()

    df = df.rename(columns={
        "Cod. Cliente":        "cod_cliente",
        "Cod. EAN Producto":   "ean",
        "Vendedor":            "vendedor",
        "Supervisor":          "supervisor",
        "Cantidades Totales":  "cant_total",
    })
    return df


def _msl_por_grupo(df_periodo: pd.DataFrame, eans_msl: set, llave: str) -> pd.DataFrame:
    """
    Calcula MSL_% por valor de la columna `llave` (ej. 'vendedor' o
    'supervisor'). MSL_% = #PDVs con al menos un EAN MSL vendido / #PDVs activos.
    Devuelve df con columnas [llave, MSL_%, MSL_PDVS_OK, MSL_PDVS_TOTAL].
    """
    if df_periodo.empty:
        return pd.DataFrame(columns=[llave, "MSL_%", "MSL_PDVS_OK", "MSL_PDVS_TOTAL"])

    df_v = df_periodo[df_periodo["cant_total"] > 0]

    # Universo: TODOS los PDVs vistos por el grupo en el periodo (con cant>0)
    universo = (df_v.groupby(llave)["cod_cliente"]
                    .apply(lambda s: set(s.unique())).reset_index(name="pdvs_universo"))
    msl = (df_v[df_v["ean"].isin(eans_msl)]
                .groupby(llave)["cod_cliente"]
                .apply(lambda s: set(s.unique())).reset_index(name="pdvs_msl"))

    out = universo.merge(msl, on=llave, how="left")
    out["pdvs_msl"] = out["pdvs_msl"].apply(lambda x: x if isinstance(x, set) else set())
    out["MSL_PDVS_OK"]    = out["pdvs_msl"].apply(len)
    out["MSL_PDVS_TOTAL"] = out["pdvs_universo"].apply(len)
    out["MSL_%"] = out.apply(
        lambda r: (len(r["pdvs_msl"] & r["pdvs_universo"]) / len(r["pdvs_universo"]))
                  if len(r["pdvs_universo"]) > 0 else np.nan,
        axis=1,
    )
    return out[[llave, "MSL_%", "MSL_PDVS_OK", "MSL_PDVS_TOTAL"]]


def _prodnuevos_por_grupo(
    df_periodo: pd.DataFrame,
    listas: dict,
    llave: str,
) -> pd.DataFrame:
    """
    Calcula PROD_NUEVOS_% por valor de `llave`. Sprint 14.4: ahora con los
    4 segmentos V2 (Listerine_Sandia, DoyPack_J&J_Baby, Liserine_Kids, Visine).

    Para cada hoja del segmento:
       hoja_% = #PDVs target del grupo con al menos un EAN del segmento
                vendido / #PDVs target del grupo
    PROD_NUEVOS_% = avg(hoja_% disponibles).

    Devuelve df con [llave, <seg1>_%, <seg2>_%, ..., PROD_NUEVOS_%,
                     PROD_NUEVOS_PDVS_OK, PROD_NUEVOS_PDVS_TOTAL].
    """
    # Importar la lista canónica V2 (defina una sola vez en etl_impactos_segmentos)
    try:
        from etl_impactos_segmentos import SEGMENTOS_PRODNUEVOS as _SEGS
    except Exception:
        _SEGS = ["Listerine_Sandia", "DoyPack_J&J_Baby", "Liserine_Kids", "Visine"]

    pct_cols     = [f"{s}_%" for s in _SEGS]
    cols_finales = [llave] + pct_cols + [
        "PROD_NUEVOS_%", "PROD_NUEVOS_PDVS_OK", "PROD_NUEVOS_PDVS_TOTAL",
    ]
    if df_periodo.empty:
        return pd.DataFrame(columns=cols_finales)

    df_v = df_periodo[df_periodo["cant_total"] > 0]

    # Universo de PDVs por grupo (todos los activos en el periodo)
    universo = (df_v.groupby(llave)["cod_cliente"]
                    .apply(lambda s: set(s.unique())).reset_index(name="pdvs_universo"))

    acum = universo[[llave]].copy()
    acum["_ok_total"]     = 0
    acum["_target_total"] = 0

    for nombre_hoja in _SEGS:
        eans_hoja   = listas.get(nombre_hoja, {}).get("eans") or set()
        pdvs_target = listas.get(nombre_hoja, {}).get("pdvs") or set()

        u = universo.copy()
        u["pdvs_target"] = u["pdvs_universo"].apply(lambda s: s & pdvs_target)

        df_h = df_v[df_v["ean"].isin(eans_hoja)]
        pdvs_v = (df_h.groupby(llave)["cod_cliente"]
                       .apply(lambda s: set(s.unique())).reset_index(name="pdvs_vendieron"))

        u = u.merge(pdvs_v, on=llave, how="left")
        u["pdvs_vendieron"] = u["pdvs_vendieron"].apply(lambda x: x if isinstance(x, set) else set())
        u[f"{nombre_hoja}_OK"]     = u.apply(lambda r: len(r["pdvs_vendieron"] & r["pdvs_target"]), axis=1)
        u[f"{nombre_hoja}_TARGET"] = u["pdvs_target"].apply(len)
        u[f"{nombre_hoja}_%"]      = u.apply(
            lambda r: (r[f"{nombre_hoja}_OK"] / r[f"{nombre_hoja}_TARGET"])
                       if r[f"{nombre_hoja}_TARGET"] > 0 else np.nan,
            axis=1,
        )

        acum = acum.merge(
            u[[llave, f"{nombre_hoja}_%", f"{nombre_hoja}_OK", f"{nombre_hoja}_TARGET"]],
            on=llave, how="left",
        )
        acum["_ok_total"]     = acum["_ok_total"]     + acum[f"{nombre_hoja}_OK"].fillna(0)
        acum["_target_total"] = acum["_target_total"] + acum[f"{nombre_hoja}_TARGET"].fillna(0)

    acum["PROD_NUEVOS_%"] = acum[pct_cols].mean(axis=1, skipna=True)
    acum["PROD_NUEVOS_PDVS_OK"]    = acum["_ok_total"].astype(int)
    acum["PROD_NUEVOS_PDVS_TOTAL"] = acum["_target_total"].astype(int)

    return acum[cols_finales]


def cumplimientos_msl_prodnuevos(mes: int, anio: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Devuelve (df_gestor, df_supervisor). Si faltan inputs → frames vacíos.
    Sprint 14.4: usa los 4 segmentos V2 importados desde etl_impactos_segmentos.
    """
    try:
        from etl_impactos_segmentos import SEGMENTOS_PRODNUEVOS as _SEGS
    except Exception:
        _SEGS = ["Listerine_Sandia", "DoyPack_J&J_Baby", "Liserine_Kids", "Visine"]
    pct_seg_cols = [f"{s}_%" for s in _SEGS]
    cols_g = ["NOMBRE", "MSL_%", "MSL_PDVS_OK", "MSL_PDVS_TOTAL"] \
             + pct_seg_cols \
             + ["PROD_NUEVOS_%", "PROD_NUEVOS_PDVS_OK", "PROD_NUEVOS_PDVS_TOTAL"]
    cols_s = ["SUPERVISOR_LIDER"] + cols_g[1:]

    ruta_listas_temp = _ruta_listas_temp_cloud()
    if ruta_listas_temp is None:
        print(f"  ⚠️  Falta Listas D&P en SharePoint ({RUTA_LISTAS_CLOUD}) — MSL_% y PROD_NUEVOS_% serán NaN")
        return pd.DataFrame(columns=cols_g), pd.DataFrame(columns=cols_s)

    print(f"  Cargando ventas D&P para {MESES_INT_A_STR.get(int(mes),'?')}-{anio}...")
    df_periodo = _cargar_ventas_periodo(mes, anio)
    if df_periodo.empty:
        print(f"  ⚠️  Consolidado_Ventas sin filas para {MESES_INT_A_STR.get(int(mes),'?')}-{anio} "
              f"(o no se encontró en SharePoint: {_ruta_ventas_cloud(mes, anio)})")
        return pd.DataFrame(columns=cols_g), pd.DataFrame(columns=cols_s)

    from etl_impactos_segmentos import cargar_listas
    listas = cargar_listas(ruta_listas_temp)
    eans_msl = listas.get("MustStock", {}).get("eans") or set()

    # ── Por gestor (vendedor) ─────────────────────────────────────────────
    msl_g  = _msl_por_grupo(df_periodo, eans_msl, "vendedor")
    prod_g = _prodnuevos_por_grupo(df_periodo, listas, "vendedor")
    gestor = msl_g.merge(prod_g, on="vendedor", how="outer")
    gestor = gestor.rename(columns={"vendedor": "NOMBRE"})
    for c in cols_g:
        if c not in gestor.columns:
            gestor[c] = np.nan
    gestor = gestor[cols_g]

    # ── Por supervisor ────────────────────────────────────────────────────
    msl_s  = _msl_por_grupo(df_periodo, eans_msl, "supervisor")
    prod_s = _prodnuevos_por_grupo(df_periodo, listas, "supervisor")
    supervisor = msl_s.merge(prod_s, on="supervisor", how="outer")
    supervisor = supervisor.rename(columns={"supervisor": "SUPERVISOR_LIDER"})
    for c in cols_s:
        if c not in supervisor.columns:
            supervisor[c] = np.nan
    supervisor = supervisor[cols_s]

    return gestor, supervisor


# ─────────────────────────────────────────────────────────────────────────────
# 3) FUNCIÓN COMBINADA (la usa calcular_cumplimientos.main)
# ─────────────────────────────────────────────────────────────────────────────

def calcular_kpis_dyp(
    mes: int,
    anio: int,
    nombres_pt_gestor: set | list | None = None,
    nombres_pt_supervisor: set | list | None = None,
    base_cupos_idx: dict | None = None,
    nombres_supervisores_bc: set | list | None = None,
) -> dict:
    """
    Calcula los 4 KPIs comerciales y devuelve dict {'gestor', 'supervisor'}.

    Modos
    ─────
    • Modo "ACRONIMO" (recomendado): pasar `base_cupos_idx` (output de
      `base_cupos.construir_indices`). Cruza el COD_ASESOR_ECOM del CSV
      con la maestra para añadir las columnas:
          ACRONIMO          → llave canónica de la persona
          ACRONIMO_SUP      → llave canónica del supervisor (resuelto desde
                              el campo Supervisor del CSV, armonizado)
      `nombres_supervisores_bc` debe ser el set de nombres canónicos de
      los supervisores de Base cupos (para armonizar el `Supervisor` del
      CSV antes de mapear a ACRONIMO_SUP).

    • Modo legacy (sin base_cupos_idx): se conserva el comportamiento
      anterior — armoniza nombres contra el PT y devuelve `NOMBRE` /
      `SUPERVISOR_LIDER`. Útil si la maestra no está disponible.
    """
    print(f"\nCalculando KPIs D&P (Venta + Impactos + MSL + Prod Nuevos):")
    g_vi, s_vi = cumplimientos_venta_impactos(mes, anio)
    g_mp, s_mp = cumplimientos_msl_prodnuevos(mes, anio)

    # Merge gestor: por NOMBRE
    if g_vi.empty and g_mp.empty:
        gestor = pd.DataFrame(columns=["NOMBRE", "SUPERVISOR_LIDER",
                                       "VENTA_%", "IMPACTOS_%",
                                       "MSL_%", "PROD_NUEVOS_%"])
    else:
        gestor = (g_vi if not g_vi.empty else pd.DataFrame(columns=["NOMBRE"])) \
                    .merge(g_mp if not g_mp.empty else pd.DataFrame(columns=["NOMBRE"]),
                           on="NOMBRE", how="outer")

    # Merge supervisor: por SUPERVISOR_LIDER
    if s_vi.empty and s_mp.empty:
        supervisor = pd.DataFrame(columns=["SUPERVISOR_LIDER",
                                           "VENTA_%", "IMPACTOS_%",
                                           "MSL_%", "PROD_NUEVOS_%"])
    else:
        supervisor = (s_vi if not s_vi.empty else pd.DataFrame(columns=["SUPERVISOR_LIDER"])) \
                        .merge(s_mp if not s_mp.empty else pd.DataFrame(columns=["SUPERVISOR_LIDER"]),
                               on="SUPERVISOR_LIDER", how="outer")

    # Asegurar que las 4 columnas % existan, aunque alguna fuente no haya
    # producido datos (evita KeyError aguas abajo).
    for c in ("VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"):
        if c not in gestor.columns:
            gestor[c] = np.nan
        if c not in supervisor.columns:
            supervisor[c] = np.nan

    # Armonización de nombres con el PT (nombres truncados → completos)
    if nombres_pt_gestor:
        gestor = armonizar_nombres_a_pt(gestor, "NOMBRE", nombres_pt_gestor)
    if nombres_pt_supervisor:
        supervisor = armonizar_nombres_a_pt(supervisor, "SUPERVISOR_LIDER", nombres_pt_supervisor)

    # ── Modo ACRONIMO: enriquecer con llaves canónicas vía Base cupos ────
    if base_cupos_idx is not None:
        gestor, supervisor = _enriquecer_con_acronimo(
            gestor, supervisor, mes, anio, base_cupos_idx,
            nombres_supervisores_bc=nombres_supervisores_bc,
        )
    # Si después de armonizar quedan duplicados (ej. dos filas D&P apuntando
    # al mismo nombre PT por colisión), agregamos sumando métricas y
    # promediando %s. Es defensivo — en la práctica no debería ocurrir.
    if not gestor.empty and gestor["NOMBRE"].duplicated().any():
        gestor = gestor.groupby("NOMBRE", as_index=False, dropna=False).first()
    if not supervisor.empty and supervisor["SUPERVISOR_LIDER"].duplicated().any():
        supervisor = supervisor.groupby("SUPERVISOR_LIDER", as_index=False, dropna=False).first()

    n_g = (gestor[["VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]]
                  .notna().any(axis=1).sum()) if not gestor.empty else 0
    n_s = (supervisor[["VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]]
                      .notna().any(axis=1).sum()) if not supervisor.empty else 0
    print(f"  ✓ D&P — gestores con datos: {n_g} | supervisores: {n_s}")

    return {"gestor": gestor, "supervisor": supervisor}