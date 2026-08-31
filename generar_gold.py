"""
generar_gold.py — Sprint 16.3 (Cloud)
──────────────────────────────────────
Genera la capa GOLD en CSV (UTF-8, separador coma, decimal punto) lista para
consumo de BI (Power BI, otros). Esquema estrella. Todas las fuentes se leen
desde SharePoint (Graph API) y los CSV finales se suben a SharePoint —
ya no hay lectura/escritura en disco local.

  Facts agregadas:
    fact_cumplimientos_gestor.csv   1 fila por (persona, periodo) con TOTALes

  Facts detalle por PDV (multi-periodo):
    fact_cif_pdv.csv                CIF.xlsx tal cual
    fact_np_reporte.csv             REPORTE_NO_PRESENCIA_{MES}_{AÑO}.xlsx
    fact_np_agotados.csv            ANALISIS_AGOTADOS_{MES}_{AÑO}.xlsx
    fact_precios_captura.csv        REPORTE_CAPTURA_PRECIOS_{MES}_{AÑO}.xlsx
    fact_precios_analisis.csv       ANALISIS_PRECIOS_{MES}_{AÑO}.xlsx
    fact_sos_pdv.csv                Reporte_SOS_Final_Calculado_{MES}_{AÑO}.xlsx
    fact_exh_pagadas.csv            Resultado_exhibiciones_pagadas {Mes} {Año}.xlsx
    fact_exh_gratis.csv             Resultado exhibiciones gratis.xlsx

  Dimensiones:
    dim_personas.csv                Base cupos
    dim_pdv.csv                     únicos de CIF.xlsx
    dim_periodo.csv                 derivado de los HISTORICOs

CLI:
    python generar_gold.py                       # genera todo
    python generar_gold.py --periodos 04/2026,05/2026   # filtrar a esos meses

Output (SharePoint): {_SALIDAS_ROOT}/GOLD/*.csv

NOTA sobre nombres de archivo en la nube: varios difieren de la convención
local original de paths.py (ej. SOS ahora trae el periodo en el nombre;
el histórico de Exhibiciones Gratis se llama "KPI_Exhibiciones_Gratis_Historico.xlsx",
no "EXHIBICIONES_GRATIS_KPIS_HISTORICO.xlsx"). Los nombres usados aquí están
tomados de logs reales de ejecución — si algo no calza, el mensaje de
"no encontrado" te dice la ruta exacta que se intentó, para poder ajustarla.
"""
from __future__ import annotations

import io
import re
import sys
import argparse
import urllib.parse
from pathlib import Path
from typing import Iterable

import requests
import pandas as pd

import paths
import periodo_resolver as pr


RUTA_GOLD_CLOUD = f"{paths._SALIDAS_ROOT}/GOLD"

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
MESES_ES_INV = {v: k for k, v in MESES_ES.items()}
MESES_VARIANTES = {**MESES_ES_INV, "MARZO": 3, "MAYO": 5}


def _norm_text(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)


def _parse_periodo_filename(nombre: str) -> tuple[int, int]:
    """Extrae (mes, anio) de un filename. Ej. 'ANALISIS_PRECIOS_MAYO_2026.xlsx' → (5, 2026)."""
    base = nombre.upper().replace(".XLSX", "")
    m = re.search(r"([A-ZÁÉÍÓÚÑ]+)[_\s-]+(\d{4})", base)
    if m:
        mes_str, anio_str = m.group(1), m.group(2)
        mes = MESES_VARIANTES.get(mes_str.replace("Á", "A").replace("Ñ", "N"), 0)
        if mes:
            return mes, int(anio_str)
    return 0, 0


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
    raise FileNotFoundError(f"No se pudo descargar {descripcion} desde SharePoint ({ruta_sharepoint}): {status}")


def _subir_bytes_cloud(contenido: bytes, ruta_sharepoint: str, content_type: str) -> None:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": content_type}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(ruta_limpia)}:/content"
    response = requests.put(url, headers=headers, data=contenido)
    if response.status_code not in (200, 201):
        raise Exception(f"Error al subir '{ruta_sharepoint}' a SharePoint: {response.status_code} - {response.text}")


def _leer_excel_cloud_opcional(ruta_sharepoint: str, descripcion: str, **kwargs) -> pd.DataFrame:
    """Lee un Excel de SharePoint; devuelve DataFrame vacío (sin tronar) si no existe."""
    try:
        contenido = _descargar_bytes_cloud(ruta_sharepoint, descripcion)
    except Exception as e:
        print(f"  ⚠  {descripcion} no encontrado en SharePoint ({ruta_sharepoint}): {e}")
        return pd.DataFrame()
    try:
        return pd.read_excel(io.BytesIO(contenido), **kwargs)
    except Exception as e:
        print(f"  ⚠  {descripcion}: error al leer ({e})")
        return pd.DataFrame()


def _leer_multi_periodo_cloud(
    ruta_carpeta: str,
    patron_regex: str,
    *,
    sheet_name=0,
    filtro_periodos: list[pr.PeriodoSpec] | None = None,
) -> pd.DataFrame:
    """
    Equivalente cloud de _leer_multi_periodo: lista la carpeta en SharePoint,
    descarga y concatena todos los archivos que matcheen patron_regex.
    Si los archivos no traen MES/AÑO en columnas, los infiere del filename.
    """
    try:
        hijos = _listar_hijos_cloud(ruta_carpeta)
    except Exception as e:
        print(f"  ⚠  No se pudo listar '{ruta_carpeta}' en SharePoint: {e}")
        return pd.DataFrame()

    archivos = sorted(
        h["name"] for h in hijos
        if h.get("file") and re.match(patron_regex, h.get("name", "")) and not h["name"].startswith("~$")
    )
    if not archivos:
        return pd.DataFrame()

    periodos_set = (
        {(s.mes, s.anio) for s in filtro_periodos}
        if filtro_periodos else None
    )

    dfs = []
    for nombre in archivos:
        mes, anio = _parse_periodo_filename(nombre)
        if periodos_set is not None and mes and anio and (mes, anio) not in periodos_set:
            continue
        try:
            contenido = _descargar_bytes_cloud(f"{ruta_carpeta}/{nombre}", nombre)
            df = pd.read_excel(io.BytesIO(contenido), sheet_name=sheet_name)
        except Exception as e:
            print(f"⚠️  {nombre}: {e} — saltado")
            continue
        if isinstance(df, dict):
            df = next(iter(df.values()))
        if df.empty:
            continue
        cols_upper = {c.upper().replace("Ñ", "N"): c for c in df.columns}
        col_mes = next((cols_upper[k] for k in cols_upper if k in ("MES", "MES_")), None)
        col_anio = next((cols_upper[k] for k in cols_upper if k in ("ANO", "ANIO", "AÑO")), None)
        if col_mes is None and mes:
            df["MES"] = mes
        if col_anio is None and anio:
            df["AÑO"] = anio
        dfs.append(df)

    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True, sort=False)


def _filtrar_por_periodos(df: pd.DataFrame, filtro_periodos: list[pr.PeriodoSpec] | None) -> pd.DataFrame:
    if filtro_periodos is None or df.empty:
        return df
    periodos_set = {(s.mes, s.anio) for s in filtro_periodos}
    col_mes = next((c for c in df.columns if str(c).upper() == "MES"), None)
    col_anio = next((c for c in df.columns if str(c).upper().replace("Ñ", "N") in ("ANO", "AÑO", "ANIO")), None)
    if not (col_mes and col_anio):
        return df
    mask = df.apply(
        lambda r: (int(r[col_mes]), int(r[col_anio])) in periodos_set
        if pd.notna(r[col_mes]) and pd.notna(r[col_anio]) else False,
        axis=1,
    )
    return df[mask].copy()


def _to_csv_cloud(df: pd.DataFrame, nombre: str) -> str:
    ruta_cloud = f"{RUTA_GOLD_CLOUD}/{nombre}"
    buffer = io.StringIO()
    df.to_csv(buffer, index=False, encoding="utf-8", sep=",", decimal=".")
    _subir_bytes_cloud(buffer.getvalue().encode("utf-8"), ruta_cloud, "text/csv; charset=utf-8")
    return ruta_cloud


# ─────────────────────────────────────────────────────────────────────────────
# FACTS DETALLE POR PDV
# ─────────────────────────────────────────────────────────────────────────────

def fact_cif_pdv(filtro):
    ruta = f"{paths.RUTA_CARPETA_SALIDAS_CIF}/CIF.xlsx"
    df = _leer_excel_cloud_opcional(ruta, "CIF.xlsx")
    return _filtrar_por_periodos(df, filtro)


def fact_np_reporte(filtro):
    return _leer_multi_periodo_cloud(
        paths.RUTA_CARPETA_SALIDAS_NP, r"^REPORTE_NO_PRESENCIA_.*\.xlsx$", filtro_periodos=filtro
    )


def fact_np_agotados(filtro):
    return _leer_multi_periodo_cloud(
        paths.RUTA_CARPETA_SALIDAS_NP, r"^ANALISIS_AGOTADOS_.*\.xlsx$", filtro_periodos=filtro
    )


def fact_precios_captura(filtro):
    return _leer_multi_periodo_cloud(
        paths.RUTA_CARPETA_SALIDAS_PRECIOS, r"^REPORTE_CAPTURA_PRECIOS_.*\.xlsx$", filtro_periodos=filtro
    )


def fact_precios_analisis(filtro):
    return _leer_multi_periodo_cloud(
        paths.RUTA_CARPETA_SALIDAS_PRECIOS, r"^ANALISIS_PRECIOS_.*\.xlsx$", filtro_periodos=filtro
    )


def fact_sos_pdv(filtro):
    # En la nube el reporte final de SOS trae el periodo en el nombre
    # (Reporte_SOS_Final_Calculado_JULIO_2026.xlsx), a diferencia del
    # archivo único local original — por eso usa el lector multi-periodo.
    return _leer_multi_periodo_cloud(
        paths.RUTA_CARPETA_SALIDAS_SOS, r"^Reporte_SOS_Final_Calculado_.*\.xlsx$", filtro_periodos=filtro
    )


def fact_exh_pagadas(filtro):
    return _leer_multi_periodo_cloud(
        paths.RUTA_CARPETA_SALIDAS_EXHIB, r"^Resultado_exhibiciones_pagadas .*\.xlsx$", filtro_periodos=filtro
    )


def fact_exh_gratis(filtro):
    ruta = f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/Resultado exhibiciones gratis.xlsx"
    df = _leer_excel_cloud_opcional(ruta, "Resultado exhibiciones gratis.xlsx", sheet_name="Exhibiciones_implementadas")
    if df.empty:
        df = _leer_excel_cloud_opcional(ruta, "Resultado exhibiciones gratis.xlsx")
    if df.empty:
        return df
    df = df.rename(columns={
        "Mes": "MES", "Año": "AÑO", "Mes-Año": "PERIODO_ETIQUETA",
        "ID PDV": "ID_PDV_INVOLVES", "Empleado": "NOMBRE",
        "Rol Empleado": "ROL",
    })
    return _filtrar_por_periodos(df, filtro)


# ─────────────────────────────────────────────────────────────────────────────
# FACT AGREGADO POR GESTOR
# ─────────────────────────────────────────────────────────────────────────────

def _leer_historico_normalizado_cloud(ruta_cloud: str, descripcion: str, kpi_col: str, kpi_nombre: str) -> pd.DataFrame:
    """
    Lee un HISTORICO desde SharePoint, normaliza columnas a
    (MES, AÑO, NOMBRE_N, SUPERVISOR_LIDER, <kpi_nombre>).
    """
    cols_vacias = ["MES", "AÑO", "NOMBRE_N", "SUPERVISOR_LIDER", kpi_nombre]
    df = _leer_excel_cloud_opcional(ruta_cloud, descripcion)
    if df.empty:
        return pd.DataFrame(columns=cols_vacias)

    rename_map = {}
    for c in df.columns:
        cu = str(c).strip().upper().replace("Ñ", "N")
        if cu == "MES":               rename_map[c] = "MES"
        elif cu in ("ANO", "ANIO"):   rename_map[c] = "AÑO"
        elif cu == "NOMBRE":          rename_map[c] = "NOMBRE"
        elif cu == "EMPLEADO":        rename_map[c] = "NOMBRE"
        elif cu == "SUPERVISOR_LIDER": rename_map[c] = "SUPERVISOR_LIDER"
    df = df.rename(columns=rename_map)
    if kpi_col not in df.columns or "NOMBRE" not in df.columns:
        return pd.DataFrame(columns=cols_vacias)
    df["NOMBRE_N"] = _norm_text(df["NOMBRE"])
    cols = ["MES", "AÑO", "NOMBRE_N"]
    if "SUPERVISOR_LIDER" in df.columns:
        cols.append("SUPERVISOR_LIDER")
    cols.append(kpi_col)
    out = df[cols].rename(columns={kpi_col: kpi_nombre}).copy()
    out["MES"] = pd.to_numeric(out["MES"], errors="coerce").astype("Int64")
    out["AÑO"] = pd.to_numeric(out["AÑO"], errors="coerce").astype("Int64")
    return out


# Rutas de los HISTORICO en SharePoint. Confirmadas por log real de
# ejecución donde fue posible; el de Exhibiciones Gratis usa el nombre
# real observado ("KPI_Exhibiciones_Gratis_Historico.xlsx"), que NO sigue
# la convención de paths.py local.
_RUTA_CIF_HIST  = f"{paths.RUTA_CARPETA_SALIDAS_CIF}/{paths.CIF_OUT_KPIS_HISTORICO.name}"
_RUTA_NP_HIST   = f"{paths.RUTA_CARPETA_SALIDAS_NP}/{paths.NP_OUT_KPIS_HISTORICO.name}"
_RUTA_PR_HIST   = f"{paths.RUTA_CARPETA_SALIDAS_PRECIOS}/{paths.PR_OUT_KPIS_HISTORICO.name}"
_RUTA_SOS_HIST  = f"{paths.RUTA_CARPETA_SALIDAS_SOS}/{paths.SOS_OUT_KPIS_HISTORICO.name}"
_RUTA_EPG_HIST  = f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/{paths.EXHIB_PAG_OUT_KPIS_HISTORICO.name}"
_RUTA_EGR_HIST  = f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/KPI_Exhibiciones_Gratis_Historico.xlsx"


def fact_cumplimientos_gestor(filtro):
    cif = _leer_historico_normalizado_cloud(_RUTA_CIF_HIST, "CIF histórico", "TOTAL", "CUMP_CIF")
    np_ = _leer_historico_normalizado_cloud(_RUTA_NP_HIST, "No Presencia histórico", "EJECUCION", "CUMP_NP")
    pre = _leer_historico_normalizado_cloud(_RUTA_PR_HIST, "Precios histórico", "CUMPLIMIENTO", "CUMP_PRECIOS")
    sos = _leer_historico_normalizado_cloud(_RUTA_SOS_HIST, "SOS histórico", "CUMPLIMIENTO", "CUMP_SOS")
    epg = _leer_historico_normalizado_cloud(_RUTA_EPG_HIST, "Exh. Pagadas histórico", "CUMPLIMIENTO", "CUMP_EXH_PAG")
    egr = _leer_historico_normalizado_cloud(_RUTA_EGR_HIST, "Exh. Gratis histórico", "TOTAL", "CUMP_EXH_GRA")

    keys = ["MES", "AÑO", "NOMBRE_N"]

    def _una_fila_por_persona(df, etiqueta):
        """
        Deja una sola fila por (mes, año, persona).

        Sin esto, un histórico con la persona repetida en el mismo periodo
        (p. ej. si cambió de supervisor a mitad de mes) hace que el merge
        multiplique filas: 2 duplicados en dos fuentes distintas producen 4
        filas en la salida. No falla, solo infla el CSV en silencio, y los
        promedios de Power BI salen mal.
        """
        if df.empty or not all(k in df.columns for k in keys):
            return df
        n_antes = len(df)
        df = df.drop_duplicates(subset=keys, keep="last")
        if len(df) != n_antes:
            print(f"    ⚠ {etiqueta}: {n_antes - len(df)} fila(s) duplicada(s) "
                  f"por persona/periodo — se conserva la última")
        return df

    base = _una_fila_por_persona(cif, "CIF histórico")
    for parte, etiqueta in ((np_, "NP"), (pre, "Precios"), (sos, "SOS"),
                            (epg, "Exh. Pagadas"), (egr, "Exh. Gratis")):
        if parte.empty:
            continue
        sup_cols = [c for c in parte.columns if c == "SUPERVISOR_LIDER"]
        parte_solo_kpi = _una_fila_por_persona(
            parte.drop(columns=sup_cols), f"{etiqueta} histórico")
        base = base.merge(parte_solo_kpi, on=keys, how="outer")

    return _filtrar_por_periodos(base, filtro).sort_values(
        ["AÑO", "MES", "NOMBRE_N"]
    ).reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# DIMENSIONES
# ─────────────────────────────────────────────────────────────────────────────

def dim_personas() -> pd.DataFrame:
    ruta = f"{paths.RUTA_CARPETA_BASES_DYP}/Base_cupos.xlsx"
    df = _leer_excel_cloud_opcional(ruta, "Base_cupos.xlsx", sheet_name="Tabla total roles")
    if df.empty:
        return df
    df["NOMBRE_N"] = _norm_text(df["NOMBRE"])
    return df.drop_duplicates("NOMBRE_N").reset_index(drop=True)


def dim_pdv() -> pd.DataFrame:
    ruta = f"{paths.RUTA_CARPETA_SALIDAS_CIF}/CIF.xlsx"
    df = _leer_excel_cloud_opcional(ruta, "CIF.xlsx")
    if df.empty:
        return df
    cols = [c for c in [
        "ID_PDV_INVOLVES", "NOMBRE_PDV", "ACRONIMO",
        "VENTAS_PROMEDIO_MES", "FUENTE",
    ] if c in df.columns]
    return df[cols].drop_duplicates("ID_PDV_INVOLVES").reset_index(drop=True)


def dim_periodo(filtro) -> pd.DataFrame:
    periodos: set[tuple[int, int]] = set()
    historicos = [_RUTA_CIF_HIST, _RUTA_NP_HIST, _RUTA_PR_HIST, _RUTA_SOS_HIST, _RUTA_EPG_HIST, _RUTA_EGR_HIST]
    for ruta in historicos:
        df = _leer_excel_cloud_opcional(ruta, Path(ruta).name)
        if df.empty:
            continue
        cols_upper = {c.upper().replace("Ñ", "N"): c for c in df.columns}
        col_mes = next((cols_upper[k] for k in cols_upper if k == "MES"), None)
        col_anio = next((cols_upper[k] for k in cols_upper if k in ("ANO", "AÑO", "ANIO")), None)
        if col_mes and col_anio:
            for _, r in df[[col_mes, col_anio]].drop_duplicates().iterrows():
                try:
                    periodos.add((int(r[col_mes]), int(r[col_anio])))
                except (ValueError, TypeError):
                    pass

    if filtro is not None:
        periodos &= {(s.mes, s.anio) for s in filtro}

    rows = []
    for mes, anio in sorted(periodos, key=lambda x: (x[1], x[0])):
        rows.append({
            "MES": mes,
            "AÑO": anio,
            "PERIODO_ETIQUETA": f"{MESES_ES[mes].title()}_{anio}",
            "PERIODO_ID": f"{anio}{mes:02d}",
            "FECHA_INICIO_MES": f"{anio}-{mes:02d}-01",
        })
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

TABLAS = [
    ("fact_cumplimientos_gestor", fact_cumplimientos_gestor, True),
    ("fact_cif_pdv",               fact_cif_pdv,              True),
    ("fact_np_reporte",            fact_np_reporte,           True),
    ("fact_np_agotados",           fact_np_agotados,          True),
    ("fact_precios_captura",       fact_precios_captura,      True),
    ("fact_precios_analisis",      fact_precios_analisis,     True),
    ("fact_sos_pdv",               fact_sos_pdv,              True),
    ("fact_exh_pagadas",           fact_exh_pagadas,          True),
    ("fact_exh_gratis",            fact_exh_gratis,           True),
    ("dim_personas",               lambda f: dim_personas(),  False),
    ("dim_pdv",                    lambda f: dim_pdv(),       False),
    ("dim_periodo",                dim_periodo,               True),
]


def generar_gold(filtro_periodos: list[pr.PeriodoSpec] | None = None) -> dict[str, int]:
    resumen: dict[str, int] = {}
    print(f"☁️  Output: {RUTA_GOLD_CLOUD} (SharePoint)")
    if filtro_periodos:
        print(f"🔍 Filtrando a periodos: {', '.join(s.etiqueta for s in filtro_periodos)}")
    print()
    for nombre, fn, usa_filtro in TABLAS:
        try:
            df = fn(filtro_periodos) if usa_filtro else fn(None)
        except Exception as e:
            print(f"  ✗ {nombre:<30} FAIL: {e}")
            resumen[nombre] = -1
            continue
        if df is None or df.empty:
            print(f"  ⚠ {nombre:<30} (vacío) — no se sube")
            resumen[nombre] = 0
            continue
        ruta_cloud = _to_csv_cloud(df, f"{nombre}.csv")
        print(f"  ✓ {nombre:<30} {len(df):>7,} filas → {ruta_cloud}")
        resumen[nombre] = len(df)
    return resumen


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        description="Generar capa GOLD (CSV para BI) — Sprint 16.3 Cloud",
    )
    parser.add_argument(
        "--periodos", type=str, default=None,
        help="Filtrar a periodos específicos. Formato MM/AAAA[,MM/AAAA,...]. "
             "Sin esto: genera todo el histórico disponible.",
    )
    args = parser.parse_args()

    filtro = None
    if args.periodos:
        tokens = [t.strip() for t in args.periodos.split(",") if t.strip()]
        filtro = [pr.parsear_periodo_str(t) for t in tokens]

    print("=" * 60)
    print(" GENERACIÓN CAPA GOLD — Eficacia (Sprint 16.3, Cloud)")
    print("=" * 60)
    resumen = generar_gold(filtro)
    print()
    total_filas = sum(v for v in resumen.values() if v > 0)
    print(f"🏁 Listo. {len([v for v in resumen.values() if v > 0])} tablas escritas, {total_filas:,} filas totales.")


if __name__ == "__main__":
    main()