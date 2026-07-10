"""
generar_gold.py — Sprint 16.3
─────────────────────────────
Genera la capa GOLD en CSV (UTF-8, separador coma, decimal punto) lista para
consumo de BI (Power BI, otros). Esquema estrella:

  Facts agregadas:
    fact_cumplimientos_gestor.csv   1 fila por (persona, periodo) con TOTALes

  Facts detalle por PDV (multi-periodo):
    fact_cif_pdv.csv                CIF.xlsx tal cual
    fact_np_reporte.csv             REPORTE_NO_PRESENCIA_{MES}_{AÑO}.xlsx
    fact_np_agotados.csv            ANALISIS_AGOTADOS_{MES}_{AÑO}.xlsx
    fact_precios_captura.csv        REPORTE_CAPTURA_PRECIOS_{MES}_{AÑO}.xlsx
    fact_precios_analisis.csv       ANALISIS_PRECIOS_{MES}_{AÑO}.xlsx
    fact_sos_pdv.csv                Reporte_SOS_Final_Calculado.xlsx
    fact_exh_pagadas.csv            Resultado_exhibiciones_pagadas {Mes} {Año}.xlsx
    fact_exh_gratis.csv             Resultado exhibiciones gratis.xlsx

  Dimensiones:
    dim_personas.csv                Base cupos
    dim_pdv.csv                     únicos de CIF.xlsx
    dim_periodo.csv                 derivado de los HISTORICOs

CLI:
    python generar_gold.py                       # genera todo
    python generar_gold.py --periodos 04/2026,05/2026   # filtrar a esos meses

Output: SALIDA/GOLD/*.csv
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Iterable

import pandas as pd

import paths
import periodo_resolver as pr


GOLD_DIR = paths.BASE / "SALIDA" / "GOLD"

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
MESES_ES_INV = {v: k for k, v in MESES_ES.items()}
# Variantes con tilde
MESES_VARIANTES = {**MESES_ES_INV, "MARZO": 3, "MAYO": 5}


def _norm_text(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)


def _parse_periodo_filename(nombre: str) -> tuple[int, int]:
    """Extrae (mes, anio) de un filename. Ej. 'ANALISIS_PRECIOS_MAYO_2026.xlsx' → (5, 2026)."""
    base = nombre.upper().replace(".XLSX", "")
    # Mes en español + año
    m = re.search(r"([A-ZÁÉÍÓÚÑ]+)[_\s-]+(\d{4})", base)
    if m:
        mes_str, anio_str = m.group(1), m.group(2)
        mes = MESES_VARIANTES.get(mes_str.replace("Á", "A").replace("Ñ", "N"), 0)
        if mes:
            return mes, int(anio_str)
    return 0, 0


def _leer_multi_periodo(
    carpeta: Path,
    patron: str,
    *,
    sheet_name=0,
    filtro_periodos: list[pr.PeriodoSpec] | None = None,
) -> pd.DataFrame:
    """
    Lee y concatena todos los archivos del patrón en carpeta.
    Si los archivos no traen MES/AÑO en columnas, los infiere del filename.
    Si `filtro_periodos` está dado, solo incluye archivos de esos periodos.
    """
    if not carpeta.is_dir():
        return pd.DataFrame()
    archivos = sorted(carpeta.glob(patron))
    if not archivos:
        return pd.DataFrame()

    periodos_set = (
        {(s.mes, s.anio) for s in filtro_periodos}
        if filtro_periodos else None
    )

    dfs = []
    for ruta in archivos:
        mes, anio = _parse_periodo_filename(ruta.name)
        if periodos_set is not None and mes and anio and (mes, anio) not in periodos_set:
            continue
        try:
            df = pd.read_excel(ruta, sheet_name=sheet_name)
        except Exception as e:
            print(f"⚠️  {ruta.name}: {e} — saltado")
            continue
        if isinstance(df, dict):
            df = next(iter(df.values()))
        if df.empty:
            continue
        # Inferir MES/AÑO si no vienen
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


def _to_csv(df: pd.DataFrame, nombre: str) -> Path:
    ruta = GOLD_DIR / nombre
    df.to_csv(ruta, index=False, encoding="utf-8", sep=",", decimal=".")
    return ruta


# ─────────────────────────────────────────────────────────────────────────────
# FACTS DETALLE POR PDV
# ─────────────────────────────────────────────────────────────────────────────

def fact_cif_pdv(filtro):
    if not paths.CIF_OUT_CIF.is_file():
        return pd.DataFrame()
    df = pd.read_excel(paths.CIF_OUT_CIF)
    return _filtrar_por_periodos(df, filtro)


def fact_np_reporte(filtro):
    return _leer_multi_periodo(paths.NP_SALIDA, "REPORTE_NO_PRESENCIA_*.xlsx", filtro_periodos=filtro)


def fact_np_agotados(filtro):
    return _leer_multi_periodo(paths.NP_SALIDA, "ANALISIS_AGOTADOS_*.xlsx", filtro_periodos=filtro)


def fact_precios_captura(filtro):
    return _leer_multi_periodo(paths.PR_SALIDA, "REPORTE_CAPTURA_PRECIOS_*.xlsx", filtro_periodos=filtro)


def fact_precios_analisis(filtro):
    return _leer_multi_periodo(paths.PR_SALIDA, "ANALISIS_PRECIOS_*.xlsx", filtro_periodos=filtro)


def fact_sos_pdv(filtro):
    ruta = paths.SOS_SALIDA / "Reporte_SOS_Final_Calculado.xlsx"
    if not ruta.is_file():
        return pd.DataFrame()
    df = pd.read_excel(ruta)
    return _filtrar_por_periodos(df, filtro)


def fact_exh_pagadas(filtro):
    return _leer_multi_periodo(paths.EXHIB_SALIDA, "Resultado_exhibiciones_pagadas *.xlsx", filtro_periodos=filtro)


def fact_exh_gratis(filtro):
    ruta = paths.EXHIB_SALIDA / "Resultado exhibiciones gratis.xlsx"
    if not ruta.is_file():
        return pd.DataFrame()
    try:
        df = pd.read_excel(ruta, sheet_name="Exhibiciones_implementadas")
    except Exception:
        df = pd.read_excel(ruta)
    # Normalizar nombres
    df = df.rename(columns={
        "Mes": "MES", "Año": "AÑO", "Mes-Año": "PERIODO_ETIQUETA",
        "ID PDV": "ID_PDV_INVOLVES", "Empleado": "NOMBRE",
        "Rol Empleado": "ROL",
    })
    return _filtrar_por_periodos(df, filtro)


# ─────────────────────────────────────────────────────────────────────────────
# FACT AGREGADO POR GESTOR
# ─────────────────────────────────────────────────────────────────────────────

def _leer_historico_normalizado(ruta: Path, kpi_col: str, kpi_nombre: str) -> pd.DataFrame:
    """
    Lee un HISTORICO, normaliza columnas a (MES, AÑO, NOMBRE_N, SUPERVISOR_LIDER, <kpi_nombre>).
    """
    if not ruta.is_file():
        return pd.DataFrame(columns=["MES", "AÑO", "NOMBRE_N", "SUPERVISOR_LIDER", kpi_nombre])
    df = pd.read_excel(ruta)
    # Normalizar nombres de columnas
    rename_map = {}
    for c in df.columns:
        cu = str(c).strip().upper().replace("Ñ", "N")
        if cu == "MES":           rename_map[c] = "MES"
        elif cu in ("ANO", "ANIO"): rename_map[c] = "AÑO"
        elif cu == "NOMBRE":      rename_map[c] = "NOMBRE"
        elif cu == "EMPLEADO":    rename_map[c] = "NOMBRE"
        elif cu == "SUPERVISOR_LIDER": rename_map[c] = "SUPERVISOR_LIDER"
    df = df.rename(columns=rename_map)
    if kpi_col not in df.columns or "NOMBRE" not in df.columns:
        return pd.DataFrame(columns=["MES", "AÑO", "NOMBRE_N", "SUPERVISOR_LIDER", kpi_nombre])
    df["NOMBRE_N"] = _norm_text(df["NOMBRE"])
    cols = ["MES", "AÑO", "NOMBRE_N"]
    if "SUPERVISOR_LIDER" in df.columns:
        cols.append("SUPERVISOR_LIDER")
    cols.append(kpi_col)
    out = df[cols].rename(columns={kpi_col: kpi_nombre}).copy()
    out["MES"] = pd.to_numeric(out["MES"], errors="coerce").astype("Int64")
    out["AÑO"] = pd.to_numeric(out["AÑO"], errors="coerce").astype("Int64")
    return out


def fact_cumplimientos_gestor(filtro):
    cif = _leer_historico_normalizado(paths.CIF_OUT_KPIS_HISTORICO, "TOTAL", "CUMP_CIF")
    np_ = _leer_historico_normalizado(paths.NP_OUT_KPIS_HISTORICO, "EJECUCION", "CUMP_NP")
    pre = _leer_historico_normalizado(paths.PR_OUT_KPIS_HISTORICO, "CUMPLIMIENTO", "CUMP_PRECIOS")
    sos = _leer_historico_normalizado(paths.SOS_OUT_KPIS_HISTORICO, "CUMPLIMIENTO", "CUMP_SOS")
    epg = _leer_historico_normalizado(paths.EXHIB_PAG_OUT_KPIS_HISTORICO, "CUMPLIMIENTO", "CUMP_EXH_PAG")
    egr = _leer_historico_normalizado(paths.EXHIB_GRA_OUT_KPIS_HISTORICO, "TOTAL", "CUMP_EXH_GRA")

    base = cif
    for parte in (np_, pre, sos, epg, egr):
        if parte.empty:
            continue
        # Outer merge para mantener todos los gestores
        keys = ["MES", "AÑO", "NOMBRE_N"]
        sup_cols = [c for c in parte.columns if c == "SUPERVISOR_LIDER"]
        parte_solo_kpi = parte.drop(columns=sup_cols)
        base = base.merge(parte_solo_kpi, on=keys, how="outer")

    return _filtrar_por_periodos(base, filtro).sort_values(
        ["AÑO", "MES", "NOMBRE_N"]
    ).reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# DIMENSIONES
# ─────────────────────────────────────────────────────────────────────────────

def dim_personas() -> pd.DataFrame:
    if not paths.DYP_BASE_CUPOS.is_file():
        return pd.DataFrame()
    df = pd.read_excel(paths.DYP_BASE_CUPOS, sheet_name="Tabla total roles")
    df["NOMBRE_N"] = _norm_text(df["NOMBRE"])
    return df.drop_duplicates("NOMBRE_N").reset_index(drop=True)


def dim_pdv() -> pd.DataFrame:
    if not paths.CIF_OUT_CIF.is_file():
        return pd.DataFrame()
    df = pd.read_excel(paths.CIF_OUT_CIF)
    cols = [c for c in [
        "ID_PDV_INVOLVES", "NOMBRE_PDV", "ACRONIMO",
        "VENTAS_PROMEDIO_MES", "FUENTE",
    ] if c in df.columns]
    return df[cols].drop_duplicates("ID_PDV_INVOLVES").reset_index(drop=True)


def dim_periodo(filtro) -> pd.DataFrame:
    periodos: set[tuple[int, int]] = set()
    historicos = [
        paths.CIF_OUT_KPIS_HISTORICO, paths.NP_OUT_KPIS_HISTORICO,
        paths.PR_OUT_KPIS_HISTORICO, paths.SOS_OUT_KPIS_HISTORICO,
        paths.EXHIB_PAG_OUT_KPIS_HISTORICO, paths.EXHIB_GRA_OUT_KPIS_HISTORICO,
    ]
    for ruta in historicos:
        if not ruta.is_file():
            continue
        df = pd.read_excel(ruta)
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
    GOLD_DIR.mkdir(parents=True, exist_ok=True)
    resumen: dict[str, int] = {}
    print(f"📁 Output: {GOLD_DIR}")
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
            print(f"  ⚠ {nombre:<30} (vacío) — no se escribe")
            resumen[nombre] = 0
            continue
        ruta = _to_csv(df, f"{nombre}.csv")
        print(f"  ✓ {nombre:<30} {len(df):>7,} filas → {ruta.name}")
        resumen[nombre] = len(df)
    return resumen


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        description="Generar capa GOLD (CSV para BI) — Sprint 16.3",
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
    print(" GENERACIÓN CAPA GOLD — Eficacia (Sprint 16.3)")
    print("=" * 60)
    resumen = generar_gold(filtro)
    print()
    total_filas = sum(v for v in resumen.values() if v > 0)
    print(f"🏁 Listo. {len([v for v in resumen.values() if v > 0])} tablas escritas, {total_filas:,} filas totales.")


if __name__ == "__main__":
    main()
