"""
etl_impactos.py
───────────────
Consolida los archivos Excel mensuales de **Impactos D&P** (Cuota, Venta,
Clientes Visitar/Visitados/Impactar/Impactados/Georeferenciados, %Efectividad)
en un único CSV (`Consolidado_Impactos.csv`).

Origen
──────
    paths.DYP_IMPACTOS_DIR/<año>/<archivo>.xlsx

Filtros y modo upsert: idénticos a `etl_ventas.py` (ver docstring).

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

import os
import re
import glob
import warnings
import locale
import numpy as np
import pandas as pd

import paths

warnings.filterwarnings("ignore", category=FutureWarning)

try:
    locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, "Spanish")
    except locale.Error:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE COLUMNAS
# ─────────────────────────────────────────────────────────────────────────────
COLUMN_MAPPING_ORIGEN_A_DESTINO = {
    "Supervisor": "Supervisor", "Asesor": "Asesor", "Ciudad": "Ciudad",
    "Cuota": "Cuota", "Venta": "Venta",
    "Total Clientes a Visitar":  "Total Clientes Visitar",
    "Clientes Visitados":        "Clientes Visitados",
    "Total Clientes a Impactar": "Total Clientes Impactar",
    "Clientes Impactados":       "Clientes Impactados",
    "Clientes Georeferenciados": "Clientes Georeferenciados",
    # Variantes del nombre de las 4 columnas de efectividad. Pandas añade
    # ".1/.2/.3" cuando hay duplicados en el header; algunos exports
    # corporativos vienen con "_1/_2/_3" porque el archivo origen ya viene
    # con ese sufijo. Soportamos ambas:
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
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _descubrir_excels(directorio: os.PathLike) -> list[str]:
    """Igual que en etl_ventas: recursivo, filtra años ≤ 2023 y 'MES ACTUAL'."""
    if not os.path.isdir(directorio):
        return []
    todos = (
        glob.glob(os.path.join(directorio, "**", "*.xlsx"), recursive=True)
        + glob.glob(os.path.join(directorio, "**", "*.xls"), recursive=True)
    )
    seleccionados = []
    for ruta in todos:
        if "MES ACTUAL" in ruta.upper():
            continue
        partes = ruta.split(os.sep)
        es_valido = True
        for p in partes:
            if p.isdigit() and len(p) == 4 and int(p) <= 2023:
                es_valido = False
                break
        if es_valido:
            seleccionados.append(ruta)
    return sorted(seleccionados)


def _detectar_mes_anio(ruta: str, mapping: dict) -> tuple[str | None, str | None]:
    nombre = os.path.basename(ruta)
    mes_val: str | None = None
    anio_val: str | None = None
    patron = r"([A-Za-z]+)\s(\d{4})"

    try:
        df_head = pd.read_excel(ruta, nrows=10)
        df_head.rename(columns=mapping, inplace=True)
        if "MES" in df_head.columns and not df_head["MES"].dropna().empty:
            mes_val = str(df_head["MES"].dropna().iloc[0]).strip().capitalize()
        if "AÑO" in df_head.columns and not df_head["AÑO"].dropna().empty:
            try:
                anio_val = str(int(df_head["AÑO"].dropna().iloc[0]))
            except Exception:
                anio_val = None
        if not mes_val or not anio_val:
            m = re.search(patron, nombre)
            if m:
                if not mes_val:
                    mes_val = m.group(1).capitalize()
                if not anio_val:
                    anio_val = m.group(2)
    except Exception:
        m = re.search(patron, nombre)
        if m:
            mes_val = m.group(1).capitalize()
            anio_val = m.group(2)

    # Casos raros: MES viene como fecha ISO → extraer mes
    if mes_val and len(mes_val) > 10 and re.match(r"^\d{4}-\d{2}-\d{2}", mes_val):
        try:
            mes_val = pd.to_datetime(mes_val).strftime("%B").capitalize()
        except Exception:
            pass

    return mes_val, anio_val


# ─────────────────────────────────────────────────────────────────────────────
# CONSOLIDACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def consolidar(
    rutas_excel: list[str],
    salida_csv: os.PathLike,
    mapping: dict = COLUMN_MAPPING_ORIGEN_A_DESTINO,
    delimitador: str = DELIMITADOR_CSV,
) -> int:
    if not rutas_excel:
        print("❌ Sin archivos de Impactos para consolidar.")
        return 0

    salida_csv = str(salida_csv)
    consolidado_existe = os.path.exists(salida_csv)
    archivos_a_procesar = [rutas_excel[-1]] if consolidado_existe else rutas_excel

    bloques: list[pd.DataFrame] = []
    orden_columnas: list[str] | None = None
    mes_act = anio_act = None

    if consolidado_existe:
        mes_act, anio_act = _detectar_mes_anio(archivos_a_procesar[0], mapping)
        if not (mes_act and anio_act):
            print("  ⚠️  No pude detectar mes/año del archivo más reciente; reconstruyo")
            archivos_a_procesar = rutas_excel
            consolidado_existe = False

    if consolidado_existe:
        try:
            df_existente = pd.read_csv(
                salida_csv, sep=delimitador, decimal=",",
                encoding="utf-8",
                dtype={"MES": str, "AÑO": str},
                low_memory=False,
            )
            orden_columnas = list(df_existente.columns)
            df_existente["__M"] = df_existente["MES"].astype(str).str.strip().str.capitalize()
            df_existente["__A"] = df_existente["AÑO"].astype(str).str.strip()
            df_filtrado = df_existente[
                (df_existente["__M"] != mes_act) | (df_existente["__A"] != anio_act)
            ].drop(columns=["__M", "__A"])
            bloques.append(df_filtrado)
            print(f"  Modo upsert — reprocesando: {os.path.basename(archivos_a_procesar[0])}")
        except Exception as e:
            print(f"  ⚠️  No pude leer el consolidado previo ({e}); reconstruyo desde cero")
            archivos_a_procesar = rutas_excel
            bloques = []

    for ruta in archivos_a_procesar:
        nombre = os.path.basename(ruta)
        mes, anio = (mes_act, anio_act) if (consolidado_existe) else _detectar_mes_anio(ruta, mapping)
        if not mes or not anio:
            print(f"  ⚠️  Sin mes/año detectable, omitido: {nombre}")
            continue

        print(f"  → Procesando: {nombre}")
        try:
            try:
                df = pd.read_excel(ruta)
            except Exception:
                df_list = pd.read_html(ruta)
                df = df_list[0] if df_list else None
            if df is None or df.empty:
                continue

            df.rename(columns=mapping, inplace=True)
            df = df.loc[:, ~df.columns.duplicated()]
            df["MES"] = mes
            df["AÑO"] = anio

            if orden_columnas is None:
                orden_columnas = list(df.columns)
                for c in ("MES", "AÑO"):
                    if c not in orden_columnas:
                        orden_columnas.append(c)

            df = df[[c for c in orden_columnas if c in df.columns]]

            # Limpieza códigos
            for col in COLUMNS_TO_CLEAN_DOT_ZERO:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")
                    df[col] = (
                        df[col].astype(str)
                              .replace(r"\.0$", "", regex=True)
                              .replace("nan", "")
                    )

            # Numéricos: cuota y venta — formato colombiano (1.234.567,89)
            for col in VTA_COLUMNS_CASH:
                if col in df.columns:
                    s = (df[col].astype(str).str.strip()
                                  .str.replace(",", "", regex=False)
                                  .str.replace(".", "", regex=False))
                    df[col] = pd.to_numeric(s, errors="coerce").fillna(0).astype(np.int64)

            # Porcentajes (vienen como "1.234,5" → convertir a float)
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
        print("❌ No se generó ningún bloque consolidable.")
        return 0

    df_final = pd.concat(bloques, ignore_index=True)
    df_final["AÑO"] = (
        pd.to_numeric(df_final["AÑO"], errors="coerce")
          .fillna("")
          .astype(pd.StringDtype())
    )

    os.makedirs(os.path.dirname(salida_csv) or ".", exist_ok=True)
    df_final.to_csv(
        salida_csv,
        sep=delimitador, index=False, encoding="utf-8",
        decimal=",", mode="w", header=True,
    )
    print(f"  ✅ Consolidado guardado: {salida_csv}  ({len(df_final):,} filas)")
    return len(df_final)


# ─────────────────────────────────────────────────────────────────────────────
# PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────────────────────────

def run() -> int:
    print(f"📍 Origen: {paths.DYP_IMPACTOS_DIR}")
    print(f"📁 Salida: {paths.DYP_OUT_IMPACTOS}")

    if not paths.DYP_IMPACTOS_DIR.is_dir():
        print(
            f"❌ Carpeta de origen no existe: {paths.DYP_IMPACTOS_DIR}\n"
            f"   Crea la carpeta con los Excel de impactos (organizados por año)."
        )
        return 0

    rutas = _descubrir_excels(paths.DYP_IMPACTOS_DIR)
    print(f"✅ Archivos encontrados (2024+): {len(rutas)}")
    return consolidar(rutas, paths.DYP_OUT_IMPACTOS)


if __name__ == "__main__":
    run()
