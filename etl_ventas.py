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
# CONFIGURACIÓN DE COLUMNAS  (preservada del archivo de ejemplo del usuario)
# ─────────────────────────────────────────────────────────────────────────────
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
    # Cantidades + ventas por mes
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


# ─────────────────────────────────────────────────────────────────────────────
# DESCUBRIMIENTO DE ARCHIVOS DE ORIGEN
# ─────────────────────────────────────────────────────────────────────────────

def _descubrir_excels(directorio: os.PathLike) -> list[str]:
    """
    Busca recursivamente *.xlsx / *.xls dentro de `directorio`, excluyendo:
      • subcarpetas con "MES ACTUAL" (cortes parciales).
      • subcarpetas con nombre numérico ≤ 2023.
    Devuelve la lista ordenada (cronológicamente, asumiendo nombres con fecha).
    """
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
        # Filtrar carpetas año <= 2023
        partes = ruta.split(os.sep)
        es_valido = True
        for p in partes:
            if p.isdigit() and len(p) == 4 and int(p) <= 2023:
                es_valido = False
                break
        if es_valido:
            seleccionados.append(ruta)
    return sorted(seleccionados)


def _detectar_mes_anio(excel_path: str, mapping: dict) -> tuple[str | None, str | None]:
    """
    Intenta deducir (mes, año) del archivo. Primero busca en las primeras 10
    filas las columnas MES/AÑO; si fallan, lo extrae del nombre del archivo
    con la regex `<Palabra> <YYYY>`.
    """
    nombre = os.path.basename(excel_path)
    mes_val: str | None = None
    anio_val: str | None = None

    try:
        df_head = pd.read_excel(excel_path, nrows=10)
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
        m = re.search(r"([A-Za-z]+)\s(\d{4})", nombre)
        if m:
            if not mes_val:
                mes_val = m.group(1).capitalize()
            if not anio_val:
                anio_val = m.group(2)

    return mes_val, anio_val


# ─────────────────────────────────────────────────────────────────────────────
# CARGA + CONSOLIDACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def consolidar(
    rutas_excel: list[str],
    salida_csv: os.PathLike,
    mapping: dict = COLUMN_MAPPING_ORIGEN_A_DESTINO,
    delimitador: str = DELIMITADOR_CSV,
) -> int:
    """
    Procesa los Excel y escribe el CSV consolidado.
    Devuelve el número de filas finales escritas.
    """
    if not rutas_excel:
        print("❌ Sin archivos de Ventas para consolidar.")
        return 0

    salida_csv = str(salida_csv)
    consolidado_existe = os.path.exists(salida_csv)
    archivos_a_procesar = [rutas_excel[-1]] if consolidado_existe else rutas_excel

    bloques: list[pd.DataFrame] = []
    orden_columnas: list[str] | None = None

    if consolidado_existe:
        mes_act, anio_act = _detectar_mes_anio(archivos_a_procesar[0], mapping)
        try:
            df_existente = pd.read_csv(
                salida_csv, sep=delimitador, decimal=",",
                encoding="utf-8",
                dtype={"MES": str, "AÑO": str},
                low_memory=False,
            )
            orden_columnas = list(df_existente.columns)
            df_filtrado = df_existente[
                (df_existente["MES"].astype(str).str.strip().str.capitalize() != mes_act)
                | (df_existente["AÑO"].astype(str).str.strip() != anio_act)
            ]
            bloques.append(df_filtrado)
            print(f"  Modo upsert — reprocesando: {os.path.basename(archivos_a_procesar[0])}")
        except Exception as e:
            print(f"  ⚠️  No pude leer el consolidado previo ({e}); reconstruyo desde cero")
            archivos_a_procesar = rutas_excel
            bloques = []

    for ruta in archivos_a_procesar:
        nombre = os.path.basename(ruta)
        mes, anio = _detectar_mes_anio(ruta, mapping)
        if not mes or not anio:
            print(f"  ⚠️  Sin mes/año detectable, omitido: {nombre}")
            continue

        print(f"  → Procesando: {nombre}")
        try:
            try:
                df = pd.read_excel(ruta)
            except Exception:
                df = pd.read_html(ruta, encoding="latin-1")[0]

            df.rename(columns=mapping, inplace=True)
            df = df.loc[:, ~df.columns.duplicated()]
            df["MES"], df["AÑO"] = mes, anio

            if orden_columnas is None:
                orden_columnas = list(df.columns)
                if "MES" not in orden_columnas:
                    orden_columnas.extend(["MES", "AÑO"])

            df = df[[c for c in orden_columnas if c in df.columns]]

            # Limpieza códigos / cantidades
            for col in COLUMNS_TO_CLEAN_DOT_ZERO:
                if col in df.columns:
                    df[col] = (
                        pd.to_numeric(df[col], errors="coerce")
                          .astype(str)
                          .replace(r"\.0$", "", regex=True)
                          .replace("nan", "")
                    )
            # Numéricos de venta
            for col in VTA_COLUMNS:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

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
    """Función esperada por el orquestador `run_all.py`."""
    print(f"📍 Origen: {paths.DYP_VENTAS_DIR}")
    print(f"📁 Salida: {paths.DYP_OUT_VENTAS}")

    if not paths.DYP_VENTAS_DIR.is_dir():
        print(
            f"❌ Carpeta de origen no existe: {paths.DYP_VENTAS_DIR}\n"
            f"   Crea la carpeta con los Excel de ventas (organizados por año)."
        )
        return 0

    rutas = _descubrir_excels(paths.DYP_VENTAS_DIR)
    print(f"✅ Archivos encontrados (2024+): {len(rutas)}")
    return consolidar(rutas, paths.DYP_OUT_VENTAS)


if __name__ == "__main__":
    run()
