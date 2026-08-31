"""
generar_sql_tablas.py
─────────────────────
Lee los archivos de salida reales desde SharePoint y emite el script SQL de
creación de las tablas de detalle en Supabase.

Cuándo se usa
─────────────
El SQL vigente ya está en DOCS/supabase_tablas.sql, generado de los archivos
reales de JULIO 2026. Este script NO hace falta para la operación normal.

Se usa cuando una carga empieza a fallar con un error de columna, porque 4 de
los 8 archivos tienen estructura VARIABLE: se arman haciendo merge con un Excel
de origen cuyos encabezados puede cambiar Involves o el equipo de D&P.

    Fijas     → CIF, Análisis Agotados, Análisis Precios, Exhibiciones Gratis
    Variables → Consolidado Ventas, Consolidado Impactos,
                Exhibiciones Pagadas (detalle), Reporte SOS Final

En ese caso, este script lee los archivos del periodo indicado y regenera el
CREATE TABLE con las columnas y tipos que tienen HOY, para ver qué cambió.

Usa el mismo criterio de tipos que `supabase_io.preparar_registros`, así que
lo que declara es exactamente lo que la ETL va a insertar.

Uso
───
    python generar_sql_tablas.py --mes 7 --anio 2026
    python generar_sql_tablas.py --mes 7 --anio 2026 --out DOCS/supabase_tablas.sql

Luego se pega el resultado en el editor SQL de Supabase.
"""

from __future__ import annotations

import argparse
import io
import sys
import urllib.parse

import msal
import pandas as pd
import requests
from dotenv import load_dotenv

import paths
import periodo_resolver as pr
from supabase_io import COLUMNA_CARGA, _intentar_numerico, normalizar_columnas

load_dotenv()

try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass


# ─────────────────────────────────────────────────────────────────────────────
# ACCESO A SHAREPOINT
# ─────────────────────────────────────────────────────────────────────────────

def _token() -> str:
    app = msal.ConfidentialClientApplication(
        paths.AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{paths.AZURE_TENANT_ID}",
        client_credential=paths.AZURE_CLIENT_SECRET,
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"Error autenticando en Azure: {res.get('error_description')}")
    return res["access_token"]


def _site_id(headers) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{paths.SHAREPOINT_SITE_NAME}"
    res = requests.get(url, headers=headers).json()
    if "id" not in res:
        raise RuntimeError(f"No se encontró el sitio '{paths.SHAREPOINT_SITE_NAME}'")
    return res["id"]


def _descargar(headers, site_id: str, ruta: str) -> bytes | None:
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{urllib.parse.quote(ruta)}:/content"
    )
    r = requests.get(url, headers=headers)
    return r.content if r.status_code == 200 else None


def _buscar_en_carpeta(headers, site_id: str, carpeta: str, contiene: str) -> str | None:
    """Devuelve la ruta completa del primer archivo de `carpeta` que contenga `contiene`."""
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{urllib.parse.quote(carpeta)}:/children"
    )
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        return None
    for item in r.json().get("value", []):
        if contiene.lower() in item.get("name", "").lower():
            return f"{carpeta}/{item['name']}"
    return None


# ─────────────────────────────────────────────────────────────────────────────
# INFERENCIA DE TIPOS
# ─────────────────────────────────────────────────────────────────────────────

def tipo_sql(serie: pd.Series, columna: str) -> str:
    """
    Deduce el tipo Postgres a partir de los datos reales de la columna.

    Mismo criterio con el que se generó DOCS/supabase_tablas.sql:
      · Enteros del archivo → bigint
      · CUALQUIER decimal   → numeric, aunque la muestra traiga solo enteros
        redondos. Tentador ponerle bigint, pero basta un decimal en un mes
        futuro para que la carga falle.
      · Texto que en realidad es número con coma decimal → numeric
        (`preparar_registros` lo convierte antes de insertar)
      · Texto ambiguo (p. ej. '30,000,000', separador de miles) → text
    """
    if columna in ("mes", "anio"):
        return "integer"
    if columna == "cargado_en":
        return "timestamp"

    if pd.api.types.is_bool_dtype(serie):
        return "boolean"
    if pd.api.types.is_datetime64_any_dtype(serie):
        return "timestamp"
    if pd.api.types.is_integer_dtype(serie):
        return "bigint"
    if pd.api.types.is_float_dtype(serie):
        return "numeric"
    if _intentar_numerico(serie) is not None:
        return "numeric"

    # Los IDs se dejan como texto a propósito: vienen con ceros a la izquierda
    # y mezclados con letras según la fuente, y convertirlos a número rompe
    # los cruces silenciosamente.
    return "text"


def generar_create_table(tabla: str, df: pd.DataFrame) -> str:
    df = normalizar_columnas(df)

    # Columnas que agrega el cargue y que no vienen en el archivo.
    if "mes" not in df.columns:
        df["mes"] = 0
    if "anio" not in df.columns:
        df["anio"] = 0
    df[COLUMNA_CARGA] = pd.NaT

    ancho = max(len(c) for c in df.columns) + 2
    lineas = [f"    {col:<{ancho}}{tipo_sql(df[col], col)}" for col in df.columns]

    cuerpo = ",\n".join(lineas)
    return f"""-- ─────────────────────────────────────────────────────────────
-- {tabla}   ({len(df.columns)} columnas · muestra de {len(df):,} filas)
-- ─────────────────────────────────────────────────────────────
drop table if exists public.{tabla};

create table public.{tabla} (
{cuerpo}
);

-- Índice por periodo: lo usa el DELETE de cada corrida y los filtros de BI.
create index if not exists idx_{tabla}_periodo
    on public.{tabla} (anio, mes);

alter table public.{tabla} enable row level security;

create policy "lectura_publica_{tabla}"
    on public.{tabla} for select
    using (true);
"""


# ─────────────────────────────────────────────────────────────────────────────
# CATÁLOGO DE ARCHIVOS
# ─────────────────────────────────────────────────────────────────────────────

def catalogo(spec: pr.PeriodoSpec) -> list[tuple[str, str, str, str]]:
    """
    (tabla, carpeta, texto_a_buscar, formato)

    `formato`: 'excel' o 'csv_pipe' (los consolidados de D&P salen con
    separador '|' y coma decimal).
    """
    mes_up = spec.mes_str_upper
    mes_cap = spec.mes_str
    anio = spec.anio

    return [
        ("cif_detalle",
         paths.RUTA_CARPETA_SALIDAS_CIF, "CIF.xlsx", "excel"),

        ("np_agotados_detalle",
         paths.RUTA_CARPETA_SALIDAS_NP, f"ANALISIS_AGOTADOS_{mes_up}_{anio}", "excel"),

        ("precios_analisis_detalle",
         paths.RUTA_CARPETA_SALIDAS_PRECIOS, f"ANALISIS_PRECIOS_{mes_up}_{anio}", "excel"),

        ("sos_detalle",
         paths.RUTA_CARPETA_SALIDAS_SOS, f"Reporte_SOS_Final_Calculado_{mes_up}_{anio}", "excel"),

        ("exhibiciones_gratis_detalle",
         paths.RUTA_CARPETA_SALIDAS_EXHIB, "Resultado exhibiciones gratis.xlsx", "excel"),

        ("exhibiciones_pagadas_detalle",
         paths.RUTA_CARPETA_SALIDAS_EXHIB, f"Resultado_exhibiciones_pagadas {mes_cap} {anio}", "excel"),

        ("dyp_ventas_detalle",
         paths.RUTA_CARPETA_SALIDAS_DYP, f"Consolidado_Ventas_{mes_up}_{anio}", "csv_pipe"),

        ("dyp_impactos_detalle",
         paths.RUTA_CARPETA_SALIDAS_DYP, f"Consolidado_Impactos_{mes_up}_{anio}", "csv_pipe"),
    ]


def main():
    parser = argparse.ArgumentParser(
        description="Genera el SQL de las tablas de detalle leyendo los archivos reales",
    )
    pr.cli_add_periodo_args(parser)
    parser.add_argument("--out", default=None,
                        help="Archivo donde guardar el SQL (por defecto lo imprime)")
    args = parser.parse_args()
    spec = pr.periodo_de_args(args)

    print(f"🔒 Autenticando con Microsoft Graph…")
    headers = {"Authorization": f"Bearer {_token()}"}
    site_id = _site_id(headers)

    print(f"📖 Leyendo archivos de salida del periodo {spec.etiqueta}\n")

    partes = [
        "-- ═══════════════════════════════════════════════════════════════",
        "--  TABLAS DE DETALLE — Proyecto Eficacia (Kenvue)",
        f"--  Generado a partir de los archivos reales del periodo {spec.etiqueta}",
        "--",
        "--  Cada corrida de la ETL hace:",
        "--     DELETE FROM <tabla> WHERE mes = M AND anio = A",
        "--     INSERT  INTO <tabla> ...",
        "--  Así, reejecutar el mismo mes reemplaza sus filas en vez de duplicarlas.",
        "-- ═══════════════════════════════════════════════════════════════",
        "",
    ]

    faltantes = []
    for tabla, carpeta, contiene, formato in catalogo(spec):
        ruta = _buscar_en_carpeta(headers, site_id, carpeta, contiene)
        if not ruta:
            print(f"  ✗ {tabla:<32} no se encontró '{contiene}' en {carpeta}")
            faltantes.append((tabla, contiene, carpeta))
            continue

        contenido = _descargar(headers, site_id, ruta)
        if not contenido:
            print(f"  ✗ {tabla:<32} no se pudo descargar {ruta}")
            faltantes.append((tabla, contiene, carpeta))
            continue

        try:
            if formato == "csv_pipe":
                df = pd.read_csv(io.BytesIO(contenido), sep="|", decimal=",",
                                 encoding="utf-8", low_memory=False)
            else:
                df = pd.read_excel(io.BytesIO(contenido))
        except Exception as e:
            print(f"  ✗ {tabla:<32} no se pudo leer: {e}")
            faltantes.append((tabla, contiene, carpeta))
            continue

        print(f"  ✓ {tabla:<32} {len(df.columns):>3} columnas · {len(df):>7,} filas")
        partes.append(generar_create_table(tabla, df))
        partes.append("")

    sql = "\n".join(partes)

    if args.out:
        from pathlib import Path
        destino = Path(args.out)
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text(sql, encoding="utf-8")
        print(f"\n💾 SQL guardado en: {destino}")
    else:
        print("\n" + "=" * 70 + "\n")
        print(sql)

    if faltantes:
        print("\n⚠️  No se pudo generar el SQL de estas tablas:")
        for tabla, contiene, carpeta in faltantes:
            print(f"     · {tabla}: falta '{contiene}' en {carpeta}")
        print("   Corre primero el ETL de ese módulo para el periodo indicado.")


if __name__ == "__main__":
    main()
