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
import supabase_io

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
# Columnas de Impactos donde la COMA es separador de MILES, no decimal.
# El Excel de origen las trae como texto ('62,010,000', '1,961'), y una coma
# ahí no se puede interpretar mirando el texto: '1,961' son 1.961 clientes
# georeferenciados, pero '196,100' en una columna de % es 196,1%. La misma
# forma, sentidos opuestos. Por eso la decisión va por columna y no por patrón.
#
# Sin esto, supabase_io adivinaba: dejaba Cuota y Venta como TEXTO con comas
# (Postgres no puede sumarlas) y convertía '1,961' en 1.961 — un conteo de
# clientes con decimales. Las alertas nunca se vieron afectadas porque
# cumplimiento_dyp._num_es ya quitaba todas las comas por su cuenta.
COLUMNAS_MILES_IMPACTOS = [
    "Cuota", "Venta",
    "Total Clientes a Visitar", "Clientes Visitados",
    "Total Clientes a Impactar", "Clientes Impactados",
    "Clientes Georeferenciados",
]


def limpiar_miles_impactos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte a número las columnas de plata y de conteos de Impactos.

    Las de porcentaje (`% Efect*`) NO se tocan: ahí la coma sí es decimal
    ('546,0' → 546.0) y supabase_io ya las resuelve bien.
    """
    salida = df.copy()
    tocadas = []
    for col in COLUMNAS_MILES_IMPACTOS:
        if col not in salida.columns:
            continue
        original = salida[col]
        convertido = pd.to_numeric(
            original.astype(str).str.strip().str.replace(",", "", regex=False),
            errors="coerce",
        )
        # Si la conversión rompe valores que antes tenían contenido, se deja la
        # columna como está: mejor un texto que un dato inventado.
        tenia = original.notna() & (original.astype(str).str.strip() != "")
        if (tenia & convertido.isna()).any():
            print(f"  ⚠️  Impactos: '{col}' tiene valores que no son número — se deja sin convertir.")
            continue
        salida[col] = convertido
        tocadas.append(col)
    if tocadas:
        print(f"  ℹ️  Impactos → Supabase: separador de miles quitado en {len(tocadas)} columnas "
              f"({', '.join(tocadas)})")
    return salida


# Grano de lo que se guarda en Supabase: una fila por supervisor x vendedor x
# cliente x marca x mes. Se dejan los CÓDIGOS y no los nombres porque hay
# códigos con el nombre escrito de dos formas, y agrupar por nombre partía en
# dos al mismo cliente (151.152 filas contra 151.118 agrupando solo por código).
LLAVE_VENTAS = ["Cod. Supervisor", "Cod. Vendedor", "Cod. Cliente", "Marca", "MES", "AÑO"]


def a_formato_vertical(df: pd.DataFrame, spec: pr.PeriodoSpec) -> pd.DataFrame:
    """
    Deja el DataFrame que va a Supabase: sin el formato horizontal, agrupado
    por marca y con los códigos limpios.

    Quita las 24 columnas del formato horizontal del reporte de origen
    (`Cant. Enero` … `Vta. Diciembre`) y deja el valor en CANTIDAD/VENTA.

    Por qué: el archivo de ventas viene con una columna por mes del año, pero
    cada corrida trae UN solo mes, así que 22 de esas 24 columnas son cero en
    todas las filas. En un CSV un cero ocupa un carácter; en Postgres cada una
    es un `bigint` de 8 bytes ocupe o no, y sobre ~390.000 filas eso son unos
    37 MB de tabla más el índice, sin aportar un dato.

    No se pierde nada: el mes lo identifica la columna MES/AÑO, y el valor ya
    está duplicado en `Cantidades Totales` / `Ventas Totales`, que para el mes
    del periodo son idénticas a `Cant. <mes>` / `Vta. <mes>`.

    Esa igualdad se VERIFICA antes de borrar. Si algún mes no cuadra (porque el
    reporte de origen cambió de forma), se devuelve el DataFrame intacto y se
    avisa: es preferible una tabla gorda a una que perdió plata en silencio.

    Solo se aplica a lo que va a Supabase. El CSV que queda en SharePoint
    conserva su formato original, por si alguien lo consume desde ahí.
    """
    cols_mes = [c for c in df.columns
                if any(m.lower() in c.lower() for m in MESES_ESPANOL.values())
                and any(p in c.lower() for p in ("cant", "vta", "venta"))]
    col_cant = next((c for c in df.columns if "cantidades totales" in c.lower()), None)
    col_vta  = next((c for c in df.columns if "ventas totales"     in c.lower()), None)

    if not cols_mes or not col_cant or not col_vta:
        return df

    def _num(s):
        return pd.to_numeric(
            s.astype(str).str.replace(",", ".", regex=False), errors="coerce"
        ).fillna(0)

    nombre_mes = MESES_ESPANOL[spec.mes].lower()
    c_mes = next((c for c in cols_mes if nombre_mes in c.lower() and "cant" in c.lower()), None)
    v_mes = next((c for c in cols_mes if nombre_mes in c.lower()
                  and ("vta" in c.lower() or "venta" in c.lower())), None)

    if not c_mes or not v_mes:
        print(f"  ⚠️  Ventas: no existe la columna del mes '{nombre_mes}' en el origen "
              f"— se sube el formato horizontal completo.")
        return df

    # Quedarse SOLO con el periodo antes de comparar y agrupar. El consolidado
    # que llega acumula todos los meses; `cargar_detalle` lo filtra, pero lo
    # hace DESPUÉS de esta función. Sin este recorte, la verificación comparaba
    # `Vta. Agosto` contra `Ventas Totales` en filas de marzo — donde la primera
    # es 0 y la segunda tiene la venta de marzo — y la diferencia daba justo el
    # total de los otros cinco meses, así que la salvaguarda se disparaba
    # siempre y nunca se agrupaba nada.
    if "MES" in df.columns and "AÑO" in df.columns:
        del_periodo = (
            df["MES"].astype(str).str.strip().str.lower().eq(nombre_mes)
            & df["AÑO"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
                 .eq(str(spec.anio))
        )
        if not del_periodo.any():
            print(f"  ⚠️  Ventas: el consolidado no trae filas de "
                  f"{nombre_mes} {spec.anio} — se sube sin agrupar.")
            return df
        if not del_periodo.all():
            print(f"  ℹ️  Ventas: {len(df):,} filas en el consolidado → "
                  f"{int(del_periodo.sum()):,} del periodo {spec.mes:02d}/{spec.anio}")
        df = df[del_periodo].copy()

    dif_c = (_num(df[c_mes]) - _num(df[col_cant])).abs().sum()
    dif_v = (_num(df[v_mes]) - _num(df[col_vta])).abs().sum()
    if dif_c or dif_v:
        print(f"  ⚠️  Ventas: '{v_mes}' no coincide con '{col_vta}' "
              f"(dif cantidad={dif_c:,.0f}, dif venta={dif_v:,.0f}). "
              f"NO se quitan las columnas de mes, para no perder datos.")
        return df

    plano = df.copy()
    plano["CANTIDAD"] = _num(plano[col_cant])
    plano["VENTA"] = _num(plano[col_vta])

    # Códigos: llegan como '1000.0' porque el origen los trae en float. Se
    # guardan como texto sin el sufijo — son identificadores, no números para
    # sumar, y como texto no se pierden ceros a la izquierda.
    for c in LLAVE_VENTAS:
        if c.startswith("Cod.") and c in plano.columns:
            plano[c] = (plano[c].astype(str).str.strip()
                                .str.replace(r"\.0+$", "", regex=True)
                                .replace({"nan": "", "None": ""}))

    faltan = [c for c in LLAVE_VENTAS if c not in plano.columns]
    if faltan:
        print(f"  ⚠️  Ventas: faltan columnas para agrupar {faltan} "
              f"— se sube el formato horizontal completo.")
        return df

    salida = (plano.groupby(LLAVE_VENTAS, dropna=False, as_index=False)
                   .agg(CANTIDAD=("CANTIDAD", "sum"), VENTA=("VENTA", "sum")))

    # Control: agrupar no puede mover el total.
    d_v = abs(salida["VENTA"].sum() - plano["VENTA"].sum())
    d_c = abs(salida["CANTIDAD"].sum() - plano["CANTIDAD"].sum())
    if round(d_v, 2) or round(d_c, 2):
        print(f"  ⚠️  Ventas: agrupar movió los totales "
              f"(dif venta={d_v:,.2f}, cantidad={d_c:,.2f}). Se sube sin agrupar.")
        return df

    print(f"  ℹ️  Ventas → Supabase: {len(df):,} filas × {len(df.columns)} col  →  "
          f"{len(salida):,} × {len(salida.columns)} (agrupado por marca). "
          f"El CSV de SharePoint conserva el detalle por producto.")
    return salida


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

def obtener_archivos_carpeta_sharepoint_recursivo(headers, site_id, anio_proceso: int):
    # Sin valor por defecto a propósito: un default de 2026 hacía que una
    # llamada sin año leyera de esa carpeta en silencio, para siempre.
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
    # Archivo ÚNICO acumulado: el mismo nombre sirve para leer el consolidado
    # previo (y quitarle las filas de este periodo) y para volver a escribirlo.
    # Ver la nota en paths.cloud_dyp_ventas sobre por qué ya no lleva el mes.
    nombre_consolidado_cloud = os.path.basename(paths.cloud_dyp_ventas())
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

    # Elegir qué archivos reprocesar.
    #
    # Antes esto era `[archivos_ventas[-1]]`: el ÚLTIMO archivo de la carpeta
    # del año, ordenado por ruta. Si en la carpeta convivían cortes de varios
    # meses (lo normal), un reproceso de marzo terminaba procesando el archivo
    # de mayo sin avisar. Ahora se busca el que corresponde al periodo pedido.
    if consolidado_existe:
        mes_pedido = MESES_ESPANOL[spec.mes].capitalize()
        del_periodo = []
        for archivo in archivos_ventas:
            mes_det, anio_det = _detectar_mes_anio_cloud(
                archivo["@microsoft.graph.downloadUrl"], archivo["name"], mapping)
            if (mes_det or "").capitalize() == mes_pedido and str(anio_det) == str(spec.anio):
                del_periodo.append(archivo)

        if del_periodo:
            archivos_a_procesar = del_periodo
            print(f"  ℹ️  Reprocesando {len(del_periodo)} archivo(s) de {spec.etiqueta}: "
                  f"{[a['name'] for a in del_periodo]}")
        else:
            # Sin coincidencia no se procesa nada: tomar "el último" sería
            # cargar otro mes encima del que se pidió.
            raise FileNotFoundError(
                f"No se encontró ningún archivo de ventas de {spec.etiqueta} en la carpeta del año.\n"
                f"  Archivos disponibles: {[a['name'] for a in archivos_ventas]}"
            )
    else:
        archivos_a_procesar = archivos_ventas

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
            # Se comparan ambos lados en MAYÚSCULAS.
            #
            # Antes era `.str.capitalize() != MESES_ESPANOL[spec.mes]`, o sea
            # 'Agosto' contra 'AGOSTO': nunca coincidían, el filtro no quitaba
            # nada y cada corrida AÑADÍA otro juego completo de filas del mes al
            # consolidado. Se veía en el crecimiento: 366.925 → 417.109 →
            # 467.293 filas en tres corridas del mismo periodo.
            mes_up = MESES_ESPANOL[spec.mes].upper()
            n_antes = len(df_existente)
            df_filtrado = df_existente[
                (df_existente["MES"].astype(str).str.strip().str.upper() != mes_up)
                | (df_existente["AÑO"].astype(str).str.strip() != str(spec.anio))
            ]
            bloques.append(df_filtrado)
            print(f"  Modo upsert cloud (Ventas) — {n_antes - len(df_filtrado):,} filas "
                  f"previas de {spec.etiqueta} reemplazadas")
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
        supabase_io.cargar_detalle_seguro(
            "dyp_ventas_detalle",
            a_formato_vertical(df_final, spec),
            spec.mes, spec.anio,
        )
    else:
        print(f"  ❌ Error al subir CSV consolidado de ventas: {res_subida.text}")

    # La carga a Supabase la hace `cargar_detalle_seguro` arriba, contra la
    # tabla dyp_ventas_detalle. La antigua tabla `dyp_ventas` usaba upsert sin
    # llave única, así que reejecutar el mismo mes duplicaba las filas.

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
    # Archivo ÚNICO acumulado — ver la nota en paths.cloud_dyp_impactos.
    nombre_consolidado_cloud = os.path.basename(paths.cloud_dyp_impactos())
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
        supabase_io.cargar_detalle_seguro(
            "dyp_impactos_detalle",
            limpiar_miles_impactos(df_final),
            spec.mes, spec.anio,
        )
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
        # Las carpetas de BASES llevan el año en la ruta; sin esto se
        # buscarían los insumos en la carpeta del año del reloj.
        paths.usar_anio(spec.anio)

        print(f"\n🎯 Procesando periodo: {spec.etiqueta} (Año: {spec.anio})")
        
        rutas_cloud = obtener_archivos_carpeta_sharepoint_recursivo(headers, site_id, anio_proceso=spec.anio)
        print(f"✅ Archivos encontrados en la nube para {spec.anio}: {len(rutas_cloud)}")
        
        total_filas += consolidar_ventas(rutas_cloud, spec, headers, site_id)
        consolidar_impactos(rutas_cloud, spec, headers, site_id)

        # `consolidar_impacto_segmentos` ya no se llama aquí.
        #
        # Buscaba archivos con "segmento" en BASES DE RESPUESTAS/DYP, pero el
        # archivo de segmentos no es un insumo: lo CALCULA etl_impactos_segmentos
        # y lo publica en SALIDAS/DYP/impacto_segmentos.xlsx. Es decir, buscaba
        # en la carpeta de entrada algo que se produce en la de salida, así que
        # nunca encontraba nada y dejaba un "❌ Sin archivos de impacto_segmentos"
        # en cada corrida — un error que no lo era.

    return total_filas


if __name__ == "__main__":
    run()