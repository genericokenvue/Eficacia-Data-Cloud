"""
etl_espacios_errores.py
────────────────────────
Detecta dos tipos de captura sospechosa en la encuesta SOS (espacios):

  1. MOTIVO_UNIVERSO — la suma de cm de marca supera el universo
     capturado para esa categoría en ese PDV: algo imposible, no puede haber
     más centímetros de marca que centímetros totales de la categoría.
  2. MOTIVO_CAMBIO — una marca puntual cambió mucho de cm
     de un mes al otro: el anaquel no cambia de tamaño de golpe, así que un
     salto grande suele ser un dígito de más o de menos, no una realidad.

Los dos motivos comparten el mismo archivo de salida y el mismo mecanismo de
ID_ERROR — la columna MOTIVO distingue cuál disparó cada fila.

CÓMO CAPTURA LA ENCUESTA
────────────────────────
Cada visita SOS trae, por (Empleado, PDV, Línea de producto):
  · el universo en cm de la categoría — una vez por línea
  · el cm de cada marca presente (propia y competencia) — una vez por marca

MOTIVO 1 — "Error de captura del mes"  (constante MOTIVO_UNIVERSO)
─────────────────────────────────────────────────────────────────
Se agrupa por (Empleado, PDV, Categoría real):

    UNIVERSO_CM    = promedio del universo (debería ser un único valor
                      constante dentro del grupo; se promedia por si hay
                      variación de tipeo entre filas de la misma categoría)
    SUMA_MARCAS_CM = suma de cm de todas las marcas de esa categoría
    DIFERENCIA     = UNIVERSO_CM - SUMA_MARCAS_CM

Error si DIFERENCIA < 0. Es el único criterio para este motivo — a
diferencia del motivo 2, acá no hay severidad — o la suma de marcas superó
al universo (imposible) o no hay error.

Esta lógica se validó 1 a 1 contra el chequeo manual que ya existía en Excel
("DETALLE ERRORES" / "UNIVERSO" de ARCHIVO BASE ESPACIOS): mismas 12.064
combinaciones y mismas 38 en error para agosto 2026 — pero ese archivo
manual es una versión ya curada. La encuesta cruda que consume este script
(`Encuesta_Sos_Consolidada_<periodo>.xlsx`) tiene tres particularidades que
el chequeo manual no tenía y que sí hay que resolver acá:

  1. La columna de PDV se llama 'PDV', no 'Nombre de PDV'.
  2. 'Línea de producto' viene como "Categoría - Marca" (ej. "Toallas -
     Nosotras"), no como categoría sola — agrupar por esa columna tal cual
     compara universo contra una sola marca repetida, no contra la suma de
     todas. Hay que restarle la Marca (columna aparte) para recuperar la
     categoría real, y esa categoría se escribe con ortografía distinta
     según qué exportación la trajo (p.ej. "Enj. Bucales Masivos" vs.
     "Enjuagues Bucales Masivos") — se normaliza contra las 17 categorías
     oficiales, igual que hace `etl_sos.py`.
  3. El nombre de la columna de cm-por-marca cambia de exportación a
     exportación ('¿Cuántos cms/CMS/CM tiene la marca?', con distinta
     tilde/mayúscula) y las tres variantes son mutuamente excluyentes fila a
     fila — hay que sumarlas, no quedarse con la primera que aparezca.
  4. Históricamente el archivo ACUMULABA varios meses (una consolidada de
     agosto traía marzo-agosto adentro). Eso rompía el chequeo: sumaba 6
     meses de capturas contra un universo de un solo mes y marcaba el 58% de
     los grupos. Se corrigió en `etl_sos.py` (paso 'enc'), que ahora filtra
     por 'Mes del año'/'Año' — la consolidada de agosto es SOLO agosto. Igual
     este script vuelve a acotar por periodo por si acaso (`_preparar`).
     Ver conversación del 2026-09-10.

Adrede NO se marcan las brechas grandes en sentido contrario (mucho universo
capturado y poca marca) — hay casos así, pero decidir su umbral quedó
pendiente; ver la misma conversación.

MOTIVO 2 — "Comparativa contra el promedio de los últimos 3 meses"
(constante MOTIVO_CAMBIO)
────────────────────────────────────────────────────────────────────────
Compara, para cada marca PROPIA (columna 'Origen de la línea de producto'
== 'Propio' — NO se revisan marcas de competencia acá), su PARTICIPACIÓN
de este mes contra el promedio de su participación en los 3 meses
anteriores:

    PARTICIPACION = CM_MARCA / UNIVERSO_CM        (campos origen de la
                                                     encuesta, sin tocar)

Se compara participación (un %) y no cm crudo a propósito: el universo de
una categoría en un PDV cambia de un mes a otro (remodelaciones, nuevas
mediciones), así que un cm distinto puede ser solo porque cambió el
universo, no porque la marca perdió o ganó espacio real. La participación
normaliza eso.

Los 3 meses anteriores salen de sus propios archivos
(`Encuesta_Sos_Consolidada_<MES>_<AÑO>.xlsx`, uno por mes), descargados
aparte. El promedio se calcula con los que sí existan (mínimo 1) — si
ninguno existe, la comparativa se salta. Por marca, el promedio solo cuenta
los meses donde esa marca tuvo participación > 0 (que entre o salga del
surtido no es, por sí sola, un error).

    DIFERENCIA = |PARTICIPACION - PARTICIPACION_PROMEDIO_3M|

    DIFERENCIA > 20 puntos porcentuales  →  error

Umbral único, sin niveles de severidad — la severidad por tramos (antes
50/100/200% sobre cm crudo) generaba ruido sin aportar a la decisión de
revisar o no. Ver conversación del 2026-09-14.

DE DÓNDE SALE EL SUPERVISOR
───────────────────────────
Se arma cruzando SOS_KPIS.xlsx (los propios gestores de espacios) primero, y
para quien no aparezca ahí, los otros KPIs (No presencia, CIF, Precios) que
también traen SUPERVISOR_LIDER — igual que en exhibiciones.

CADA CASO SE VE COMPLETO, NO SOLO EL TOTAL
───────────────────────────────────────────
El archivo no trae una fila por caso con el total nada más — trae una fila
por CADA MARCA que compone ese total, todas con el mismo ID_ERROR. Así el
supervisor ve, para "Enjuagues Bucales" en tal PDV, el cm de Listerine, de
Plax, de Marca Propia, etc. uno por uno — y puede detectar cuál marca en
particular quedó mal tipeada, en vez de solo saber que el total no cuadra.

EL ARCHIVO DE SALIDA ES TAMBIÉN EL FORMULARIO
─────────────────────────────────────────────
`ERRORES_ESPACIOS_<MES>_<AÑO>.xlsx` trae tres columnas vacías —
CM_MARCA_CORREGIDO, UNIVERSO_CORREGIDO y OBSERVACION_SUPERVISOR — porque acá,
a diferencia de precios/exhibiciones, el dato equivocado puede ser cualquiera
de los dos: una marca puntual con cm de más, o el universo mal tipeado. En
cada corrida se conservan las correcciones ya escritas — el cm de cada marca
se conserva por su propia fila (Empleado+PDV+Categoría+Marca); si el
supervisor escribe la observación o el universo corregido en una sola fila
del caso, esa fila puntual es la que lo conserva entre corridas.

USO
───
    python etl_espacios_errores.py --mes 8 --anio 2026
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
import time
import unicodedata
import urllib.parse

import msal
import pandas as pd
import requests
from dotenv import load_dotenv

import paths
import periodo_resolver as pr

load_dotenv()

# ─────────────────────────────────────────────────────────────────────────────
# PARÁMETROS
# ─────────────────────────────────────────────────────────────────────────────

COLS_RESPUESTA = ["CM_MARCA_CORREGIDO", "UNIVERSO_CORREGIDO", "OBSERVACION_SUPERVISOR"]

# Identifica el CASO (gestor, PDV, categoría) — de esto sale el ID_ERROR del
# motivo MOTIVO_UNIVERSO. Varias filas de marca comparten la misma llave
# de caso, y por lo tanto el mismo ID_ERROR: son el mismo error, visto marca
# por marca.
COLS_LLAVE_CASO = ["Empleado", "PDV", "LINEA_PRODUCTO"]

# Identifica cada FILA de detalle (una marca dentro de un caso) — de esto
# sale qué corrección puntual se conserva entre corridas, y también el
# ID_ERROR del motivo MOTIVO_CAMBIO (ahí cada fila YA es su
# propio caso, no hay nada que agrupar).
COLS_LLAVE_FILA = COLS_LLAVE_CASO + ["MARCA"]

# Motivos posibles de una fila en error. El texto es el que ve el supervisor
# en la columna MOTIVO, así que va en lenguaje llano (antes eran códigos en
# mayúscula que nadie entendía). El detalle de qué meses mira cada chequeo
# va aparte, en la columna PERIODO.
MOTIVO_UNIVERSO = "Error de captura del mes"
MOTIVO_CAMBIO = "Comparativa contra el promedio de los últimos 3 meses"

# Umbral del motivo 2 — ver docstring del módulo. No se expone por línea de
# comandos por la misma razón que en precios/exhibiciones: si cada corrida
# pudiera elegirlo, dos meses dejarían de ser comparables.
UMBRAL_CAMBIO_PARTICIPACION = 0.20   # 20 puntos porcentuales de diferencia

# 'Origen de la línea de producto' distingue marca propia (Kenvue) de
# competencia en la encuesta cruda — el motivo 2 solo mira marcas propias.
ORIGEN_PROPIO = "PROPIO"

# Las 17 categorías oficiales del negocio (mismas que usa etl_sos.py). Sirven
# para normalizar 'Línea de producto', que trae ortografía distinta según la
# exportación de origen (ver docstring del módulo, punto 2).
LISTA_17_CATEGORIAS = [
    "SHAMPOO BEBE EN ADULTOS", "SHAMPOO BEBE", "JABONES SOLIDOS BEBE",
    "JABONES LIQUIDOS BEBE", "CREMAS CORPORALES BEBE", "ASEO DEL BEBE",
    "TOALLAS", "TAMPONES", "PROTECTORES", "PROTECCION SOLAR",
    "JABONES SOLIDOS ADULTOS", "JABONES LIQUIDOS ADULTOS", "FACIALES",
    "ENJUAGUES BUCALES MASIVOS", "ENJUAGUES BUCALES ESPECIALIZADOS",
    "ENJUAGUES BUCALES", "CREMAS CORPORALES ADULTO",
]

# KPIs de donde se saca el vínculo empleado → supervisor, en orden de
# preferencia. SOS primero porque son los propios gestores de espacios; el
# resto es respaldo para quien no aparezca ahí.
KPIS_CON_SUPERVISOR = [
    (paths.RUTA_CARPETA_SALIDAS_SOS, "SOS_KPIS.xlsx"),
    (paths.RUTA_CARPETA_SALIDAS_NP, "NO_PRESENCIA_KPIS.xlsx"),
    (paths.RUTA_CARPETA_SALIDAS_CIF, "KPIS_CIF.xlsx"),
    (paths.RUTA_CARPETA_SALIDAS_PRECIOS, "PRECIOS_KPIS.xlsx"),
]

TENANT_ID = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")


def eliminar_tildes(texto) -> str:
    if pd.isna(texto):
        return ""
    s = "".join(c for c in unicodedata.normalize("NFD", str(texto))
                if unicodedata.category(c) != "Mn")
    return s.upper().strip()


# ─────────────────────────────────────────────────────────────────────────────
# SHAREPOINT
# ─────────────────────────────────────────────────────────────────────────────

def obtener_token_azure() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"Azure rechazó la autenticación: {res.get('error_description')}")
    return res["access_token"]


def obtener_site_id(headers: dict) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{paths.SHAREPOINT_SITE_NAME}"
    res = requests.get(url, headers=headers).json()
    if "id" not in res:
        raise RuntimeError(f"No se encontró el sitio '{paths.SHAREPOINT_SITE_NAME}'")
    return res["id"]


def obtener_archivos_carpeta(headers: dict, site_id: str, ruta_carpeta: str) -> list[dict]:
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(ruta_carpeta)}:/children")
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        return []
    return r.json().get("value", [])


def leer_excel_cloud(headers: dict, site_id: str, ruta: str, descripcion: str,
                     obligatorio: bool = True) -> pd.DataFrame | None:
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(ruta)}:/content")
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        if obligatorio:
            raise FileNotFoundError(f"No se pudo leer {descripcion}: {ruta} ({r.status_code})")
        print(f"  ℹ️  No existe todavía {descripcion} — se creará nuevo.")
        return None
    print(f"  ✓ {descripcion} leído")
    return pd.read_excel(io.BytesIO(r.content))


def _formatear_porcentajes(ws, columnas: list[str], porcentaje: set[str]) -> None:
    """
    Los valores de `porcentaje` ya vienen multiplicados por 100 (ej. 50.0 =
    50%), así que el formato usa un '%' literal entre comillas — el formato
    nativo de Excel ('0.0%') multiplicaría por 100 otra vez.
    """
    from openpyxl.utils import get_column_letter
    for idx, col in enumerate(columnas, start=1):
        if col not in porcentaje:
            continue
        letra = get_column_letter(idx)
        for fila in range(2, ws.max_row + 1):
            ws[f"{letra}{fila}"].number_format = '0.0"%"'


def _embellecer_hoja(ws, df: pd.DataFrame) -> None:
    """
    Formato visual de la hoja: encabezado resaltado y congelado, columnas
    con ancho ajustado al contenido, bordes finos, alineación por tipo de
    dato (numérica a la derecha, texto a la izquierda — detectado por el
    dtype real de la columna, no por nombre, para que sirva igual en
    cualquier archivo de errores), filtro automático, y — si la hoja trae
    PARTICIPACION/PARTICIPACION_PROMEDIO_3M — esa columna en verde/rojo
    según subió o bajó, mismo código de color que ya usa el correo.
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    columnas = list(df.columns)
    n_filas = len(df)
    if n_filas == 0:
        return

    borde = Border(*(Side(style="thin", color="DDDDDD"),) * 4)
    relleno_encabezado = PatternFill("solid", fgColor="DCE6F1")
    alin_izq, alin_der = Alignment(horizontal="left"), Alignment(horizontal="right")
    alineaciones = [alin_der if pd.api.types.is_numeric_dtype(df[c]) else alin_izq
                   for c in columnas]

    for idx, col in enumerate(columnas, start=1):
        letra = get_column_letter(idx)
        celda_enc = ws[f"{letra}1"]
        celda_enc.font = Font(bold=True)
        celda_enc.fill = relleno_encabezado
        celda_enc.alignment = Alignment(horizontal="center", vertical="center")
        celda_enc.border = borde

        largo = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str)])
        tope = 90 if col in ("DIAGNOSTICO", "OBSERVACION_SUPERVISOR") else 45
        ws.column_dimensions[letra].width = min(max(largo + 2, 10), tope)

    # Una sola pasada fila por fila (iter_rows) en vez de direcciones de celda
    # repetidas — en archivos de miles de filas (encuestas crudas) la
    # diferencia de tiempo es real.
    for fila in ws.iter_rows(min_row=2, max_row=n_filas + 1, max_col=len(columnas)):
        for idx, celda in enumerate(fila):
            celda.border = borde
            celda.alignment = alineaciones[idx]

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    if "PARTICIPACION" in columnas and "PARTICIPACION_PROMEDIO_3M" in columnas:
        letra_part = get_column_letter(columnas.index("PARTICIPACION") + 1)
        letra_prom = get_column_letter(columnas.index("PARTICIPACION_PROMEDIO_3M") + 1)
        verde, rojo = Font(bold=True, color="0A7D1F"), Font(bold=True, color="CC0000")
        for fila in range(2, n_filas + 2):
            v, p = ws[f"{letra_part}{fila}"].value, ws[f"{letra_prom}{fila}"].value
            if isinstance(v, (int, float)) and isinstance(p, (int, float)):
                ws[f"{letra_part}{fila}"].font = verde if v > p else rojo


def subir_excel_cloud(headers: dict, site_id: str, carpeta: str, nombre: str,
                      hojas: dict[str, pd.DataFrame],
                      formato_porcentaje: dict[str, set[str]] | None = None,
                      embellecer: bool = False) -> None:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for hoja, df in hojas.items():
            df.to_excel(writer, sheet_name=hoja[:31], index=False)
            ws = writer.sheets[hoja[:31]]
            porcentaje = (formato_porcentaje or {}).get(hoja)
            if porcentaje:
                _formatear_porcentajes(ws, list(df.columns), porcentaje)
            if embellecer:
                _embellecer_hoja(ws, df)
    buffer.seek(0)
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(carpeta)}/{urllib.parse.quote(nombre)}:/content")
    cab = {**headers,
           "Content-Type": "application/vnd.openxmlformats-officedocument."
                           "spreadsheetml.sheet"}

    # 423 = alguien lo tiene abierto en el Excel de escritorio.
    espera = 20
    for intento in range(1, 4):
        r = requests.put(url, headers=cab, data=buffer.getvalue())
        if r.status_code in (200, 201):
            print(f"  ✅ Guardado en SharePoint: {carpeta}/{nombre}")
            return
        if r.status_code != 423:
            raise RuntimeError(f"Error al subir {nombre}: {r.status_code} - {r.text}")
        if intento < 3:
            print(f"  ⏳ {nombre} está abierto por alguien "
                  f"(intento {intento}/3) — reintento en {espera}s…")
            time.sleep(espera)
            espera *= 2

    raise RuntimeError(
        f"No se pudo guardar {nombre}: alguien lo tiene abierto en Excel de "
        f"escritorio y SharePoint no deja escribir encima (HTTP 423).\n"
        f"  · Los cálculos se hicieron bien; lo único que falló fue guardar.\n"
        f"  · Pedile que lo cierre, o que lo edite desde Excel Online.\n"
        f"  · Nada se perdió: las correcciones escritas siguen en el archivo."
    )


# ─────────────────────────────────────────────────────────────────────────────
# DETECCIÓN
# ─────────────────────────────────────────────────────────────────────────────

def _llave_por(cols: list[str], df: pd.DataFrame) -> pd.Series:
    partes = []
    for c in cols:
        col = df[c] if c in df.columns else pd.Series([""] * len(df), index=df.index)
        partes.append(col.astype(str).str.strip())
    return partes[0].str.cat(partes[1:], sep="|")


def _llave_caso(df: pd.DataFrame) -> pd.Series:
    """Llave del CASO (Empleado+PDV+Categoría) sobre las columnas visibles."""
    return _llave_por(COLS_LLAVE_CASO, df)


def _llave_fila(df: pd.DataFrame) -> pd.Series:
    """Llave de la FILA de detalle (+ Marca) sobre las columnas visibles."""
    return _llave_por(COLS_LLAVE_FILA, df)


def _llave_id(df: pd.DataFrame) -> pd.Series:
    """
    Llave para asignar/leer el ID_ERROR — depende del motivo:
      · MOTIVO_UNIVERSO: llave de CASO (varias marcas comparten un
        mismo ID_ERROR, son un solo caso visto marca por marca).
      · MOTIVO_CAMBIO: llave de FILA (cada marca es ya su
        propio caso, no hay nada que agrupar).
    """
    if "MOTIVO" in df.columns:
        es_cambio = df["MOTIVO"] == MOTIVO_CAMBIO
    else:
        es_cambio = pd.Series(False, index=df.index)
    return _llave_fila(df).where(es_cambio, _llave_caso(df))


def _decrementar_mes(mes: int, anio: int) -> tuple[int, int]:
    if mes == 1:
        return 12, anio - 1
    return mes - 1, anio


def _mes_anterior(spec: pr.PeriodoSpec) -> tuple[int, int]:
    return _decrementar_mes(spec.mes, spec.anio)


def _meses_anteriores(spec: pr.PeriodoSpec, n: int = 3) -> list[tuple[int, int]]:
    """Los N meses antes del periodo, del más antiguo al más reciente."""
    out = []
    mes, anio = spec.mes, spec.anio
    for _ in range(n):
        mes, anio = _decrementar_mes(mes, anio)
        out.append((mes, anio))
    return list(reversed(out))


def _llave_origen(df: pd.DataFrame) -> pd.Series:
    """
    Misma llave, pero sobre las columnas de la encuesta CRUDA ('PDV' en vez
    de 'PDV' visible ya coincide; la categoría se deriva con
    `_categoria_desde_linea`, no viene tal cual en 'Línea de producto').
    Cada fila de marca de un grupo en error recibe el mismo ID_ERROR — así,
    filtrando por ese ID en el archivo original, se ve el detalle marca por
    marca que arma esa diferencia.
    """
    categoria = df.apply(_categoria_desde_linea, axis=1).fillna("")
    return (df["Empleado"].astype(str).str.strip()
            .str.cat(df["PDV"].astype(str).str.strip(), sep="|")
            .str.cat(categoria, sep="|"))


def _categoria_desde_linea(row) -> str | None:
    """
    'Línea de producto' viene como "Categoría - Marca" (ej. "Toallas -
    Nosotras"). Se le resta la Marca de esa misma fila para recuperar la
    categoría, y se normaliza contra las 17 categorías oficiales — la
    ortografía de esa parte cambia según qué exportación trajo la fila (ej.
    "Enj. Bucales Masivos" en vez de "Enjuagues Bucales Masivos").
    """
    lin = str(row.get("Línea de producto", "") or "")
    marca = str(row.get("Marca", "") or "")
    base = lin
    if marca and lin.endswith(marca):
        base = lin[: -len(marca)].rstrip(" -")
    elif marca:
        partido = re.split(r"\s*-\s*" + re.escape(marca) + r"\s*$", lin)
        if len(partido) > 1:
            base = partido[0]

    t = re.sub(r"[^A-Z0-9\s]", " ", eliminar_tildes(base))
    t = re.sub(r"\bENJ\b\.?", "ENJUAGUES", t)   # "Enj." es abreviatura real en esta encuesta
    t = " ".join(t.split())

    if "SHAMPOO BEBE EN ADULTOS" in t:
        return "SHAMPOO BEBE EN ADULTOS"
    if re.search(r"\bSHAMPOO BEBE\b", t):
        return "SHAMPOO BEBE"
    for cat in LISTA_17_CATEGORIAS:
        if cat not in ("SHAMPOO BEBE", "SHAMPOO BEBE EN ADULTOS") and cat in t:
            return cat
    return None


def _detectar_columnas(df: pd.DataFrame) -> tuple[str, list[str]]:
    """
    Ubica las columnas de universo y de cm-por-marca por contenido, no por
    nombre exacto: la encuesta consolidada trae varias variantes según la
    exportación de origen — '¿Cuántos cms/CMS/CM tiene la marca?', con
    distinta tilde/mayúscula — mutuamente excluyentes fila a fila. Se
    devuelven TODAS para sumarlas; quedarse con la primera pierde ~2 de cada
    3 filas (la variante 'CM' sola es la más común, pero no la única).
    """
    limpias = {eliminar_tildes(c): c for c in df.columns}
    col_universo = next((limpias[c] for c in limpias if "UNIVERSO" in c and "CMS" in c), None)
    cols_marca_cm = [limpias[c] for c in limpias
                     if ("CM" in c.split() or "CMS" in c.split()) and "MARCA" in c]
    if not col_universo or not cols_marca_cm:
        raise KeyError(
            f"No se pudieron identificar las columnas de cm en la encuesta. "
            f"Columnas disponibles: {list(df.columns)}"
        )
    return col_universo, cols_marca_cm


def _detectar_universo(d: pd.DataFrame, col_universo: str, periodo_label: str) -> pd.DataFrame:
    """Motivo 1: la suma de cm de marcas supera el universo capturado."""
    casos = d.groupby(["Empleado", "PDV", "_CATEGORIA"], as_index=False).agg(
        UNIVERSO_CM=(col_universo, "mean"),
        SUMA_MARCAS_CM=("_CM", "sum"),
    )
    casos["DIFERENCIA"] = casos["UNIVERSO_CM"] - casos["SUMA_MARCAS_CM"]
    print(f"  Combinaciones Empleado × PDV × Categoría: {len(casos):,}")

    casos_error = casos[casos["DIFERENCIA"] < 0].copy()
    print(f"  → {len(casos_error):,} caso(s) con suma de marcas > universo")
    if casos_error.empty:
        return pd.DataFrame()

    casos_error["DIAGNOSTICO"] = casos_error.apply(
        lambda r: (f"Las marcas suman {r['SUMA_MARCAS_CM']:,.0f} cm pero el universo "
                   f"capturado es de {r['UNIVERSO_CM']:,.0f} cm "
                   f"(exceso de {abs(r['DIFERENCIA']):,.0f} cm)"),
        axis=1)

    # Detalle marca por marca de los casos en error. Si la misma marca tuvo
    # más de una captura en el mes, se suman (mismo criterio que
    # SUMA_MARCAS_CM del caso, para que las partes sigan sumando el total).
    detalle = (
        d.merge(casos_error[["Empleado", "PDV", "_CATEGORIA"]],
               on=["Empleado", "PDV", "_CATEGORIA"], how="inner")
         .groupby(["Empleado", "PDV", "_CATEGORIA", "Marca"], as_index=False)
         .agg(CM_MARCA=("_CM", "sum"))
    )
    detalle = detalle.merge(casos_error, on=["Empleado", "PDV", "_CATEGORIA"], how="left")
    detalle = detalle.rename(columns={"_CATEGORIA": "LINEA_PRODUCTO", "Marca": "MARCA"})
    detalle["MOTIVO"] = MOTIVO_UNIVERSO
    detalle["PERIODO"] = periodo_label
    return detalle


def _participacion_por_marca_propia(d: pd.DataFrame, col_universo: str) -> pd.DataFrame:
    """
    Participación de cada marca PROPIA = CM_MARCA / UNIVERSO_CM del caso
    (Empleado+PDV+Categoría). El universo es el de TODA la categoría (todas
    las marcas, propias y competencia) — la participación es la porción de
    ese total que corresponde a la marca propia. Se excluyen casos con
    universo o cm en 0 (no hay participación que calcular).
    """
    universo_caso = d.groupby(["Empleado", "PDV", "_CATEGORIA"], as_index=False).agg(
        UNIVERSO_CM=(col_universo, "mean"))
    propias = d[d["_ORIGEN"] == ORIGEN_PROPIO]
    cm_marca = propias.groupby(["Empleado", "PDV", "_CATEGORIA", "Marca"], as_index=False).agg(
        CM_MARCA=("_CM", "sum"))
    m = cm_marca.merge(universo_caso, on=["Empleado", "PDV", "_CATEGORIA"], how="left")
    m = m[(m["CM_MARCA"] > 0) & (m["UNIVERSO_CM"] > 0)].copy()
    m["PARTICIPACION"] = m["CM_MARCA"] / m["UNIVERSO_CM"]
    return m


def _detectar_cambio_participacion(d: pd.DataFrame, col_universo: str,
                                   piezas_prev: list[tuple[str, pd.DataFrame]],
                                   periodo_label: str) -> pd.DataFrame:
    """
    Motivo 2: la participación de una marca PROPIA (ver
    `_participacion_por_marca_propia`) cambió mucho vs. el promedio de los
    meses anteriores disponibles (hasta 3). `piezas_prev` trae, por cada mes
    anterior que sí existe, su (col_universo, dataframe ya preparado). Por
    marca, el promedio solo cuenta los meses donde esa marca tuvo
    participación > 0 — que entre o salga del surtido no es, por sí sola,
    un error.
    """
    actual = _participacion_por_marca_propia(d, col_universo)

    piezas = [_participacion_por_marca_propia(d_p, col_u_p)[
                  ["Empleado", "PDV", "_CATEGORIA", "Marca",
                   "CM_MARCA", "UNIVERSO_CM", "PARTICIPACION"]]
              for col_u_p, d_p in piezas_prev]
    if not piezas:
        return pd.DataFrame()
    historico = pd.concat(piezas, ignore_index=True)
    promedio = historico.groupby(["Empleado", "PDV", "_CATEGORIA", "Marca"], as_index=False).agg(
        CM_MARCA_PROMEDIO_3M=("CM_MARCA", "mean"),
        UNIVERSO_CM_PROMEDIO_3M=("UNIVERSO_CM", "mean"),
        PARTICIPACION_PROMEDIO_3M=("PARTICIPACION", "mean"),
        MESES_CON_DATO=("PARTICIPACION", "size"))

    m = actual.merge(promedio, on=["Empleado", "PDV", "_CATEGORIA", "Marca"], how="inner")
    print(f"  Marcas propias con participación este mes y en algún mes anterior: {len(m):,}")

    m["_dif_pp"] = (m["PARTICIPACION"] - m["PARTICIPACION_PROMEDIO_3M"]).abs()
    m = m[m["_dif_pp"] > UMBRAL_CAMBIO_PARTICIPACION].copy()
    print(f"  → {len(m):,} caso(s) con variación > "
          f"{UMBRAL_CAMBIO_PARTICIPACION * 100:.0f} puntos en participación")
    if m.empty:
        return pd.DataFrame()

    m["DIAGNOSTICO"] = m.apply(
        lambda r: (f"{'Subió' if r['PARTICIPACION'] > r['PARTICIPACION_PROMEDIO_3M'] else 'Bajó'} "
                   f"de {r['PARTICIPACION_PROMEDIO_3M'] * 100:,.1f}% de participación en promedio "
                   f"los últimos {int(r['MESES_CON_DATO'])} mes(es) a "
                   f"{r['PARTICIPACION'] * 100:,.1f}% este mes "
                   f"({r['_dif_pp'] * 100:,.1f} puntos de diferencia)"),
        axis=1)

    m["PARTICIPACION"] = (m["PARTICIPACION"] * 100).round(1)
    m["PARTICIPACION_PROMEDIO_3M"] = (m["PARTICIPACION_PROMEDIO_3M"] * 100).round(1)
    m["CM_MARCA_PROMEDIO_3M"] = m["CM_MARCA_PROMEDIO_3M"].round(1)
    m["UNIVERSO_CM_PROMEDIO_3M"] = m["UNIVERSO_CM_PROMEDIO_3M"].round(1)

    m = m.rename(columns={"_CATEGORIA": "LINEA_PRODUCTO", "Marca": "MARCA"})
    m["MOTIVO"] = MOTIVO_CAMBIO
    m["PERIODO"] = periodo_label
    return m.drop(columns=["_dif_pp", "MESES_CON_DATO"])


def _preparar(df: pd.DataFrame, col_universo: str, cols_marca_cm: list[str],
              mes: int, anio: int) -> pd.DataFrame:
    """
    Deja la encuesta cruda lista para agrupar: normaliza textos, calcula
    `_CM` (suma de las variantes de cm-por-marca), deriva `_CATEGORIA`, y
    acota al periodo (mes/año) por si el archivo todavía trae otros meses.
    """
    d = df.copy()
    d["Empleado"] = d["Empleado"].astype(str).str.strip()
    d["PDV"] = d["PDV"].astype(str).str.strip()
    d["Marca"] = d["Marca"].astype(str).str.strip()
    d[col_universo] = pd.to_numeric(d[col_universo], errors="coerce")
    d["_CM"] = d[cols_marca_cm].apply(pd.to_numeric, errors="coerce").sum(axis=1, skipna=True)
    d["_CATEGORIA"] = d.apply(_categoria_desde_linea, axis=1)
    d = d.dropna(subset=["_CATEGORIA"])
    if "Origen de la línea de producto" in d.columns:
        d["_ORIGEN"] = d["Origen de la línea de producto"].apply(eliminar_tildes)
    else:
        d["_ORIGEN"] = ""

    if "Mes del año" in d.columns and "Año" in d.columns:
        n_antes = len(d)
        d = d[(pd.to_numeric(d["Mes del año"], errors="coerce") == mes)
              & (pd.to_numeric(d["Año"], errors="coerce") == anio)]
        if len(d) != n_antes:
            print(f"    ℹ️  Acotado a {mes:02d}/{anio}: {n_antes:,} → {len(d):,} filas "
                  f"(el archivo todavía traía otros meses).")
    return d


def detectar(df: pd.DataFrame, spec: pr.PeriodoSpec,
             dfs_prev: list[pd.DataFrame | None]) -> pd.DataFrame:
    """
    Corre los dos motivos (ver docstring del módulo) y devuelve el detalle
    combinado — una fila por marca, con MOTIVO indicando cuál disparó cada
    fila.

    `dfs_prev` son las encuestas consolidadas de los 3 meses anteriores, del
    más antiguo al más reciente (mismo orden que `_meses_anteriores`), una
    por archivo aparte. Los que no existan van como None; si ninguno existe,
    no se puede hacer la comparativa y solo corre el motivo 1.
    """
    col_universo, cols_marca_cm = _detectar_columnas(df)

    n_sin_cat = int(df.apply(_categoria_desde_linea, axis=1).isna().sum())
    if n_sin_cat:
        print(f"  ⚠️  {n_sin_cat:,} fila(s) sin categoría reconocible — se excluyen del chequeo.")

    d = _preparar(df, col_universo, cols_marca_cm, spec.mes, spec.anio)
    print(f"  Filas del periodo {spec.mes:02d}/{spec.anio}: {len(d):,}")
    if d.empty:
        raise ValueError(f"La encuesta de {spec.mes:02d}/{spec.anio} quedó sin filas.")

    meses_prev = _meses_anteriores(spec, 3)
    piezas_prev: list[tuple[str, pd.DataFrame]] = []
    for (mes_p, anio_p), df_p in zip(meses_prev, dfs_prev):
        if df_p is None or df_p.empty:
            continue
        col_u_p, cols_m_p = _detectar_columnas(df_p)
        d_p = _preparar(df_p, col_u_p, cols_m_p, mes_p, anio_p)
        if not d_p.empty:
            piezas_prev.append((col_u_p, d_p))

    etiqueta_mes = f"{spec.mes_str} {spec.anio}"
    (m1, a1), (m2, a2) = meses_prev[0], meses_prev[-1]
    etiqueta_3m = (f"{pr.MESES_ES[m1]}-{pr.MESES_ES[m2]} {a1}" if a1 == a2
                   else f"{pr.MESES_ES[m1]} {a1} - {pr.MESES_ES[m2]} {a2}")

    print("\n  Motivo 1 — suma de marcas > universo:")
    detalle_universo = _detectar_universo(d, col_universo, etiqueta_mes)

    if piezas_prev:
        print(f"\n  Motivo 2 — participación vs. promedio {etiqueta_3m} "
              f"({len(piezas_prev)} de 3 mes(es) disponible(s)):")
        detalle_cambio = _detectar_cambio_participacion(
            d, col_universo, piezas_prev, f"{etiqueta_mes} vs. promedio {etiqueta_3m}")
    else:
        print("\n  Motivo 2 — se salta (sin archivos de meses anteriores).")
        detalle_cambio = pd.DataFrame()

    return pd.concat([detalle_universo, detalle_cambio], ignore_index=True, sort=False)


def mapa_supervisores(headers: dict, site_id: str) -> dict:
    """Empleado → supervisor, combinando los KPIs que traen SUPERVISOR_LIDER."""
    mapa: dict = {}
    for carpeta, archivo in KPIS_CON_SUPERVISOR:
        k = leer_excel_cloud(headers, site_id, f"{carpeta}/{archivo}",
                             archivo, obligatorio=False)
        if k is None or "SUPERVISOR_LIDER" not in k.columns or "NOMBRE" not in k.columns:
            continue
        sup = k["SUPERVISOR_LIDER"].astype(str).str.upper().str.strip()
        sup = sup.where(~sup.isin(["", "NAN", "NONE"]), "")
        nom = k["NOMBRE"].astype(str).str.upper().str.strip()
        nuevos = 0
        for n, s in zip(nom, sup):
            if s and n and n not in mapa:
                mapa[n] = s
                nuevos += 1
        print(f"    {archivo}: +{nuevos} personas")
    return mapa


def asignar_supervisor(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    df = df.copy()
    df["SUPERVISOR_LIDER"] = (
        df["Empleado"].astype(str).str.upper().str.strip()
          .map(mapa).replace("", pd.NA).fillna("(sin cruzar)"))
    n = int((df["SUPERVISOR_LIDER"] == "(sin cruzar)").sum())
    if n:
        nombres = sorted(df.loc[df["SUPERVISOR_LIDER"] == "(sin cruzar)", "Empleado"].unique())
        print(f"  ⚠️  {n} fila(s) sin supervisor: el empleado no aparece en ningún KPI "
              f"({', '.join(nombres[:4])}{'…' if len(nombres) > 4 else ''}). "
              f"Se agrupan bajo '(sin cruzar)' para que no se pierdan.")
    return df


def asignar_ids(nuevo: pd.DataFrame, previo: pd.DataFrame | None,
                prefijo: str, spec: pr.PeriodoSpec) -> pd.Series:
    """
    ID legible y ESTABLE entre corridas: ESP-202608-001. La llave depende
    del motivo (ver `_llave_id`): en MOTIVO_UNIVERSO varias filas de
    marca comparten un mismo ID_ERROR (mismo caso visto marca por marca);
    en MOTIVO_CAMBIO cada fila ya es su propio caso. Mismo
    mecanismo que etl_precios_errores.py / etl_exhibiciones_errores.py,
    adaptado para no numerar dos veces un caso que aparece en varias filas.
    """
    base = f"{prefijo}-{spec.anio}{spec.mes:02d}-"
    ya: dict[str, str] = {}
    maximo = 0
    if previo is not None and not previo.empty and "ID_ERROR" in previo.columns:
        for k, i in zip(_llave_id(previo), previo["ID_ERROR"].astype(str)):
            if not i or i == "nan":
                continue
            ya.setdefault(k, i)
            try:
                maximo = max(maximo, int(i.rsplit("-", 1)[-1]))
            except ValueError:
                pass

    llave_nuevo = _llave_id(nuevo)
    mapa: dict[str, str] = {}
    nuevos = 0
    for k in llave_nuevo.unique():
        if k in ya:
            mapa[k] = ya[k]
        else:
            maximo += 1
            nuevos += 1
            mapa[k] = f"{base}{maximo:03d}"
    if ya:
        print(f"  🔖 IDs: {len(mapa) - nuevos} conservados, {nuevos} nuevos "
              f"({len(nuevo)} fila(s) de detalle)")
    return llave_nuevo.map(mapa)


def _nombre_hoja(headers: dict, site_id: str, ruta: str) -> str:
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(ruta)}:/content")
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        return "Sheet1"
    return pd.ExcelFile(io.BytesIO(r.content)).sheet_names[0]


def escribir_id_en_origen(headers: dict, site_id: str, carpeta: str, archivo: str,
                          err: pd.DataFrame, spec: pr.PeriodoSpec) -> None:
    """
    Devuelve el ID_ERROR a la encuesta consolidada cruda, para poder ubicar
    ahí el detalle marca por marca de cada caso.

    Solo se escribe el motivo MOTIVO_UNIVERSO: ese cruza por caso
    completo (Empleado+PDV+Categoría), que es la misma llave con la que se
    identifica cada fila cruda de esa categoría. El motivo
    MOTIVO_CAMBIO identifica cada fila por marca — cruzarlo
    acá pisaría el ID_ERROR de filas de OTRAS marcas de la misma categoría
    que no tienen nada que ver. Ese motivo queda solo en el archivo de
    salida, donde cada fila ya distingue la marca.

    Igual que en exhibiciones, este archivo ACUMULA varios periodos (ver
    docstring del módulo) — el mes que se está procesando se reescribe
    entero; los demás conservan el ID_ERROR que ya tuvieran.

    OJO CON EL ORDEN: `etl_sos.py` regenera este archivo desde cero en cada
    corrida, y con eso se lleva puesta esta columna. El workflow de errores
    debe correr DESPUÉS del ETL de SOS, igual que con exhibiciones y precios.
    """
    ruta = f"{carpeta}/{archivo}"
    df = leer_excel_cloud(headers, site_id, ruta, archivo, obligatorio=False)
    if df is None or df.empty:
        return

    hoja = _nombre_hoja(headers, site_id, ruta)

    if "Mes del año" in df.columns and "Año" in df.columns:
        del_periodo = ((pd.to_numeric(df["Mes del año"], errors="coerce") == spec.mes)
                       & (pd.to_numeric(df["Año"], errors="coerce") == spec.anio))
    else:
        del_periodo = pd.Series(True, index=df.index)

    anterior = (df["ID_ERROR"].fillna("").astype(str).replace("nan", "")
               if "ID_ERROR" in df.columns else pd.Series("", index=df.index))
    df["ID_ERROR"] = anterior.where(~del_periodo, "")

    err_universo = err[err["MOTIVO"] == MOTIVO_UNIVERSO] if "MOTIVO" in err.columns else err

    # `err_universo` ya trae la categoría limpia en LINEA_PRODUCTO (columnas
    # visibles, llave de CASO); `df` es la encuesta cruda y necesita
    # _llave_origen para derivarla de 'Línea de producto' - 'Marca'. Se usa
    # la llave de caso (no la de fila) porque todas las marcas de un mismo
    # caso van con el mismo ID_ERROR.
    nuevos = (dict(zip(_llave_caso(err_universo), err_universo["ID_ERROR"].astype(str)))
             if not err_universo.empty else {})
    df.loc[del_periodo, "ID_ERROR"] = _llave_origen(df[del_periodo]).map(nuevos).fillna("")
    n = int((df["ID_ERROR"] != "").sum())

    subir_excel_cloud(headers, site_id, carpeta, archivo, {hoja: df}, embellecer=True)
    print(f"  🔗 ID_ERROR escrito en {archivo}: {n} fila(s) marcadas")


def conservar_correcciones(nuevo: pd.DataFrame, previo: pd.DataFrame | None) -> pd.DataFrame:
    """
    Arrastra lo que el supervisor ya escribió en la corrida anterior. Se
    conserva por FILA (Empleado+PDV+Categoría+Marca): la corrección de una
    marca puntual sigue a esa marca, no se mezcla con las demás del caso.
    """
    for c in COLS_RESPUESTA:
        nuevo[c] = ""
    if previo is None or previo.empty:
        return nuevo
    if any(c not in previo.columns for c in COLS_LLAVE_FILA):
        print("  ⚠️  El archivo anterior no trae las columnas llave; no se pueden "
              "conservar las correcciones. Se genera limpio.")
        return nuevo
    llave_previo, llave_nuevo = _llave_fila(previo), _llave_fila(nuevo)
    for c in COLS_RESPUESTA:
        if c not in previo.columns:
            continue
        mapa = dict(zip(llave_previo, previo[c].fillna("")))
        nuevo[c] = llave_nuevo.map(mapa).fillna("")
    n = int((nuevo["OBSERVACION_SUPERVISOR"].astype(str).str.strip() != "").sum())
    if n:
        print(f"  ✓ Se conservaron {n} corrección(es) que ya había escrito el supervisor.")
    return nuevo


# ─────────────────────────────────────────────────────────────────────────────
# PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

# G=UNIVERSO_CM y H=CM_MARCA quedan fijos ahí a propósito (pedido explícito).
COLS_SALIDA = [
    "ID_ERROR",
    "Empleado", "PDV", "LINEA_PRODUCTO", "MARCA",
    "SUPERVISOR_LIDER",
    "UNIVERSO_CM", "CM_MARCA", "SUMA_MARCAS_CM",
    "UNIVERSO_CM_PROMEDIO_3M", "CM_MARCA_PROMEDIO_3M",
    "PARTICIPACION", "PARTICIPACION_PROMEDIO_3M",
    "MOTIVO", "PERIODO", "DIAGNOSTICO",
    "CM_MARCA_CORREGIDO", "UNIVERSO_CORREGIDO", "OBSERVACION_SUPERVISOR",
]


def _nombre_encuesta_periodo(mes: int, anio: int) -> str:
    return f"Encuesta_Sos_Consolidada_{pr.MESES_ES[mes].upper()}_{anio}.xlsx"


def _nombre_encuesta(spec: pr.PeriodoSpec) -> str:
    return _nombre_encuesta_periodo(spec.mes, spec.anio)


def run(spec: pr.PeriodoSpec) -> int:
    print("\n" + "=" * 60)
    print(f"  ERRORES DE ESPACIOS (SOS) — {spec.etiqueta}")
    print(f"  Motivo 1: suma de cm de marcas > universo capturado (imposible)")
    print(f"  Motivo 2: participación de marca propia vs. promedio de los 3 "
          f"meses anteriores, variación > {UMBRAL_CAMBIO_PARTICIPACION * 100:.0f} puntos")
    print("=" * 60)

    headers = {"Authorization": f"Bearer {obtener_token_azure()}"}
    site_id = obtener_site_id(headers)
    carpeta_bases = paths.RUTA_CARPETA_BASES_SOS
    carpeta_salidas = paths.RUTA_CARPETA_SALIDAS_SOS
    archivo_origen = _nombre_encuesta(spec)
    nombre_salida = f"ERRORES_ESPACIOS_{spec.mes_str_upper}_{spec.anio}.xlsx"

    print("\nLeyendo insumos:")
    df = leer_excel_cloud(headers, site_id, f"{carpeta_bases}/{archivo_origen}", archivo_origen)
    # La comparativa del motivo 2 se hace contra los 3 meses anteriores, cada
    # uno en su propio archivo, no contra una "sección" del de este mes.
    dfs_prev = []
    for mes_p, anio_p in _meses_anteriores(spec, 3):
        archivo_p = _nombre_encuesta_periodo(mes_p, anio_p)
        dfs_prev.append(leer_excel_cloud(
            headers, site_id, f"{carpeta_bases}/{archivo_p}",
            f"{archivo_p} (para la comparativa)", obligatorio=False))
    previo = leer_excel_cloud(headers, site_id, f"{carpeta_salidas}/{nombre_salida}",
                              f"{nombre_salida} (corrida anterior)", obligatorio=False)

    print("\nArmando el mapa empleado → supervisor:")
    mapa = mapa_supervisores(headers, site_id)
    print(f"    total: {len(mapa)} personas con supervisor conocido")

    print("\nDetectando:")
    err = detectar(df, spec, dfs_prev)

    if err.empty:
        print("  ✓ Sin errores en ninguno de los dos motivos.")
        escribir_id_en_origen(headers, site_id, carpeta_bases, archivo_origen, err, spec)
        subir_excel_cloud(headers, site_id, carpeta_salidas, nombre_salida,
                          {"Errores": pd.DataFrame(columns=COLS_SALIDA)})
        return 0

    err = asignar_supervisor(err, mapa)
    err["ID_ERROR"] = asignar_ids(err, previo, "ESP", spec)
    err = conservar_correcciones(err, previo)
    # Dentro de cada caso (mismo ID_ERROR), la marca con más cm queda primero
    # — suele ser la candidata más probable a estar mal tipeada.
    err = err.sort_values(["SUPERVISOR_LIDER", "Empleado", "ID_ERROR", "CM_MARCA"],
                          ascending=[True, True, True, False])

    casos_unicos = err.drop_duplicates("ID_ERROR")
    n_casos = len(casos_unicos)
    n_universo = int((casos_unicos["MOTIVO"] == MOTIVO_UNIVERSO).sum())
    n_cambio = n_casos - n_universo
    print(f"\n  {n_casos} caso(s) a revisar ({len(err)} fila(s) de detalle, marca por marca)")
    print(f"  {n_universo} por universo superado, {n_cambio} por cambio abrupto")
    print(f"  Repartidos en {err['SUPERVISOR_LIDER'].nunique()} supervisor(es)")

    resumen = (casos_unicos
                  .groupby("SUPERVISOR_LIDER", as_index=False)
                  .agg(CASOS=("ID_ERROR", "size"),
                       GESTORES=("Empleado", "nunique"),
                       UNIVERSO=("MOTIVO", lambda s: int((s == MOTIVO_UNIVERSO).sum())),
                       CAMBIO=("MOTIVO", lambda s: int((s == MOTIVO_CAMBIO).sum())))
                  .sort_values("CASOS", ascending=False))
    print()
    for _, r in resumen.head(10).iterrows():
        print(f"     {str(r['SUPERVISOR_LIDER'])[:38]:38} {r['CASOS']:>3} "
              f"({r['UNIVERSO']} universo, {r['CAMBIO']} cambio abrupto)")

    escribir_id_en_origen(headers, site_id, carpeta_bases, archivo_origen, err, spec)

    err = err.reindex(columns=COLS_SALIDA)

    print("\nGuardando:")
    subir_excel_cloud(headers, site_id, carpeta_salidas, nombre_salida, {"Errores": err},
                      formato_porcentaje={"Errores": {"PARTICIPACION", "PARTICIPACION_PROMEDIO_3M"}},
                      embellecer=True)
    return len(err)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--mes", type=int, required=True, choices=range(1, 13))
    ap.add_argument("--anio", type=int, required=True)
    args = ap.parse_args()
    run(pr.resolver(int(args.mes), int(args.anio)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
