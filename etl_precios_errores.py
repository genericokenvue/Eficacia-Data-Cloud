"""
etl_errores_precios.py
──────────────────────
Detecta precios mal capturados en `ANALISIS_PRECIOS_<MES>_<AÑO>.xlsx` y arma
un archivo para que cada supervisor los revise y corrija.

CÓMO DETECTA
────────────
Compara cada precio contra la MEDIANA de su propio SKU:

    ratio = PRECIO_REGULAR / mediana_del_SKU

    ratio fuera de [1/3, 3]  →  severidad ALTA
    ratio fuera de [1/2, 2]  →  severidad MEDIA

Los errores reales son casi siempre de dígito: un cero de más o de menos.
Medido sobre agosto 2026 (14.289 capturas con PRESENCIA=1 y precio > 0):

    mediana del ratio        1.00
    percentil 99             1.24     ← un precio legítimo no se aleja más
    percentil 99.9           1.85
    máximo                   9.42     ← acá viven los errores

Con umbral 2x salen 46 filas (0.3%), repartidas en 15 supervisores: mediana de
3 por supervisor, máximo 7.

POR QUÉ MEDIANA Y NO MEDIA NI PERCENTIL
───────────────────────────────────────
· La MEDIA se contamina con el propio error que se busca: un precio 10 veces
  mayor arrastra el promedio y sube el umbral justo cuando no debería.
· El PERCENTIL 90 marca, por definición, al 10% de las capturas. Sobre estos
  datos serían 1.131 filas de precios correctos. Un reporte así se ignora a la
  segunda semana.
La mediana no se mueve aunque haya outliers, que es exactamente lo que hace
falta para detectarlos.

SKUs CON POCAS CAPTURAS
───────────────────────
Con menos de MIN_CAPTURAS registros la mediana no es confiable, así que esos
SKUs no se evalúan. Quedan listados aparte para que no desaparezcan en
silencio.

EL ARCHIVO DE SALIDA ES TAMBIÉN EL FORMULARIO
─────────────────────────────────────────────
`ERRORES_PRECIOS_<MES>_<AÑO>.xlsx` trae dos columnas vacías,
PRECIO_CORREGIDO y OBSERVACION_SUPERVISOR, que el supervisor llena en Excel
Online. En cada corrida se vuelve a leer ese mismo archivo y **se conservan
las correcciones ya escritas**: si se regenerara de cero, el supervisor
perdería su trabajo y no volvería a llenarlo nunca más.

USO
───
    python etl_errores_precios.py --mes 8 --anio 2026
    python etl_errores_precios.py --mes 8 --anio 2026 --umbral 3
"""
from __future__ import annotations

import argparse
import io
import os
import sys
import time
import urllib.parse

import msal
import pandas as pd
import requests
from dotenv import load_dotenv

import paths
import periodo_resolver as pr

load_dotenv()

# ─────────────────────────────────────────────────────────────────────────────
# PARÁMETROS DE DETECCIÓN
# ─────────────────────────────────────────────────────────────────────────────
UMBRAL_MEDIO = 2.0      # ratio fuera de [1/2, 2]  → revisar
UMBRAL_ALTO = 3.0       # ratio fuera de [1/3, 3]  → casi seguro un error
MIN_CAPTURAS = 5        # SKUs con menos capturas no se evalúan

# Columnas que el supervisor llena. Se conservan entre corridas.
COLS_RESPUESTA = ["PRECIO_CORREGIDO", "OBSERVACION_SUPERVISOR"]

# Identifica una captura de forma estable entre corridas, para poder arrastrar
# lo que el supervisor ya escribió aunque el archivo se regenere.
COLS_LLAVE = ["ID del PDV", "CODIGO_SKU", "Fecha de la encuesta", "Empleado"]

TENANT_ID = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")


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


def leer_excel_cloud(headers: dict, site_id: str, ruta: str, descripcion: str,
                     obligatorio: bool = True) -> pd.DataFrame | None:
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(ruta)}:/content")
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        if obligatorio:
            raise FileNotFoundError(f"No se pudo leer {descripcion}: {ruta} ({r.status_code})")
        print(f"  ℹ️  No existe todavía {descripcion} ({ruta}) — se creará nuevo.")
        return None
    print(f"  ✓ {descripcion} leído")
    return pd.read_excel(io.BytesIO(r.content))


def subir_excel_cloud(headers: dict, site_id: str, carpeta: str, nombre: str,
                      hojas: dict[str, pd.DataFrame]) -> None:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for hoja, df in hojas.items():
            df.to_excel(writer, sheet_name=hoja[:31], index=False)
    buffer.seek(0)
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(carpeta)}/{urllib.parse.quote(nombre)}:/content")
    cab = {**headers,
           "Content-Type": "application/vnd.openxmlformats-officedocument."
                           "spreadsheetml.sheet"}

    # SharePoint devuelve 423 (locked) cuando alguien tiene el archivo abierto
    # en el Excel de escritorio. Acá eso NO es raro: el archivo es justamente
    # el formulario que los supervisores editan, así que se reintenta un rato
    # antes de rendirse, y si igual no se puede se explica qué pasó en vez de
    # soltar un traceback que nadie va a saber leer.
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
        f"  · Pedile a quien lo tenga abierto que lo cierre, o que lo edite "
        f"desde Excel Online (ahí no bloquea), y volvé a correr.\n"
        f"  · Nada se perdió: las correcciones que ya estaban escritas siguen "
        f"en el archivo."
    )


# ─────────────────────────────────────────────────────────────────────────────
# DETECCIÓN
# ─────────────────────────────────────────────────────────────────────────────

# Las mismas cuatro columnas llave, pero con el nombre que tienen ya en la
# hoja de salida (Fecha de la encuesta se renombró a FECHA).
COLS_LLAVE_VISIBLES = ["ID PDV", "CODIGO_SKU", "FECHA", "Empleado"]


def _llave_visible(df: pd.DataFrame) -> pd.Series:
    """Llave compuesta a partir de las columnas ya renombradas de la hoja."""
    partes = [df[c].astype(str).str.strip() if c in df.columns
              else pd.Series([""] * len(df), index=df.index)
              for c in COLS_LLAVE_VISIBLES]
    return partes[0].str.cat(partes[1:], sep="|")


def _llave(df: pd.DataFrame) -> pd.Series:
    """Identificador estable de una captura, para cruzar entre corridas."""
    partes = []
    for c in COLS_LLAVE:
        col = df[c] if c in df.columns else pd.Series([""] * len(df), index=df.index)
        partes.append(col.astype(str).str.strip())
    return partes[0].str.cat(partes[1:], sep="|")


def detectar(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Devuelve (errores, skus_sin_evaluar).

    Solo se miran capturas con PRESENCIA = 1 y precio > 0: PRESENCIA = 0
    significa que el producto no estaba y el precio viene en 0 por definición,
    no es un error de captura.
    """
    d = df.copy()
    d["PRECIO_REGULAR"] = pd.to_numeric(d["PRECIO_REGULAR"], errors="coerce")
    d = d[(pd.to_numeric(d["PRESENCIA"], errors="coerce") == 1)
          & (d["PRECIO_REGULAR"] > 0)].copy()
    print(f"  Capturas evaluables (PRESENCIA=1, precio>0): {len(d):,}")

    g = d.groupby("CODIGO_SKU")["PRECIO_REGULAR"]
    d["MEDIANA_SKU"] = g.transform("median")
    # La media va solo como referencia visible, NO se usa para decidir: se
    # contamina con el mismo error que se busca. En el SKU 7702031318316, por
    # ejemplo, un precio de 356.750 contra una mediana de 37.890 empuja la
    # media a 38.739. Tenerlas lado a lado deja ver ese efecto.
    d["MEDIA_SKU"] = g.transform("mean").round(0)
    d["CAPTURAS_SKU"] = g.transform("size")
    # Se guarda el ratio sin redondear para comparar y describir; el redondeo
    # es solo para mostrarlo. Con .round(2), un precio muy chico contra una
    # mediana grande daba 0.0 exacto y el texto del diagnóstico dividía por cero.
    d["_RATIO_EXACTO"] = d["PRECIO_REGULAR"] / d["MEDIANA_SKU"]
    d["RATIO"] = d["_RATIO_EXACTO"].round(2)

    pocos = d[d["CAPTURAS_SKU"] < MIN_CAPTURAS]
    sin_evaluar = (
        pocos.groupby(["CODIGO_SKU", "NOMBRE_PRODUCTO"], as_index=False)
             .agg(CAPTURAS=("PRECIO_REGULAR", "size"),
                  PRECIO_MIN=("PRECIO_REGULAR", "min"),
                  PRECIO_MAX=("PRECIO_REGULAR", "max"))
    )
    if not sin_evaluar.empty:
        print(f"  ℹ️  {len(sin_evaluar)} SKU(s) con menos de {MIN_CAPTURAS} capturas: "
              f"no se evalúan (la mediana no sería confiable).")

    d = d[d["CAPTURAS_SKU"] >= MIN_CAPTURAS].copy()

    ex = d["_RATIO_EXACTO"]
    fuera_medio = (ex > UMBRAL_MEDIO) | (ex < 1 / UMBRAL_MEDIO)
    fuera_alto = (ex > UMBRAL_ALTO) | (ex < 1 / UMBRAL_ALTO)

    err = d[fuera_medio].copy()
    err["SEVERIDAD"] = "MEDIA"
    err.loc[fuera_alto, "SEVERIDAD"] = "ALTA"

    def _texto(r: float) -> str:
        # Dice "promedio" aunque el cálculo use la MEDIANA. Es a propósito: en
        # campo "promedio" se entiende de una, y la distinción estadística no
        # le cambia nada a quien tiene que corregir el precio. Lo que se
        # compara sigue siendo la mediana, por lo explicado más arriba.
        if r > 1:
            return f"{r:.1f}x por ENCIMA del promedio del producto"
        if r > 0:
            return f"{1/r:.1f}x por DEBAJO del promedio del producto"
        return "precio en cero o inválido"

    err["DIAGNOSTICO"] = err["_RATIO_EXACTO"].apply(_texto)

    # Todos los registros evaluados, marcando cuál falla. La hoja "Errores"
    # es la accionable, pero acá queda la foto completa: sirve para ver el
    # rango normal de cada producto y para confiar en que lo marcado es
    # realmente la excepción y no un criterio que barre con todo.
    todos = d.copy()
    todos["TIENE_ERROR"] = "NO"
    todos.loc[fuera_medio, "TIENE_ERROR"] = "SI"
    todos["SEVERIDAD"] = ""
    todos.loc[fuera_medio, "SEVERIDAD"] = "MEDIA"
    todos.loc[fuera_alto, "SEVERIDAD"] = "ALTA"
    todos["DIAGNOSTICO"] = todos["_RATIO_EXACTO"].apply(_texto)
    todos.loc[~fuera_medio, "DIAGNOSTICO"] = "dentro del rango normal"

    return (err.drop(columns=["_RATIO_EXACTO"]),
            sin_evaluar,
            todos.drop(columns=["_RATIO_EXACTO"]))


def asignar_supervisor(err: pd.DataFrame, kpis: pd.DataFrame,
                       silencioso: bool = False) -> pd.DataFrame:
    """
    Agrega SUPERVISOR_LIDER cruzando por nombre del empleado.

    El vínculo gestor→supervisor vive en el Plan de Trabajo; `PRECIOS_KPIS.xlsx`
    ya lo trae resuelto por gestor, así que se lee de ahí en vez de rehacer el
    cruce.
    """
    if kpis is None or kpis.empty:
        err["SUPERVISOR_LIDER"] = "(sin cruzar)"
        return err
    sup = kpis["SUPERVISOR_LIDER"].astype(str).str.upper().str.strip()
    # Un supervisor vacío en el KPI llega como "NAN" al pasar por astype(str),
    # y sin esto salía en el reporte como si fuera el nombre de alguien.
    sup = sup.where(~sup.isin(["", "NAN", "NONE"]), "")
    mapa = {n: s for n, s in zip(
        kpis["NOMBRE"].astype(str).str.upper().str.strip(), sup) if s}
    err["SUPERVISOR_LIDER"] = (
        err["Empleado"].astype(str).str.upper().str.strip()
           .map(mapa).replace("", pd.NA).fillna("(sin cruzar)")
    )
    n = int((err["SUPERVISOR_LIDER"] == "(sin cruzar)").sum())
    if n and not silencioso:
        nombres = sorted(err.loc[err["SUPERVISOR_LIDER"] == "(sin cruzar)", "Empleado"].unique())
        print(f"  ⚠️  {n} fila(s) sin supervisor: el empleado no está en PRECIOS_KPIS "
              f"({', '.join(nombres[:5])}{'…' if len(nombres) > 5 else ''}). "
              f"Se agrupan bajo '(sin cruzar)' para que no se pierdan.")
    return err


def asignar_ids(nuevo: pd.DataFrame, previo: pd.DataFrame | None,
                prefijo: str, spec: pr.PeriodoSpec, llave_fn) -> pd.Series:
    """
    ID legible y ESTABLE entre corridas: PRE-202608-001.

    Funciona como un sistema de tickets. Al regenerar el archivo se leen los
    IDs ya asignados y solo los casos NUEVOS toman el siguiente número libre.
    Sin esto, un correlativo simple se correría en cada corrida: al aparecer
    un error nuevo cambiaría el orden y "el caso 12" pasaría a ser otro, que
    es justo lo que rompe cualquier conversación con el supervisor.

    Si un error se corrige y desaparece, su número queda retirado y no se
    reutiliza — así el histórico no termina apuntando a otra cosa.

    El ID es para hablar y hacer seguimiento; el cruce contra el archivo
    original se hace igual con las columnas llave, que siguen visibles.
    """
    base = f"{prefijo}-{spec.anio}{spec.mes:02d}-"
    ya: dict[str, str] = {}
    maximo = 0
    if previo is not None and not previo.empty and "ID_ERROR" in previo.columns:
        for k, i in zip(llave_fn(previo), previo["ID_ERROR"].astype(str)):
            if not i or i == "nan":
                continue
            ya[k] = i
            try:
                maximo = max(maximo, int(i.rsplit("-", 1)[-1]))
            except ValueError:
                pass  # un ID con otro formato no debe romper la numeración

    ids, nuevos = [], 0
    for k in llave_fn(nuevo):
        if k in ya:
            ids.append(ya[k])
        else:
            maximo += 1
            nuevos += 1
            ids.append(f"{base}{maximo:03d}")
    if ya:
        print(f"  🔖 IDs: {len(ids) - nuevos} conservados, {nuevos} nuevos")
    return pd.Series(ids, index=nuevo.index)


def _nombre_hoja(headers: dict, site_id: str, ruta: str) -> str:
    """Nombre de la primera hoja del libro, para no renombrarla al reescribir."""
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(ruta)}:/content")
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        return "Sheet1"
    return pd.ExcelFile(io.BytesIO(r.content)).sheet_names[0]


def escribir_id_en_origen(headers: dict, site_id: str, carpeta: str, archivo: str,
                          err: pd.DataFrame, llave_origen_fn, llave_err_fn) -> None:
    """
    Devuelve el ID_ERROR al archivo oficial, para poder ubicar ahí cada caso.

    Las filas sin error quedan con la celda vacía. Los IDs de OTROS periodos
    que ya estuvieran en el archivo se respetan: en exhibiciones el archivo
    acumula todos los meses y sería un problema borrarlos.

    OJO CON EL ORDEN
        El ETL de origen regenera su archivo desde cero en cada corrida, y con
        eso se lleva puesta esta columna. Por eso el workflow de errores corre
        DESPUÉS del ETL diario (día 1 y 16, una hora más tarde): si se corriera
        antes, el ID quedaría escrito y luego borrado el mismo día.
    """
    ruta = f"{carpeta}/{archivo}"
    df = leer_excel_cloud(headers, site_id, ruta, archivo, obligatorio=False)
    if df is None or df.empty:
        return

    # Se conserva el nombre de la hoja original: al reescribir el archivo con
    # un nombre distinto se rompería cualquier consulta o tabla dinamica que
    # apunte a esa hoja por nombre.
    hoja = _nombre_hoja(headers, site_id, ruta)

    nuevos = dict(zip(llave_err_fn(err), err["ID_ERROR"].astype(str)))
    previos = {}
    if "ID_ERROR" in df.columns:
        # IDs de otros periodos que ya estaban: no se tocan.
        for k, v in zip(llave_origen_fn(df), df["ID_ERROR"].fillna("").astype(str)):
            if v and v != "nan" and k not in nuevos:
                previos[k] = v

    llaves = llave_origen_fn(df)
    df["ID_ERROR"] = llaves.map({**previos, **nuevos}).fillna("")
    n = int((df["ID_ERROR"] != "").sum())

    subir_excel_cloud(headers, site_id, carpeta, archivo, {hoja: df})
    print(f"  🔗 ID_ERROR escrito en {archivo}: {n} fila(s) marcadas")


def conservar_correcciones(nuevo: pd.DataFrame, previo: pd.DataFrame | None) -> pd.DataFrame:
    """
    Arrastra PRECIO_CORREGIDO y OBSERVACION_SUPERVISOR de la corrida anterior.

    Sin esto, cada corrida borraría lo que el supervisor escribió. Es la parte
    que hace que el archivo funcione como formulario y no solo como reporte.
    """
    for c in COLS_RESPUESTA:
        nuevo[c] = ""
    if previo is None or previo.empty:
        return nuevo
    if any(c not in previo.columns for c in COLS_LLAVE_VISIBLES):
        print("  ⚠️  El archivo anterior no trae las columnas llave; no se pueden "
              "conservar las correcciones. Se genera limpio.")
        return nuevo

    # La llave se rearma desde las 4 columnas visibles, en los dos lados. Ya
    # no hay una columna ID concatenada: se separó para que la hoja se pueda
    # leer y filtrar, pero el cruce necesita las cuatro juntas igual.
    previo = previo.copy()
    llave_previo = _llave_visible(previo)
    llave_nuevo = _llave_visible(nuevo)
    for c in COLS_RESPUESTA:
        if c not in previo.columns:
            continue
        mapa = dict(zip(llave_previo, previo[c].fillna("")))
        nuevo[c] = llave_nuevo.map(mapa).fillna("")
    n = int((nuevo["PRECIO_CORREGIDO"].astype(str).str.strip() != "").sum())
    if n:
        print(f"  ✓ Se conservaron {n} corrección(es) que ya había escrito el supervisor.")
    return nuevo


# ─────────────────────────────────────────────────────────────────────────────
# PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

# La hoja que ve el supervisor: lo mínimo para entender el error y arreglarlo.
# Las 4 columnas llave van SEPARADAS (no concatenadas) para que se puedan leer
# y filtrar. Juntas identifican la captura de forma única, que es lo que
# permite reencontrar la fila entre corridas y conservar lo ya escrito.
COLS_SALIDA = [
    "ID_ERROR",
    "ID PDV", "CODIGO_SKU", "FECHA", "Empleado",
    "SUPERVISOR_LIDER", "PDV", "NOMBRE_PRODUCTO",
    "PRECIO_CAPTURADO", "DIAGNOSTICO",
    "PRECIO_CORREGIDO", "OBSERVACION_SUPERVISOR",
]

# La hoja con TODOS los registros: mismas columnas más la marca, y sin las
# dos de respuesta (ahí solo se corrige lo que está marcado como error).
COLS_TODOS = [
    "TIENE_ERROR", "SEVERIDAD", "SUPERVISOR_LIDER", "Empleado", "PDV",
    "ID del PDV", "Fecha de la encuesta", "Marca", "CODIGO_SKU",
    "NOMBRE_PRODUCTO", "PRESENCIA", "PRECIO_REGULAR", "MEDIANA_SKU",
    "MEDIA_SKU", "RATIO", "CAPTURAS_SKU", "DIAGNOSTICO",
]


def run(spec: pr.PeriodoSpec, umbral: float | None = None) -> int:
    global UMBRAL_MEDIO
    if umbral:
        UMBRAL_MEDIO = float(umbral)

    print("\n" + "=" * 60)
    print(f"  ERRORES DE PRECIOS — {spec.etiqueta}")
    print(f"  Umbral: fuera de [1/{UMBRAL_MEDIO:g}, {UMBRAL_MEDIO:g}] vs mediana del SKU")
    print("=" * 60)

    headers = {"Authorization": f"Bearer {obtener_token_azure()}"}
    site_id = obtener_site_id(headers)
    carpeta = paths.RUTA_CARPETA_SALIDAS_PRECIOS
    periodo_nom = f"{spec.mes_str_upper}_{spec.anio}"
    nombre_salida = f"ERRORES_PRECIOS_{periodo_nom}.xlsx"

    print("\nLeyendo insumos:")
    df = leer_excel_cloud(headers, site_id,
                          f"{carpeta}/ANALISIS_PRECIOS_{periodo_nom}.xlsx",
                          f"ANALISIS_PRECIOS_{periodo_nom}.xlsx")
    kpis = leer_excel_cloud(headers, site_id, f"{carpeta}/{paths.PR_OUT_KPIS.name}",
                            paths.PR_OUT_KPIS.name, obligatorio=False)
    previo = leer_excel_cloud(headers, site_id, f"{carpeta}/{nombre_salida}",
                              f"{nombre_salida} (corrida anterior)", obligatorio=False)

    print("\nDetectando:")
    err, sin_evaluar, todos = detectar(df)
    todos = asignar_supervisor(todos, kpis, silencioso=True)
    todos = todos.reindex(columns=COLS_TODOS).sort_values(
        ["TIENE_ERROR", "CODIGO_SKU", "RATIO"], ascending=[False, True, False])

    if err.empty:
        print("  ✓ No se encontró ningún precio fuera de rango.")
        subir_excel_cloud(headers, site_id, carpeta, nombre_salida,
                          {"Errores": pd.DataFrame(columns=COLS_SALIDA)})
        return 0

    err = asignar_supervisor(err, kpis)
    # Nombres pensados para quien corrige, no para quien programó el ETL.
    err = err.rename(columns={"PRECIO_REGULAR": "PRECIO_CAPTURADO",
                              "Fecha de la encuesta": "FECHA",
                              "ID del PDV": "ID PDV"})
    err["ID_ERROR"] = asignar_ids(err, previo, "PRE", spec, _llave_visible)
    err = conservar_correcciones(err, previo)
    # Se ordena ANTES de recortar columnas: SEVERIDAD sirve para dejar arriba
    # los casos más graves, pero no se muestra.
    # ALTA antes que MEDIA, y dentro de cada supervisor agrupado por gestor.
    # RATIO ya no está en esta hoja (se simplificó), así que se ordena por el
    # precio capturado, que deja juntos los casos parecidos.
    err = err.sort_values(["SEVERIDAD", "SUPERVISOR_LIDER", "Empleado", "PRECIO_CAPTURADO"],
                          ascending=[True, True, True, False])

    alta = int((err["SEVERIDAD"] == "ALTA").sum())
    print(f"\n  {len(err)} precio(s) a revisar — {alta} de severidad ALTA, "
          f"{len(err) - alta} MEDIA")
    print(f"  Repartidos en {err['SUPERVISOR_LIDER'].nunique()} supervisor(es)")

    resumen = (err.groupby("SUPERVISOR_LIDER", as_index=False)
                  .agg(ERRORES=("SEVERIDAD", "size"),
                       ALTA=("SEVERIDAD", lambda s: int((s == "ALTA").sum())),
                       GESTORES=("Empleado", "nunique"))
                  .sort_values("ERRORES", ascending=False))
    print()
    for _, r in resumen.head(10).iterrows():
        print(f"     {r['SUPERVISOR_LIDER'][:38]:38} {r['ERRORES']:>3} "
              f"({r['ALTA']} altas)")


    # El ID vuelve al archivo oficial para poder ubicar ahi cada caso.
    escribir_id_en_origen(
        headers, site_id, carpeta, f'ANALISIS_PRECIOS_{periodo_nom}.xlsx', err,
        llave_origen_fn=lambda d: (
            d['ID del PDV'].astype(str).str.strip().str.cat(
                [d['CODIGO_SKU'].astype(str).str.strip(),
                 d['Fecha de la encuesta'].astype(str).str.strip(),
                 d['Empleado'].astype(str).str.strip()], sep='|')),
        llave_err_fn=_llave_visible)

    # Recién acá se recortan las columnas: SEVERIDAD hizo falta para ordenar,
    # contar las graves y armar el resumen, pero no se muestra en la hoja.
    err = err.reindex(columns=COLS_SALIDA)

    print("\nGuardando:")
    # Una sola hoja: el archivo es un formulario, no un informe. Las otras
    # hojas (resumen, todos los registros, SKUs sin evaluar) hacían que el
    # supervisor tuviera que buscar dónde trabajar.
    subir_excel_cloud(headers, site_id, carpeta, nombre_salida, {"Errores": err})
    return len(err)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--mes", type=int, required=True, choices=range(1, 13))
    ap.add_argument("--anio", type=int, required=True)
    ap.add_argument("--umbral", type=float, default=None,
                    help=f"Ratio contra la mediana del SKU (default {UMBRAL_MEDIO:g})")
    args = ap.parse_args()
    # El envío del archivo a los supervisores vive en `envio_errores.py`, con
    # su propio workflow: este script solo detecta y escribe.
    run(pr.resolver(int(args.mes), int(args.anio)), args.umbral)
    return 0


if __name__ == "__main__":
    sys.exit(main())
