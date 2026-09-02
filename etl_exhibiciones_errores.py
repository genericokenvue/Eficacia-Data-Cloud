"""
etl_exhibiciones_errores.py
───────────────────────────
Detecta cantidades sospechosas en `Resultado exhibiciones gratis.xlsx` y arma
un archivo para que cada supervisor las revise y corrija.

CÓMO DETECTA
────────────
Regla del negocio: una exhibición con más de UMBRAL_CANTIDAD unidades en el
mismo PDV es sospechosa.

    Cantidad > 30   →  severidad ALTA
    Cantidad > 15   →  severidad MEDIA

A diferencia de precios, acá NO se compara contra una mediana: el criterio es
un tope fijo acordado con el negocio, no algo que salga de los datos.

QUÉ TAN COMÚN ES
────────────────
Medido sobre agosto 2026 (2.822 exhibiciones registradas):

    mediana de Cantidad        1
    percentil 90              10
    percentil 99              32
    máximo                   192

Con el tope en 15 salen 244 filas (8,7%). Es bastante más que en precios, así
que conviene mirar el reparto por supervisor antes de mandarlo.

DE DÓNDE SALE EL SUPERVISOR
───────────────────────────
El archivo de exhibiciones trae Empleado pero no su supervisor, y
KPI_Exhibiciones_Gratis tampoco. Se arma el mapa combinando los KPIs de los
otros módulos, que sí traen SUPERVISOR_LIDER. Se usan varios porque ninguno
solo cubre a toda la gente: CIF llega al 94%, No Presencia al 97%, y juntos
prácticamente al 100%.

EL ARCHIVO DE SALIDA ES TAMBIÉN EL FORMULARIO
─────────────────────────────────────────────
`ERRORES_EXHIBICIONES_<MES>_<AÑO>.xlsx` trae dos columnas vacías,
CANTIDAD_CORREGIDA y OBSERVACION_SUPERVISOR, que el supervisor llena en Excel
Online. En cada corrida se vuelve a leer ese mismo archivo y se conservan las
correcciones ya escritas: si se regenerara de cero, el supervisor perdería su
trabajo y no volvería a llenarlo nunca más.

USO
───
    python etl_exhibiciones_errores.py --mes 8 --anio 2026
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
# Tope fijo acordado con el negocio. No se pasa por línea de comandos en la
# operación normal: si cada corrida pudiera elegirlo, dos quincenas dejarían de
# ser comparables — no se sabría si bajaron los errores o si se movió el criterio.
UMBRAL_CANTIDAD = 15    # más de esto ya es sospechoso
UMBRAL_ALTO = 30        # más de esto es casi seguro un error

COLS_RESPUESTA = ["CANTIDAD_CORREGIDA", "OBSERVACION_SUPERVISOR"]

# Identifica una exhibición de forma estable entre corridas, para arrastrar lo
# que el supervisor ya escribió aunque el archivo se regenere.
COLS_LLAVE = ["ID PDV", "Tipo Exhibición", "Marca", "Empleado"]

ARCHIVO_ORIGEN = "Resultado exhibiciones gratis.xlsx"

# KPIs de donde se saca el vínculo empleado → supervisor, en orden de
# preferencia. Se combinan: el primero que tenga a la persona, gana.
KPIS_CON_SUPERVISOR = [
    (paths.RUTA_CARPETA_SALIDAS_NP, "NO_PRESENCIA_KPIS.xlsx"),
    (paths.RUTA_CARPETA_SALIDAS_CIF, "KPIS_CIF.xlsx"),
    (paths.RUTA_CARPETA_SALIDAS_SOS, "SOS_KPIS.xlsx"),
    (paths.RUTA_CARPETA_SALIDAS_PRECIOS, "PRECIOS_KPIS.xlsx"),
]

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
        print(f"  ℹ️  No existe todavía {descripcion} — se creará nuevo.")
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

    # 423 = alguien lo tiene abierto en el Excel de escritorio. Acá eso no es
    # raro: el archivo es el formulario que los supervisores editan.
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

def _llave(df: pd.DataFrame) -> pd.Series:
    partes = []
    for c in COLS_LLAVE:
        col = df[c] if c in df.columns else pd.Series([""] * len(df), index=df.index)
        partes.append(col.astype(str).str.strip())
    return partes[0].str.cat(partes[1:], sep="|")


def detectar(df: pd.DataFrame, spec: pr.PeriodoSpec) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (errores, todos_los_registros_del_periodo)."""
    d = df.copy()
    d["Cantidad"] = pd.to_numeric(d["Cantidad"], errors="coerce")

    # El archivo acumula todos los meses: hay que recortar al periodo.
    del_periodo = ((pd.to_numeric(d["Mes"], errors="coerce") == spec.mes)
                   & (pd.to_numeric(d["Año"], errors="coerce") == spec.anio))
    if not del_periodo.any():
        presentes = sorted(set(zip(pd.to_numeric(d["Mes"], errors="coerce"),
                                   pd.to_numeric(d["Año"], errors="coerce"))))
        raise ValueError(
            f"{ARCHIVO_ORIGEN} no tiene filas de {spec.mes:02d}/{spec.anio}. "
            f"Periodos presentes: {presentes}. Corré primero el ETL de exhibiciones."
        )
    d = d[del_periodo & d["Cantidad"].notna()].copy()
    print(f"  Exhibiciones del periodo {spec.mes:02d}/{spec.anio}: {len(d):,}")

    err = d[d["Cantidad"] > UMBRAL_CANTIDAD].copy()
    err["SEVERIDAD"] = "MEDIA"
    err.loc[err["Cantidad"] > UMBRAL_ALTO, "SEVERIDAD"] = "ALTA"
    err["DIAGNOSTICO"] = err["Cantidad"].apply(
        lambda c: f"{c:,.0f} unidades en un solo PDV (el tope es {UMBRAL_CANTIDAD})")

    todos = d.copy()
    todos["TIENE_ERROR"] = "NO"
    todos.loc[todos["Cantidad"] > UMBRAL_CANTIDAD, "TIENE_ERROR"] = "SI"
    todos["SEVERIDAD"] = ""
    todos.loc[todos["Cantidad"] > UMBRAL_CANTIDAD, "SEVERIDAD"] = "MEDIA"
    todos.loc[todos["Cantidad"] > UMBRAL_ALTO, "SEVERIDAD"] = "ALTA"
    return err, todos


def mapa_supervisores(headers: dict, site_id: str) -> dict:
    """
    Empleado → supervisor, combinando los KPIs que sí traen SUPERVISOR_LIDER.

    Se recorren en orden y el primero que tenga a la persona gana. Ninguno solo
    cubre a todos: juntos llegan a casi el 100%.
    """
    mapa: dict = {}
    for carpeta, archivo in KPIS_CON_SUPERVISOR:
        k = leer_excel_cloud(headers, site_id, f"{carpeta}/{archivo}",
                             archivo, obligatorio=False)
        if k is None or "SUPERVISOR_LIDER" not in k.columns or "NOMBRE" not in k.columns:
            continue
        sup = k["SUPERVISOR_LIDER"].astype(str).str.upper().str.strip()
        # Un supervisor vacío llega como "NAN" al pasar por astype(str), y sin
        # esto aparecería en el reporte como si fuera el nombre de alguien.
        sup = sup.where(~sup.isin(["", "NAN", "NONE"]), "")
        nom = k["NOMBRE"].astype(str).str.upper().str.strip()
        nuevos = 0
        for n, s in zip(nom, sup):
            if s and n and n not in mapa:
                mapa[n] = s
                nuevos += 1
        print(f"    {archivo}: +{nuevos} personas")
    return mapa


def asignar_supervisor(df: pd.DataFrame, mapa: dict, silencioso: bool = False) -> pd.DataFrame:
    df = df.copy()
    df["SUPERVISOR_LIDER"] = (
        df["Empleado"].astype(str).str.upper().str.strip()
          .map(mapa).replace("", pd.NA).fillna("(sin cruzar)"))
    n = int((df["SUPERVISOR_LIDER"] == "(sin cruzar)").sum())
    if n and not silencioso:
        nombres = sorted(df.loc[df["SUPERVISOR_LIDER"] == "(sin cruzar)", "Empleado"].unique())
        print(f"  ⚠️  {n} fila(s) sin supervisor: el empleado no aparece en ningún KPI "
              f"({', '.join(nombres[:4])}{'…' if len(nombres) > 4 else ''}). "
              f"Se agrupan bajo '(sin cruzar)' para que no se pierdan.")
    return df


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


def conservar_correcciones(nuevo: pd.DataFrame, previo: pd.DataFrame | None) -> pd.DataFrame:
    """Arrastra lo que el supervisor ya escribió en la corrida anterior."""
    for c in COLS_RESPUESTA:
        nuevo[c] = ""
    if previo is None or previo.empty:
        return nuevo
    if any(c not in previo.columns for c in COLS_LLAVE):
        print("  ⚠️  El archivo anterior no trae las columnas llave; no se pueden "
              "conservar las correcciones. Se genera limpio.")
        return nuevo
    # La llave se rearma desde las columnas visibles, en los dos lados.
    llave_previo, llave_nuevo = _llave(previo), _llave(nuevo)
    for c in COLS_RESPUESTA:
        if c not in previo.columns:
            continue
        mapa = dict(zip(llave_previo, previo[c].fillna("")))
        nuevo[c] = llave_nuevo.map(mapa).fillna("")
    n = int((nuevo["CANTIDAD_CORREGIDA"].astype(str).str.strip() != "").sum())
    if n:
        print(f"  ✓ Se conservaron {n} corrección(es) que ya había escrito el supervisor.")
    return nuevo


# ─────────────────────────────────────────────────────────────────────────────
# PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

# Las columnas llave van SEPARADAS (no concatenadas) para que se puedan leer
# y filtrar. Juntas identifican la exhibición de forma única, que es lo que
# permite reencontrar la fila entre corridas y conservar lo ya escrito.
COLS_SALIDA = [
    "ID_ERROR",
    "ID PDV", "Tipo Exhibición", "Marca", "Empleado",
    "SUPERVISOR_LIDER", "Nivel Impacto",
    "CANTIDAD", "DIAGNOSTICO",
    "CANTIDAD_CORREGIDA", "OBSERVACION_SUPERVISOR",
    # Va ÚLTIMA a propósito: es fea de leer, pero evita armar la fórmula
    # concatenada a mano para cruzar contra el archivo original.
    "LLAVE",
]

COLS_TODOS = [
    "TIENE_ERROR", "SEVERIDAD", "SUPERVISOR_LIDER", "Empleado", "ID PDV",
    "Categoría", "Tipo Exhibición", "Nivel Impacto", "Marca", "Cantidad",
    "Rol Empleado", "Mes", "Año",
]


def run(spec: pr.PeriodoSpec) -> int:
    print("\n" + "=" * 60)
    print(f"  ERRORES DE EXHIBICIONES — {spec.etiqueta}")
    print(f"  Regla: Cantidad > {UMBRAL_CANTIDAD} (ALTA si supera {UMBRAL_ALTO})")
    print("=" * 60)

    headers = {"Authorization": f"Bearer {obtener_token_azure()}"}
    site_id = obtener_site_id(headers)
    carpeta = paths.RUTA_CARPETA_SALIDAS_EXHIB
    nombre_salida = f"ERRORES_EXHIBICIONES_{spec.mes_str_upper}_{spec.anio}.xlsx"

    print("\nLeyendo insumos:")
    df = leer_excel_cloud(headers, site_id, f"{carpeta}/{ARCHIVO_ORIGEN}", ARCHIVO_ORIGEN)
    previo = leer_excel_cloud(headers, site_id, f"{carpeta}/{nombre_salida}",
                              f"{nombre_salida} (corrida anterior)", obligatorio=False)

    print("\nArmando el mapa empleado → supervisor:")
    mapa = mapa_supervisores(headers, site_id)
    print(f"    total: {len(mapa)} personas con supervisor conocido")

    print("\nDetectando:")
    err, todos = detectar(df, spec)
    todos = asignar_supervisor(todos, mapa, silencioso=True)
    todos = todos.reindex(columns=COLS_TODOS).sort_values(
        ["TIENE_ERROR", "Cantidad"], ascending=[False, False])

    if err.empty:
        print(f"  ✓ Ninguna exhibición supera las {UMBRAL_CANTIDAD} unidades.")
        subir_excel_cloud(headers, site_id, carpeta, nombre_salida,
                          {"Errores": pd.DataFrame(columns=COLS_SALIDA)})
        return 0

    err = asignar_supervisor(err, mapa)
    err = err.rename(columns={"Cantidad": "CANTIDAD"})
    err["ID_ERROR"] = asignar_ids(err, previo, "EXH", spec, _llave)
    err["LLAVE"] = _llave(err)
    err = conservar_correcciones(err, previo)
    # Se ordena ANTES de recortar columnas: SEVERIDAD deja arriba los casos
    # más graves, pero no se muestra.
    err = err.sort_values(["SEVERIDAD", "SUPERVISOR_LIDER", "Empleado", "CANTIDAD"],
                          ascending=[True, True, True, False])

    alta = int((err["SEVERIDAD"] == "ALTA").sum())
    print(f"\n  {len(err)} exhibición(es) a revisar — {alta} ALTA, {len(err) - alta} MEDIA")
    print(f"  Repartidas en {err['SUPERVISOR_LIDER'].nunique()} supervisor(es)")

    resumen = (err.groupby("SUPERVISOR_LIDER", as_index=False)
                  .agg(ERRORES=("SEVERIDAD", "size"),
                       ALTA=("SEVERIDAD", lambda s: int((s == "ALTA").sum())),
                       GESTORES=("Empleado", "nunique"),
                       CANTIDAD_MAX=("CANTIDAD", "max"))
                  .sort_values("ERRORES", ascending=False))
    print()
    for _, r in resumen.head(10).iterrows():
        print(f"     {str(r['SUPERVISOR_LIDER'])[:38]:38} {r['ERRORES']:>3} "
              f"({r['ALTA']} altas, máx {r['CANTIDAD_MAX']:.0f})")

    # Recién acá se recortan las columnas: SEVERIDAD hizo falta para ordenar,
    # contar las graves y armar el resumen, pero no se muestra en la hoja.
    err = err.reindex(columns=COLS_SALIDA)

    print("\nGuardando:")
    # Una sola hoja: el archivo es un formulario, no un informe.
    subir_excel_cloud(headers, site_id, carpeta, nombre_salida, {"Errores": err})
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
