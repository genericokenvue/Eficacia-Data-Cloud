# -*- coding: utf-8 -*-
"""
diagnostico_cruce.py
────────────────────
Verificación PREVIA (solo lectura) antes de correr `calcular_cumplimientos.py`.

Responde tres preguntas, sin enviar un solo correo ni escribir un solo archivo:

  1. ¿Están todos los insumos del periodo en SharePoint, con el nombre correcto?
  2. ¿Cuántas personas cruzan contra Base cupos y cuántas quedan fuera?
  3. ¿De qué periodo son REALMENTE los datos que hay hoy en la nube?

Por qué existe
──────────────
`calcular_cumplimientos.py` está diseñado para no caerse: si falta un insumo
opcional, la hoja correspondiente sale con "No aplica" y el proceso sigue. Eso
es correcto en producción (no queremos que un archivo faltante tumbe el envío
de 37 correos), pero significa que un insumo mal nombrado o de otro mes NO se
manifiesta como error — se manifiesta como una hoja vacía o, peor, como datos
del mes equivocado.

El caso más peligroso es el Rutero de D&P: si falta el del periodo, el código
cae a `Rutero_Droguerias.xlsx` (un genérico que sí existe en la carpeta) y
calcula las hojas de productos nuevos contra el mes equivocado. El reporte sale
completo, con números, y solo queda una advertencia en el log.

Este script corre ANTES y te lo dice de frente.

Uso
───
    # Verificar el periodo que hay hoy en la nube (autodetectado)
    python diagnostico_cruce.py

    # Verificar un periodo específico (antes de subir los insumos, p.ej.)
    python diagnostico_cruce.py --mes 8 --anio 2026

    # Solo revisar archivos, sin correr el cruce (mucho más rápido)
    python diagnostico_cruce.py --mes 8 --anio 2026 --solo-insumos

    # Ver además el contenido completo de las carpetas del periodo
    python diagnostico_cruce.py --listar-carpetas

Código de salida
────────────────
    0 → todo en orden
    1 → falta al menos un insumo OBLIGATORIO, o el periodo pedido no coincide
        con el que traen los KPIs
    2 → solo advertencias (opcionales ausentes o personas sin cruzar)

SEGURIDAD
─────────
Antes de importar cualquier módulo del proyecto, este script reemplaza
`requests.put/post/patch/delete` por funciones que no hacen nada y solo
registran el intento. Como todos los módulos comparten el mismo objeto
`requests` en memoria, ninguno puede escribir en SharePoint aunque lo intente.
Los GET quedan intactos, así que las lecturas funcionan con normalidad.

Al final se reporta cuántos intentos de escritura se bloquearon.
"""
from __future__ import annotations

import argparse
import io
import os
import sys
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# Consola en UTF-8 — el proyecto imprime emojis y acentos, y la consola de
# Windows por defecto (cp1252) los convierte en excepción o en basura.
# ─────────────────────────────────────────────────────────────────────────────
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

_ALERTAS_DIR = Path(__file__).resolve().parent
_SCRIPTS_DIR = _ALERTAS_DIR.parent
for _p in (str(_ALERTAS_DIR), str(_SCRIPTS_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# `paths.py` y los ETLs llaman `load_dotenv()` sin argumentos, que busca el .env
# desde el directorio de trabajo hacia arriba. Nos paramos en SCRIPTS/ para que
# lo encuentre sin importar desde dónde se invoque este script.
os.chdir(_SCRIPTS_DIR)

from dotenv import load_dotenv
load_dotenv(_SCRIPTS_DIR / ".env")


# ═════════════════════════════════════════════════════════════════════════════
# CANDADO DE ESCRITURA — debe ejecutarse ANTES de importar los módulos
# ═════════════════════════════════════════════════════════════════════════════
import requests

_ESCRITURAS_BLOQUEADAS: list[tuple[str, str]] = []


class _RespuestaBloqueada:
    """
    Imita una respuesta HTTP exitosa. Se devuelve 200 a propósito: así el código
    llamador sigue su curso normal en vez de entrar a su rama de error, y el
    diagnóstico refleja el flujo real de una corrida de producción.

    Efecto secundario a tener en cuenta al leer la salida: los módulos que
    imprimen "✓ guardado" tras un PUT lo van a imprimir igual. Ese mensaje es
    falso — la lista de bloqueos al final del reporte es la fuente de verdad.
    """
    status_code = 200
    text = "(escritura bloqueada por diagnostico_cruce.py — modo solo lectura)"

    def json(self):
        return {}


def _bloquear(verbo: str):
    def _fn(url, *args, **kwargs):
        _ESCRITURAS_BLOQUEADAS.append((verbo, str(url)))
        return _RespuestaBloqueada()
    return _fn


requests.put    = _bloquear("PUT")
requests.post   = _bloquear("POST")
requests.patch  = _bloquear("PATCH")
requests.delete = _bloquear("DELETE")

# A partir de aquí ya es imposible escribir en SharePoint.
import urllib.parse

import pandas as pd

import paths
import base_cupos as bcm
import calcular_cumplimientos as cc


# ═════════════════════════════════════════════════════════════════════════════
# Utilidades
# ═════════════════════════════════════════════════════════════════════════════
ANCHO = 78


def titulo(texto: str, char: str = "═") -> None:
    print("\n" + char * ANCHO)
    print(f"  {texto}")
    print(char * ANCHO)


def seccion(texto: str) -> None:
    print("\n" + "─" * ANCHO)
    print(f"  {texto}")
    print("─" * ANCHO)


_DRIVE_ID: dict = {}


def _drive_id() -> str:
    if "id" not in _DRIVE_ID:
        _DRIVE_ID["id"] = cc._obtener_default_drive_id(cc._obtener_token_azure())
    return _DRIVE_ID["id"]


def existe_en_nube(ruta: str) -> bool:
    """¿Existe este archivo/carpeta en SharePoint? (solo un GET de metadatos)"""
    try:
        headers = {"Authorization": f"Bearer {cc._obtener_token_azure()}"}
        ruta_cod = urllib.parse.quote(cc._limpiar_ruta_graph(ruta))
        url = f"https://graph.microsoft.com/v1.0/drives/{_drive_id()}/root:/{ruta_cod}"
        return requests.get(url, headers=headers).status_code == 200
    except Exception:
        return False


# ═════════════════════════════════════════════════════════════════════════════
# Los 6 insumos cuyo NOMBRE depende del periodo
# ═════════════════════════════════════════════════════════════════════════════
def insumos_variables(mes: int, anio: int) -> list[tuple[str, str]]:
    """
    Estos son los que hay que subir cada mes con el nombre correcto. El resto
    de insumos tiene nombre fijo y los ETLs simplemente los sobrescriben.

    OJO con las mayúsculas: cuatro usan el mes en MAYÚSCULAS y uno solo con la
    inicial ("Base Exhibiciones Planning Agosto 2026.xlsx"). El código localiza
    los de Exhibiciones con `re.match`, que distingue mayúsculas — un archivo
    bien ubicado pero mal capitalizado no se encuentra.
    """
    return [
        ("Informe de visitas (CIF)",   paths.cloud_cif_visitas(mes, anio)),
        ("Exh — Planning maestro",     paths.cloud_exhib_planning(mes, anio)),
        ("Exh — Base Planning",        paths.cloud_exhib_base_planning(mes, anio)),
        ("D&P — Consolidado Impactos", paths.cloud_dyp_impactos(mes, anio)),
        ("D&P — Consolidado Ventas",   paths.cloud_dyp_ventas(mes, anio)),
        ("D&P — Rutero",               paths.cloud_dyp_rutero(mes, anio)),
    ]


def insumos_fijos(mes: int, anio: int) -> list[tuple[str, str]]:
    """
    Nombre fijo: los sobrescribe el ETL de cada módulo en cada corrida.

    Los cuatro archivos de DETALLE se piden a `calcular_cumplimientos`, que es
    el que define cómo se llaman. Antes se repetían aquí con un nombre sin
    periodo que ningún ETL escribe, así que el diagnóstico daba por buenos unos
    archivos viejos y no alcanzaba a ver que las alertas iban con datos rezagados.
    """
    return [
        ("CIF — detalle",         cc.ruta_cif_detalle()),
        ("SOS — detalle",         cc.ruta_sos_detalle(mes, anio)),
        ("NP — detalle",          cc.ruta_np_detalle(mes, anio)),
        ("Precios — detalle",     cc.ruta_pr_detalle(mes, anio)),
        ("KPIs CIF",              f"{paths.RUTA_CARPETA_SALIDAS_CIF}/{paths.CIF_OUT_KPIS.name}"),
        ("KPIs SOS",              f"{paths.RUTA_CARPETA_SALIDAS_SOS}/{paths.SOS_OUT_KPIS.name}"),
        ("KPIs NP",               f"{paths.RUTA_CARPETA_SALIDAS_NP}/{paths.NP_OUT_KPIS.name}"),
        ("KPIs Precios",          f"{paths.RUTA_CARPETA_SALIDAS_PRECIOS}/{paths.PR_OUT_KPIS.name}"),
        ("KPIs Exhib. Pagadas",   f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/{paths.EXHIB_PAG_OUT_KPIS.name}"),
        ("KPIs Exhib. Gratis",    f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/{paths.EXHIB_GRA_OUT_KPIS.name}"),
        ("Exh — Fuera de regla",  f"{paths.RUTA_CARPETA_SALIDAS_EXHIB}/Exh_Gratis_Fuera_de_Regla.xlsx"),
        ("Base cupos",            paths.cloud_dyp_base_cupos(anio)),
        ("D&P — Listas",          paths.cloud_dyp_listas(anio)),
    ]


# ═════════════════════════════════════════════════════════════════════════════
# Bloques del reporte
# ═════════════════════════════════════════════════════════════════════════════
def bloque_base_cupos() -> tuple[pd.DataFrame, dict, set]:
    seccion("PASO 1 — Base cupos (tabla maestra de personas)")

    df_bc    = bcm.cargar()
    universo = bcm.universo_personas(df_bc)
    sups_bc  = bcm.supervisores(df_bc)

    nombres_sups = set(sups_bc["NOMBRE"].astype(str).str.upper())
    for col in ("ES_GDD", "ES_LIDER"):
        if col in df_bc.columns:
            nombres_sups.update(df_bc[df_bc[col] == True]["NOMBRE"].astype(str).str.upper())

    idx_bc = bcm.construir_indices(df_bc)

    n_gdd   = int((df_bc.get("ES_GDD")   == True).sum()) if "ES_GDD"   in df_bc.columns else 0
    n_lider = int((df_bc.get("ES_LIDER") == True).sum()) if "ES_LIDER" in df_bc.columns else 0
    descartadas = len(df_bc) - len(universo)

    print(f"\n  Filas crudas en Base_cupos ....... {len(df_bc):>5}")
    print(f"  Universo activo (con ACRONIMO) ... {len(universo):>5}")
    if descartadas > 0:
        print(f"  Descartadas (sin ACRONIMO) ...... {descartadas:>5}   ← no entran a los reportes")
    print(f"  ─────────────────────────────────────")
    print(f"  Supervisores ..................... {len(sups_bc):>5}")
    print(f"  GDD .............................. {n_gdd:>5}")
    print(f"  Líderes de ejecución ............. {n_lider:>5}")
    print(f"  Destinatarios (adjuntos a generar) {len(nombres_sups):>5}")

    return universo, idx_bc, nombres_sups


def bloque_cruce(idx_bc: dict, nombres_sups: set) -> tuple[dict, int, int, dict]:
    seccion("PASO 2 — Cruce de los KPIs contra Base cupos")

    kpis = cc.cargar_kpis_v3(idx_bc, nombres_sups)

    modulos = [
        ("CIF",            "cif_gest"),
        ("SOS",            "sos_gest"),
        ("No Presencia",   "np_gest"),
        ("Precios",        "pr_gest"),
        ("Exhib. Pagadas", "exp_gest"),
        ("Exhib. Gratis",  "egr_gest"),
    ]

    titulo("RESULTADO DEL CRUCE POR MÓDULO")
    print(f"  {'MÓDULO':<16} {'FILAS':>7} {'CRUZAN':>8} {'FUERA':>7} {'% ÉXITO':>9}  {'CÉDULA?':>8}")
    print("  " + "-" * (ANCHO - 4))

    total_filas = total_ok = 0
    for etiqueta, clave in modulos:
        df = kpis.get(clave, pd.DataFrame())
        if df is None or df.empty:
            print(f"  {etiqueta:<16} {0:>7} {'—':>8} {'—':>7} {'VACÍO':>9}  {'—':>8}")
            continue
        n  = len(df)
        ok = int((df["ACRONIMO"].astype(str).str.strip() != "").sum()) \
             if "ACRONIMO" in df.columns else 0
        pct = (ok / n * 100) if n else 0.0
        # Los módulos sin CEDULA dependen solo del nombre y cruzan peor —
        # por eso se muestra la columna.
        print(f"  {etiqueta:<16} {n:>7} {ok:>8} {n-ok:>7} {pct:>8.1f}%  "
              f"{('sí' if 'CEDULA' in df.columns else 'NO'):>8}")
        total_filas += n
        total_ok    += ok

    print("  " + "-" * (ANCHO - 4))
    pct_tot = (total_ok / total_filas * 100) if total_filas else 0.0
    print(f"  {'TOTAL':<16} {total_filas:>7} {total_ok:>8} {total_filas-total_ok:>7} {pct_tot:>8.1f}%")

    # Personas sin cruzar, deduplicadas por nombre (la misma persona aparece
    # como faltante en varios módulos a la vez).
    fuera: dict[str, str] = {}
    for c in kpis.get("candidatos_faltantes", []):
        nom = str(c.get("NOMBRE", "")).strip().upper()
        ced = str(c.get("CEDULA", "")).strip()
        if not nom:
            continue
        if nom not in fuera or (not fuera[nom] and ced):
            fuera[nom] = ced

    titulo(f"PERSONAS SIN CRUZAR: {len(fuera)} distinta(s)")
    if fuera:
        print(f"  {'NOMBRE':<45} {'CÉDULA':>14}")
        print("  " + "-" * 62)
        for nom, ced in sorted(fuera.items()):
            print(f"  {nom[:44]:<45} {(ced or '(sin cédula)'):>14}")
        print("\n  💡 Revisa arriba el detalle por persona: el log distingue")
        print("     'SÍ ESTÁ en Base cupos pero sin ACRONIMO' (se arregla llenando")
        print("     la celda) de 'NO está en Base cupos' (hay que crear la fila).")
    else:
        print("  ✅ Ninguna — todas las personas de los KPIs cruzaron.")

    return kpis, total_filas, total_ok, fuera


def bloque_insumos(mes: int, anio: int) -> tuple[list, list]:
    titulo(f"PASO 3 — Insumos del periodo {mes:02d}/{anio}")

    print("\n  ARCHIVOS QUE CAMBIAN DE NOMBRE CADA MES")
    print("  " + "-" * (ANCHO - 4))
    faltan_var = []
    for desc, ruta in insumos_variables(mes, anio):
        ok = existe_en_nube(ruta)
        print(f"  {'✓' if ok else '✗'}  {desc}")
        if not ok:
            faltan_var.append((desc, ruta))
            print(f"       falta: {ruta}")

    print("\n  ARCHIVOS DE NOMBRE FIJO (los sobrescribe el ETL)")
    print("  " + "-" * (ANCHO - 4))
    faltan_fijos = []
    for desc, ruta in insumos_fijos(mes, anio):
        ok = existe_en_nube(ruta)
        print(f"  {'✓' if ok else '✗'}  {desc}")
        if not ok:
            faltan_fijos.append((desc, ruta))
            print(f"       falta: {ruta}")

    if faltan_var:
        print("\n  ⚠️  CONSECUENCIA DE LO QUE FALTA:")
        consecuencias = {
            "Informe de visitas (CIF)":   "el correo sale SIN el rango de fechas del corte",
            "Exh — Planning maestro":     "hoja 'Exh Pagadas — No Capturadas' → 'No disponible'",
            "Exh — Base Planning":        "hoja 'Exh Pagadas — No Capturadas' → 'No disponible'",
            "D&P — Consolidado Impactos": "hoja 'Impactos — Detalle' → 'No aplica'",
            "D&P — Consolidado Ventas":   "hojas de productos nuevos → 'No aplica'",
            "D&P — Rutero":               "🔴 CAE AL GENÉRICO 'Rutero_Droguerias.xlsx' y "
                                          "calcula productos nuevos CON DATOS DE OTRO MES",
        }
        for desc, _ in faltan_var:
            print(f"     · {desc}: {consecuencias.get(desc, 'hoja incompleta')}")

    return faltan_var, faltan_fijos


def bloque_carpetas(anio: int) -> None:
    titulo("CONTENIDO REAL DE LAS CARPETAS DEL PERIODO")
    carpetas = [
        ("BASES/CIF/{}".format(anio),          paths.bases_cif(anio)),
        ("BASES/EXHIBICIONES/{}".format(anio), paths.bases_exhib(anio)),
        ("BASES/DYP/{}/Rutero".format(anio),   f"{paths.bases_dyp(anio)}/Rutero"),
        ("BASES/DYP/{}/Listas".format(anio),   f"{paths.bases_dyp(anio)}/Listas"),
        ("SALIDAS/DYP",                        paths.RUTA_CARPETA_SALIDAS_DYP),
    ]
    for etiqueta, carpeta in carpetas:
        print(f"\n  📁 {etiqueta}")
        try:
            hijos = cc._listar_hijos_cloud(carpeta)
            if not hijos:
                print("       (vacía)")
            for h in sorted(hijos, key=lambda x: str(x.get("name", ""))):
                print(f"       {'📂' if h.get('folder') else '  '} {h.get('name')}")
        except Exception as e:
            print(f"       ⚠️  No se pudo listar: {str(e)[:80]}")


# ═════════════════════════════════════════════════════════════════════════════
# Punto de entrada
# ═════════════════════════════════════════════════════════════════════════════
def main() -> int:
    ap = argparse.ArgumentParser(
        description="Verificación previa (solo lectura) de calcular_cumplimientos.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("Uso\n───")[-1].split("SEGURIDAD")[0],
    )
    ap.add_argument("--mes",  type=int, choices=range(1, 13), metavar="1-12",
                    help="Mes a verificar. Si se omite, se detecta de KPIS_CIF.xlsx.")
    ap.add_argument("--anio", type=int, metavar="AAAA",
                    help="Año a verificar. Si se omite, se detecta de KPIS_CIF.xlsx.")
    ap.add_argument("--solo-insumos", action="store_true",
                    help="Omitir el cruce de KPIs (mucho más rápido). Requiere --mes y --anio.")
    ap.add_argument("--listar-carpetas", action="store_true",
                    help="Mostrar el contenido completo de las carpetas del periodo.")
    args = ap.parse_args()

    titulo("DIAGNÓSTICO PREVIO — MODO SOLO LECTURA")
    print("  PUT/POST/PATCH/DELETE bloqueados: no se escribe nada en SharePoint.")
    print(f"\n  Sitio   : {paths.SHAREPOINT_SITE_NAME}")
    print(f"  BASES   : {paths._BASES_ROOT}")
    print(f"  SALIDAS : {paths._SALIDAS_ROOT}")

    universo = pd.DataFrame()
    fuera: dict = {}
    total_filas = total_ok = 0
    mes_kpis = anio_kpis = None

    if args.solo_insumos:
        if not (args.mes and args.anio):
            print("\n  ✗ --solo-insumos requiere --mes y --anio.")
            return 1
        mes, anio = args.mes, args.anio
    else:
        universo, idx_bc, nombres_sups = bloque_base_cupos()
        kpis, total_filas, total_ok, fuera = bloque_cruce(idx_bc, nombres_sups)

        # Periodo que traen REALMENTE los KPIs (mismo criterio que main()).
        df_cif = kpis.get("cif_gest", pd.DataFrame())
        if not df_cif.empty and "MES" in df_cif.columns:
            mes_kpis  = int(pd.to_numeric(df_cif["MES"],  errors="coerce").dropna().mode().iloc[0])
            anio_kpis = int(pd.to_numeric(df_cif["AÑO"], errors="coerce").dropna().mode().iloc[0])

        mes  = args.mes  or mes_kpis  or 0
        anio = args.anio or anio_kpis or 0
        if not (mes and anio):
            print("\n  ✗ No pude detectar el periodo y no me diste --mes/--anio.")
            return 1

    cc._ANIO_ACTIVO = int(anio)

    # Aviso clave: pediste un periodo pero los KPIs traen otro. Es exactamente
    # el error de "subí los insumos nuevos pero no volví a correr los ETLs".
    desalineado = (
        mes_kpis is not None
        and (args.mes or args.anio)
        and (mes, anio) != (mes_kpis, anio_kpis)
    )
    if desalineado:
        titulo("⚠️  PERIODO DESALINEADO")
        print(f"  Pediste verificar ....... {mes:02d}/{anio}")
        print(f"  Pero los KPIs traen ..... {mes_kpis:02d}/{anio_kpis}")
        print("\n  calcular_cumplimientos.py detecta el periodo de los KPIs, así que")
        print(f"  procesaría {mes_kpis:02d}/{anio_kpis}, no {mes:02d}/{anio}.")
        print("  Corre primero los ETLs de cada módulo para regenerar los *_KPIS.xlsx.")

    faltan_var, faltan_fijos = bloque_insumos(mes, anio)

    if args.listar_carpetas:
        bloque_carpetas(anio)

    # ── Cierre ───────────────────────────────────────────────────────────────
    titulo("RESUMEN")
    if not args.solo_insumos:
        pct = (total_ok / total_filas * 100) if total_filas else 0.0
        print(f"  Universo de personas activas .......... {len(universo)}")
        print(f"  Filas de KPI procesadas ............... {total_filas}")
        print(f"  Filas que cruzaron .................... {total_ok} ({pct:.1f}%)")
        print(f"  Personas distintas sin cruzar ......... {len(fuera)}")
    print(f"  Insumos de nombre variable ausentes ... {len(faltan_var)} de 6")
    print(f"  Insumos de nombre fijo ausentes ....... {len(faltan_fijos)} de 13")

    print(f"\n  🔒 Escrituras bloqueadas: {len(_ESCRITURAS_BLOQUEADAS)}")
    for verbo, url in _ESCRITURAS_BLOQUEADAS:
        # Mostrar solo la parte legible de la ruta, no el drive-id.
        legible = urllib.parse.unquote(url.split("/root:/")[-1].split(":/")[0]) \
                  if "/root:/" in url else url
        print(f"     · {verbo} {legible[:88]}")
    print("  (Ningún archivo de SharePoint fue modificado.)")

    titulo("VEREDICTO")
    if faltan_fijos or desalineado:
        print("  🔴 NO CORRAS calcular_cumplimientos.py todavía.")
        if faltan_fijos:
            print("     Faltan salidas de ETL — corre primero los ETLs de esos módulos.")
        if desalineado:
            print("     Los KPIs son de otro periodo.")
        codigo = 1
    elif faltan_var:
        print("  🟡 PUEDES correr, pero saldrán hojas incompletas.")
        print("     Revisa arriba la consecuencia de cada archivo faltante.")
        print("     Si falta el Rutero, los productos nuevos saldrán con datos de otro mes.")
        codigo = 2
    elif fuera:
        print("  🟡 Todos los insumos están, pero hay personas sin cruzar.")
        print("     Van a quedar fuera de los reportes. Completa sus ACRONIMO en Base_cupos.xlsx.")
        codigo = 2
    else:
        print("  🟢 TODO EN ORDEN — puedes correr calcular_cumplimientos.py.")
        codigo = 0
    print("═" * ANCHO)
    return codigo


if __name__ == "__main__":
    sys.exit(main())
