"""
aplicar_correcciones.py
─────────────────────────
Lee un archivo de respuesta — el Excel que un supervisor devolvió con las
columnas *_CORREGIDO llenas — y aplica cada corrección al archivo OFICIAL
correspondiente: el mismo que ya tiene el ID_ERROR escrito.

Cada fila corregida queda además marcada con CORREGIDO = "Sí" en esa misma
encuesta cruda (la columna se agrega sola, en "No", la primera vez que se
corrige algo en ese archivo) — para ver de un vistazo qué ya se revisó sin
tener que abrir LOG_CORRECCIONES.xlsx.

CÓMO CRUZA CADA MÓDULO
───────────────────────
Precios y Exhibiciones: por ID_ERROR, 1 a 1 — el ID ya identifica una única
fila en el archivo oficial (se escribió ahí en `escribir_id_en_origen` de
cada ETL de errores).

Espacios: por la llave natural (Empleado, PDV, Categoría, Marca, Mes, Año)
en vez de ID_ERROR, porque el motivo MOTIVO_CAMBIO no
quedó escrito en la encuesta origen (ver etl_espacios_errores.py — cruzarlo
por ID_ERROR ahí pisaría el de otras marcas de la misma categoría). Se
confirmó con datos reales que esa combinación es siempre UNA sola fila
cruda (agosto 2026: 40.505 de 40.505, sin excepción) — no hay ambigüedad de
qué fila tocar:
  · CM_MARCA_CORREGIDO → sobrescribe solo esa fila (esa marca puntual)
  · UNIVERSO_CORREGIDO → sobrescribe TODAS las filas del mismo
    (Empleado, PDV, Categoría, Mes, Año) — el universo es un valor por
    categoría, se repite en cada fila de marca de ese caso. Nunca sale de
    ahí: no toca otros PDV ni otros gestores.

DE DÓNDE SALE EL PERIODO
─────────────────────────
No hace falta pasarlo por argumento: el ID_ERROR ya lo trae en el formato
<PREFIJO>-<AAAAMM>-<NNN> (ej. "ESP-202608-010" → agosto 2026). Cada fila se
procesa contra el archivo oficial de SU propio periodo, aunque la respuesta
mezcle casos de meses distintos.

NO DISPARA RECÁLCULO
─────────────────────
Corrige el archivo fuente y ahí termina — no vuelve a correr
etl_precios.py / etl_exhibiciones_gratis.py / etl_sos.py. La próxima
corrida programada de esos ETL ya toma el valor corregido.

QUEDA AUDITADO
───────────────
Cada corrección aplicada se agrega (no se sobrescribe, se acumula) a
SALIDAS/LOG_CORRECCIONES.xlsx — cuándo, de qué archivo de respuesta salió,
módulo, ID_ERROR, campo, valor anterior → valor nuevo.

CÓMO SE ENGANCHA CON POWER AUTOMATE
─────────────────────────────────────
Power Automate solo hace UNA cosa: cuando llega un correo con asunto
"CAPTURA DE ERRORES", guarda el adjunto en
SALIDAS/RESPUESTAS_ERRORES/ (`paths.RUTA_CARPETA_RESPUESTAS_ERRORES`). No
necesita conector premium, ni credenciales de GitHub, ni saber nada de
Excel — solo "cuando llega un correo" + "crear archivo".

Este script, corriendo por cron (`--carpeta`), revisa esa carpeta, procesa
cada archivo nuevo y lo marca leyendo el prefijo "PROCESADO_" en el nombre
— así una corrida no vuelve a aplicar lo que ya aplicó otra. Si un archivo
falla, se avisa y se sigue con los demás (no corta el lote completo).

USO
───
    python aplicar_correcciones.py --archivo "SALIDAS/RESPUESTAS_ERRORES/Errores_Yuly_Alarcon_08_2026.xlsx"
    python aplicar_correcciones.py --carpeta "SALIDAS/RESPUESTAS_ERRORES"

(las rutas son relativas a la carpeta de documentos del sitio SharePoint,
misma convención que el resto de los ETL)
"""
from __future__ import annotations

import argparse
import io
import sys
import urllib.parse
from datetime import datetime

import pandas as pd
import requests

import paths
import periodo_resolver as pr
import etl_espacios_errores as esp

# ─────────────────────────────────────────────────────────────────────────────
# QUÉ MÓDULOS SE PUEDEN CORREGIR Y CÓMO
# ─────────────────────────────────────────────────────────────────────────────

def _archivo_precios(mes: int, anio: int) -> str:
    """
    Respuestas_Encuesta_<MES>_<AÑO>.xlsx — la encuesta CRUDA, no
    ANALISIS_PRECIOS. ANALISIS_PRECIOS se regenera entero cada corrida de
    etl_precios.py y perdería la corrección; la encuesta cruda es la que
    ese ETL usa como INSUMO, así que el valor corregido sí sobrevive a la
    próxima corrida. Ver conversación del 2026-09-14.
    """
    spec = pr.resolver(mes, anio)
    return f"{paths.bases_precios(anio)}/Respuestas_Encuesta_{spec.mes_str_upper}_{anio}.xlsx"


CFG_PRECIOS = {
    "archivo": _archivo_precios,
    "col_corregido": "PRECIO_CORREGIDO",
    "col_oficial": "Digite Precio Regular",
}

CFG_EXHIBICIONES = {
    "carpeta": paths.RUTA_CARPETA_SALIDAS_EXHIB,
    "archivo": lambda mes, anio: "Resultado exhibiciones gratis.xlsx",
    "col_corregido": "CANTIDAD_CORREGIDA",
    "col_oficial": "Cantidad",
}


# ─────────────────────────────────────────────────────────────────────────────
# SHAREPOINT (funciones propias sin credenciales — se reutiliza la sesión
# ya autenticada de etl_espacios_errores para todo lo demás)
# ─────────────────────────────────────────────────────────────────────────────

def _marcar_procesado(headers: dict, site_id: str, item_id: str, nombre_actual: str) -> None:
    """Renombra el archivo con el prefijo PROCESADO_ — así una corrida futura lo salta."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
    r = requests.patch(url, headers={**headers, "Content-Type": "application/json"},
                       json={"name": f"PROCESADO_{nombre_actual}"})
    if r.status_code not in (200, 201):
        print(f"  ⚠️  No se pudo marcar {nombre_actual} como procesado: {r.status_code} - {r.text}")


def _descargar_bytes(headers: dict, site_id: str, ruta: str) -> bytes:
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(ruta)}:/content")
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        raise FileNotFoundError(f"No se pudo leer {ruta}: {r.status_code}")
    return r.content


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────────────────

def _periodo_de_id(id_error: str) -> tuple[int, int] | None:
    """'ESP-202608-010' → (8, 2026). None si el formato no matchea."""
    partes = str(id_error).split("-")
    if len(partes) < 2 or len(partes[1]) != 6 or not partes[1].isdigit():
        return None
    yyyymm = partes[1]
    return int(yyyymm[4:6]), int(yyyymm[:4])


def _no_vacio(serie: pd.Series) -> pd.Series:
    return serie.astype(str).str.strip().replace({"nan": "", "None": ""}) != ""


def _valor_no_vacio(v) -> bool:
    """
    Mismo criterio que `_no_vacio` pero para un valor suelto (una celda), no
    una Series — hace falta acá porque `str(float('nan'))` da el texto
    "nan", que NO está vacío para un chequeo ingenuo de `.strip()`. Sin este
    helper, una celda vacía (NaN) se interpretaba como "el supervisor sí
    escribió algo" y se aplicaba una corrección que nadie pidió.
    """
    if pd.isna(v):
        return False
    return str(v).strip() not in ("", "nan", "none", "None")


def _log_fila(log: list, archivo_respuesta: str, modulo: str, id_error: str,
             llave: str, campo: str, viejo, nuevo) -> None:
    log.append({
        "FECHA_APLICACION": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ARCHIVO_RESPUESTA": archivo_respuesta,
        "MODULO": modulo,
        "ID_ERROR": id_error,
        "LLAVE": llave,
        "CAMPO": campo,
        "VALOR_ANTERIOR": viejo,
        "VALOR_NUEVO": nuevo,
    })


# ─────────────────────────────────────────────────────────────────────────────
# PRECIOS / EXHIBICIONES — cruce 1 a 1 por ID_ERROR
# ─────────────────────────────────────────────────────────────────────────────

def aplicar_generico(headers: dict, site_id: str, filas: pd.DataFrame, cfg: dict,
                     modulo: str, log: list, archivo_respuesta: str) -> int:
    """
    `cfg["archivo"](mes, anio)` puede devolver:
      · una ruta completa (carpeta/nombre.xlsx) si `cfg` no tiene "carpeta"
      · un nombre suelto (se combina con `cfg["carpeta"]`) si sí la tiene
      · una LISTA de cualquiera de los dos — se prueban todas, por si el
        caso puede estar en más de un archivo candidato (ej. exhibiciones,
        que junta dos encuestas crudas distintas).
    """
    filas = filas[_no_vacio(filas[cfg["col_corregido"]]) & filas["ID_ERROR"].notna()].copy()
    if filas.empty:
        print(f"  ✓ {modulo}: nada corregido en esta respuesta.")
        return 0

    filas["_periodo"] = filas["ID_ERROR"].apply(_periodo_de_id)
    sin_periodo = filas["_periodo"].isna()
    if sin_periodo.any():
        print(f"  ⚠️  {int(sin_periodo.sum())} fila(s) con ID_ERROR sin el formato esperado — se saltan.")
    filas = filas[~sin_periodo]

    aplicados = 0
    for (mes, anio), grupo in filas.groupby("_periodo"):
        rutas = cfg["archivo"](mes, anio)
        if isinstance(rutas, str):
            rutas = [rutas]
        if "carpeta" in cfg:
            rutas = [f"{cfg['carpeta']}/{r}" for r in rutas]

        mapa_valor = {k: v for k, v in zip(grupo["ID_ERROR"].astype(str), grupo[cfg["col_corregido"]])
                     if _valor_no_vacio(v)}
        encontrado = False
        for ruta in rutas:
            carpeta, nombre = ruta.rsplit("/", 1)
            oficial = esp.leer_excel_cloud(headers, site_id, ruta, nombre, obligatorio=False)
            if oficial is None:
                continue
            if "ID_ERROR" not in oficial.columns or cfg["col_oficial"] not in oficial.columns:
                continue

            ids_oficial = oficial["ID_ERROR"].astype(str)
            cambia = ids_oficial.isin(mapa_valor) & (ids_oficial != "") & (ids_oficial != "nan")
            n = int(cambia.sum())
            if n == 0:
                continue
            encontrado = True

            # CORREGIDO marca, en la encuesta cruda, cuáles CASOS DE ERROR ya
            # se corrigieron a mano — solo las filas que tienen ID_ERROR (o
            # sea, que sí son un error a revisar); las filas que nunca
            # tuvieron error quedan en blanco, no en "No" (para no dar a
            # entender que a esas también había que corregirlas).
            if "CORREGIDO" not in oficial.columns:
                tiene_id = (oficial["ID_ERROR"].notna()
                           & (oficial["ID_ERROR"].astype(str).str.strip() != "")
                           & (oficial["ID_ERROR"].astype(str).str.strip() != "nan"))
                oficial["CORREGIDO"] = ""
                oficial.loc[tiene_id, "CORREGIDO"] = "No"

            # Defensa extra: pase lo que pase arriba en el filtro, nunca se
            # escribe un valor vacío encima de un dato real — el 22/09/2026
            # un archivo de prueba sin corregir sobrescribió 42 precios
            # reales con NaN (se recuperó del log). Este chequeo es la
            # última barrera aunque el filtro de más arriba falle.
            for idx in oficial[cambia].index:
                id_err = ids_oficial.at[idx]
                nuevo = mapa_valor[id_err]
                if not _valor_no_vacio(nuevo):
                    continue
                viejo = oficial.at[idx, cfg["col_oficial"]]
                oficial.at[idx, cfg["col_oficial"]] = nuevo
                oficial.at[idx, "CORREGIDO"] = "Sí"
                _log_fila(log, archivo_respuesta, modulo, id_err, id_err,
                         cfg["col_oficial"], viejo, nuevo)

            hoja = esp._nombre_hoja(headers, site_id, ruta)
            esp.subir_excel_cloud(headers, site_id, carpeta, nombre, {hoja: oficial}, embellecer=True)
            print(f"  ✅ {nombre}: {n} corrección(es) aplicada(s)")
            aplicados += n

        if not encontrado:
            nombres = ", ".join(r.rsplit("/", 1)[-1] for r in rutas)
            print(f"  ⚠️  Ningún ID_ERROR de la respuesta se encontró en: {nombres}.")

    return aplicados


# ─────────────────────────────────────────────────────────────────────────────
# ESPACIOS — cruce por llave natural (ver docstring del módulo)
# ─────────────────────────────────────────────────────────────────────────────

def aplicar_espacios(headers: dict, site_id: str, filas: pd.DataFrame,
                     log: list, archivo_respuesta: str) -> int:
    filas = filas.copy()
    tiene_marca = _no_vacio(filas["CM_MARCA_CORREGIDO"])
    tiene_universo = _no_vacio(filas["UNIVERSO_CORREGIDO"])
    filas = filas[(tiene_marca | tiene_universo) & filas["ID_ERROR"].notna()].copy()
    if filas.empty:
        print("  ✓ Espacios: nada corregido en esta respuesta.")
        return 0

    filas["_periodo"] = filas["ID_ERROR"].apply(_periodo_de_id)
    sin_periodo = filas["_periodo"].isna()
    if sin_periodo.any():
        print(f"  ⚠️  {int(sin_periodo.sum())} fila(s) con ID_ERROR sin el formato esperado — se saltan.")
    filas = filas[~sin_periodo]

    aplicados = 0
    for (mes, anio), grupo in filas.groupby("_periodo"):
        spec = pr.resolver(mes, anio)
        nombre = f"Encuesta_Sos_Consolidada_{spec.mes_str_upper}_{anio}.xlsx"
        carpeta = paths.RUTA_CARPETA_BASES_SOS
        ruta = f"{carpeta}/{nombre}"
        origen = esp.leer_excel_cloud(headers, site_id, ruta, nombre, obligatorio=False)
        if origen is None:
            print(f"  ⚠️  {nombre}: no existe, se salta.")
            continue

        col_universo, cols_marca_cm = esp._detectar_columnas(origen)
        origen["Empleado"] = origen["Empleado"].astype(str).str.strip()
        origen["PDV"] = origen["PDV"].astype(str).str.strip()
        origen["Marca"] = origen["Marca"].astype(str).str.strip()
        origen["_CATEGORIA"] = origen.apply(esp._categoria_desde_linea, axis=1)
        periodo_ok = ((pd.to_numeric(origen["Mes del año"], errors="coerce") == mes)
                     & (pd.to_numeric(origen["Año"], errors="coerce") == anio))

        # CORREGIDO marca, en la encuesta cruda, cuáles CASOS DE ERROR ya se
        # corrigieron a mano — ver misma idea en aplicar_generico. Ojo: en
        # espacios solo el motivo MOTIVO_UNIVERSO deja ID_ERROR
        # escrito en esta encuesta (ver etl_espacios_errores.py); por eso
        # acá también se usa "tiene ID_ERROR" como el filtro de qué filas
        # cuentan como caso de error.
        if "CORREGIDO" not in origen.columns:
            origen["CORREGIDO"] = ""
            if "ID_ERROR" in origen.columns:
                tiene_id = (origen["ID_ERROR"].notna()
                           & (origen["ID_ERROR"].astype(str).str.strip() != "")
                           & (origen["ID_ERROR"].astype(str).str.strip() != "nan"))
                origen.loc[tiene_id, "CORREGIDO"] = "No"

        n_este_archivo = 0
        for _, fila in grupo.iterrows():
            llave_txt = f"{fila['Empleado']}|{fila['PDV']}|{fila['LINEA_PRODUCTO']}|{fila['MARCA']}"

            if _valor_no_vacio(fila.get("CM_MARCA_CORREGIDO")):
                coincide = (periodo_ok
                           & (origen["Empleado"] == fila["Empleado"])
                           & (origen["PDV"] == fila["PDV"])
                           & (origen["_CATEGORIA"] == fila["LINEA_PRODUCTO"])
                           & (origen["Marca"] == fila["MARCA"]))
                idx = origen[coincide].index
                if len(idx) != 1:
                    print(f"  ⚠️  {fila['ID_ERROR']} ({fila['MARCA']}): {len(idx)} fila(s) "
                          f"coinciden en origen (se esperaba 1) — se salta.")
                else:
                    i = idx[0]
                    escrito = False
                    for c in cols_marca_cm:
                        if pd.notna(origen.at[i, c]):
                            viejo = origen.at[i, c]
                            nuevo = float(fila["CM_MARCA_CORREGIDO"])
                            origen.at[i, c] = nuevo
                            origen.at[i, "CORREGIDO"] = "Sí"
                            _log_fila(log, archivo_respuesta, "Espacios", fila["ID_ERROR"],
                                     llave_txt, c, viejo, nuevo)
                            n_este_archivo += 1
                            escrito = True
                            break
                    if not escrito:
                        print(f"  ⚠️  {fila['ID_ERROR']} ({fila['MARCA']}): la fila no tenía "
                              f"ningún valor de cm poblado — se salta.")

            if _valor_no_vacio(fila.get("UNIVERSO_CORREGIDO")):
                coincide_caso = (periodo_ok
                                 & (origen["Empleado"] == fila["Empleado"])
                                 & (origen["PDV"] == fila["PDV"])
                                 & (origen["_CATEGORIA"] == fila["LINEA_PRODUCTO"]))
                idx = origen[coincide_caso].index
                if len(idx) == 0:
                    print(f"  ⚠️  {fila['ID_ERROR']}: no se encontró el caso en origen "
                          f"— se salta la corrección de universo.")
                else:
                    nuevo = float(fila["UNIVERSO_CORREGIDO"])
                    viejo = origen.at[idx[0], col_universo]
                    origen.loc[idx, col_universo] = nuevo
                    origen.loc[idx, "CORREGIDO"] = "Sí"
                    _log_fila(log, archivo_respuesta, "Espacios", fila["ID_ERROR"],
                             llave_txt, col_universo, viejo, nuevo)
                    n_este_archivo += len(idx)

        if n_este_archivo:
            origen = origen.drop(columns=["_CATEGORIA"])
            hoja = esp._nombre_hoja(headers, site_id, ruta)
            esp.subir_excel_cloud(headers, site_id, carpeta, nombre, {hoja: origen}, embellecer=True)
            print(f"  ✅ {nombre}: {n_este_archivo} corrección(es) aplicada(s)")
        aplicados += n_este_archivo

    return aplicados


# ─────────────────────────────────────────────────────────────────────────────
# LOG DE AUDITORÍA
# ─────────────────────────────────────────────────────────────────────────────

def _guardar_log(headers: dict, site_id: str, filas_nuevas: list) -> None:
    carpeta = paths._SALIDAS_ROOT
    nombre = "LOG_CORRECCIONES.xlsx"
    previo = esp.leer_excel_cloud(headers, site_id, f"{carpeta}/{nombre}", nombre,
                                  obligatorio=False)
    nuevo_df = pd.DataFrame(filas_nuevas)
    combinado = pd.concat([previo, nuevo_df], ignore_index=True) if previo is not None else nuevo_df
    esp.subir_excel_cloud(headers, site_id, carpeta, nombre, {"Log": combinado})
    print(f"\n📝 Log actualizado: {len(filas_nuevas)} fila(s) nueva(s) en {nombre} "
          f"({len(combinado)} en total)")


# ─────────────────────────────────────────────────────────────────────────────
# PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def _procesar_contenido(headers: dict, site_id: str, contenido: bytes,
                        ruta_respuesta: str) -> int:
    """
    Aplica las correcciones de un archivo de respuesta ya descargado.
    Devuelve el total aplicado.

    El módulo de cada hoja se reconoce por sus COLUMNAS, no por el nombre
    de la hoja — así funciona igual con el archivo real que arma
    `libro_del_supervisor()` (hojas "Precios"/"Exhibiciones"/"Espacios") y
    con un archivo suelto de cualquier ETL de errores (que siempre nombra
    su hoja "Errores"), sin tener que renombrar nada a mano.
    """
    xls = pd.ExcelFile(io.BytesIO(contenido))
    print(f"  Hojas encontradas: {xls.sheet_names}")

    log: list = []
    total = 0

    for hoja in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=hoja)
        if "PRECIO_CORREGIDO" in df.columns:
            print(f"\nPrecios (hoja '{hoja}'):")
            total += aplicar_generico(headers, site_id, df, CFG_PRECIOS, "Precios", log, ruta_respuesta)
        elif "CANTIDAD_CORREGIDA" in df.columns:
            print(f"\nExhibiciones (hoja '{hoja}'):")
            total += aplicar_generico(headers, site_id, df, CFG_EXHIBICIONES, "Exhibiciones", log, ruta_respuesta)
        elif "CM_MARCA_CORREGIDO" in df.columns or "UNIVERSO_CORREGIDO" in df.columns:
            print(f"\nEspacios (hoja '{hoja}'):")
            total += aplicar_espacios(headers, site_id, df, log, ruta_respuesta)
        else:
            print(f"  ⚠️  Hoja '{hoja}': no se reconoce ningún módulo por sus columnas — se salta.")

    if log:
        _guardar_log(headers, site_id, log)

    return total


def run_archivo(ruta_respuesta: str) -> int:
    """Procesa UN archivo de respuesta puntual (uso manual / de prueba)."""
    print("\n" + "=" * 60)
    print(f"  APLICAR CORRECCIONES — {ruta_respuesta}")
    print("=" * 60)

    headers = {"Authorization": f"Bearer {esp.obtener_token_azure()}"}
    site_id = esp.obtener_site_id(headers)

    print("\nLeyendo la respuesta:")
    contenido = _descargar_bytes(headers, site_id, ruta_respuesta)
    total = _procesar_contenido(headers, site_id, contenido, ruta_respuesta)

    print(f"\n{'=' * 60}")
    print(f"  {total} corrección(es) aplicada(s) en total.")
    print("=" * 60)
    return 0


def run_carpeta(ruta_carpeta: str) -> int:
    """
    Modo cron: revisa TODOS los archivos de respuesta pendientes en la
    carpeta (los que Power Automate va dejando) y los procesa uno por uno.
    Un archivo que falla no frena a los demás — queda para revisar a mano y
    la corrida sigue con el resto.
    """
    print("\n" + "=" * 60)
    print(f"  APLICAR CORRECCIONES — carpeta {ruta_carpeta}")
    print("=" * 60)

    headers = {"Authorization": f"Bearer {esp.obtener_token_azure()}"}
    site_id = esp.obtener_site_id(headers)

    archivos = esp.obtener_archivos_carpeta(headers, site_id, ruta_carpeta)
    pendientes = [a for a in archivos
                 if a["name"].lower().endswith(".xlsx")
                 and not a["name"].startswith("PROCESADO_")
                 and not a["name"].startswith("~$")]

    print(f"\n{len(pendientes)} archivo(s) pendiente(s) de {len(archivos)} en la carpeta.")
    if not pendientes:
        print("  ✓ Nada por procesar.")
        return 0

    total_general = 0
    fallos = 0
    for item in pendientes:
        ruta_item = f"{ruta_carpeta}/{item['name']}"
        print(f"\n{'-' * 60}")
        print(f"  {item['name']}")
        print("-" * 60)
        try:
            contenido = requests.get(item["@microsoft.graph.downloadUrl"]).content
            total = _procesar_contenido(headers, site_id, contenido, ruta_item)
            print(f"  {total} corrección(es) aplicada(s) — marcando como procesado.")
            _marcar_procesado(headers, site_id, item["id"], item["name"])
            total_general += total
        except Exception as e:
            fallos += 1
            print(f"  ❌ Error procesando {item['name']}: {e}")
            print(f"     No se marca como procesado — queda pendiente para revisar.")

    print(f"\n{'=' * 60}")
    print(f"  {total_general} corrección(es) aplicada(s) en total. "
          f"{len(pendientes) - fallos} archivo(s) OK, {fallos} fallido(s).")
    print("=" * 60)
    return 1 if fallos else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawTextHelpFormatter)
    grupo = ap.add_mutually_exclusive_group(required=True)
    grupo.add_argument("--archivo",
                       help="Ruta en SharePoint (relativa a Documentos) de UN Excel de respuesta")
    grupo.add_argument("--carpeta",
                       help="Ruta en SharePoint de una carpeta con varios archivos de respuesta "
                            "pendientes (modo cron)")
    args = ap.parse_args()
    if args.archivo:
        return run_archivo(args.archivo)
    return run_carpeta(args.carpeta)


if __name__ == "__main__":
    sys.exit(main())
