"""
shared_loader.py
────────────────
Lee los archivos fuente del Plan de Trabajo UNA SOLA VEZ y produce
DataFrames filtrados listos para cada ETL.

Ver docstring original para descripción de la interfaz.
"""

import os
import glob
import urllib.parse
import requests
import pandas as pd
import numpy as np
from etl_logger import get_logger


# Proxy lazy: difiere get_logger() hasta el primer uso real.
# Necesario para que run_all.py pueda llamar inicializar_logging(ruta_log=...)
# ANTES de que shared_loader registre handlers con la ruta por defecto.
class _LazyLogger:
    def __getattr__(self, attr):
        return getattr(get_logger("shared_loader"), attr)

log = _LazyLogger()

# ─────────────────────────────────────────────────────────────────────────────
COLUMNAS_ESTANDAR_UNIFICADO = {
    "Ventas Promedio Mes Oct-Dic 2024" : "VENTAS_PROMEDIO_MES",
    "Ventas Promedio Mes"              : "VENTAS_PROMEDIO_MES",
    "ID PDV INVOLVES"                  : "ID_PDV_INVOLVES",
    "Nombre Punto de Venta"            : "NOMBRE_PDV",
    "Nombre punto de venta"            : "NOMBRE_PDV",
    "ACRONIMO"                         : "ACRONIMO",
    "Acronimo"                         : "ACRONIMO",
    "Cedula"                           : "CEDULA",
    "Nombre"                           : "NOMBRE",
    "Cod Mercaderista"                 : "COD_MERCADERISTA",
    "Fecha"                            : "FECHA",
    "Hora Inicio"                      : "HORA_INICIO",
    "Hora fin"                         : "HORA_FIN",
    "Hora Fin"                         : "HORA_FIN",
    "Tiempo de servicio"               : "TIEMPO_SERVICIO",
    "Frec Semanal"                     : "FREC_SEMANAL",
    "Frec Mensual"                     : "FREC_MENSUAL",
    "Horas x visita"                   : "HORAS_X_VISITA",
    "Horas Mes"                        : "HORAS_MES",
    "Rol"                              : "ROL",
    "Supervisor Lider"                 : "SUPERVISOR_LIDER",
    "VENTAS PROMEDIO MES OCT-DIC 2024" : "VENTAS_PROMEDIO_MES",
    "VENTAS PROMEDIO MES"              : "VENTAS_PROMEDIO_MES",
    "NOMBRE PUNTO DE VENTA"            : "NOMBRE_PDV",
    "CEDULA"                           : "CEDULA",
    "NOMBRE"                           : "NOMBRE",
    "COD MERCADERISTA"                 : "COD_MERCADERISTA",
    "FECHA"                            : "FECHA",
    "HORA INICIO"                      : "HORA_INICIO",
    "HORA FIN"                         : "HORA_FIN",
    "TIEMPO DE SERVICIO"               : "TIEMPO_SERVICIO",
    "FREC SEMANAL"                     : "FREC_SEMANAL",
    "FREC MENSUAL"                     : "FREC_MENSUAL",
    "HORAS X VISITA"                   : "HORAS_X_VISITA",
    "HORAS MES"                        : "HORAS_MES",
    "ROL"                              : "ROL",
    "SUPERVISOR LIDER"                 : "SUPERVISOR_LIDER",
}

COLUMNAS_SUPERSET = [
    "ID_PDV_INVOLVES", "NOMBRE_PDV", "VENTAS_PROMEDIO_MES", "ACRONIMO",
    "CEDULA", "NOMBRE", "COD_MERCADERISTA", "FECHA", "HORA_INICIO",
    "HORA_FIN", "TIEMPO_SERVICIO", "FREC_SEMANAL", "FREC_MENSUAL",
    "HORAS_X_VISITA", "HORAS_MES", "ROL", "SUPERVISOR_LIDER", "FUENTE",
]

COLUMNAS_NUMERICAS = [
    "TIEMPO_SERVICIO", "FREC_SEMANAL", "FREC_MENSUAL",
    "HORAS_X_VISITA", "HORAS_MES", "VENTAS_PROMEDIO_MES",
]

MODULOS = {
    "no_presencia" : {"DIRECTO": "NO PRESENCIA",    "ISM": "NO PRESENCIA_FINAL"},
    "precios"      : {"DIRECTO": "PRECIOS",          "ISM": "PRECIOS_FINAL"},
    "sos"          : {"DIRECTO": "ESPACIOS",         "ISM": "ESPACIOS_FINAL"},
}


# ─────────────────────────────────────────────────────────────────────────────
# NORMALIZACIÓN DE IDs DE PDV  (helper central — usar SIEMPRE para cruces)
# ─────────────────────────────────────────────────────────────────────────────

def calcular_kpi_simple_y_escribir(
    df_origen: pd.DataFrame,
    col_planeado: str,
    col_ejecutado: str,
    ruta_kpi_mes: str,
    ruta_kpi_historico: str,
    nombre_cumplimiento: str = "CUMPLIMIENTO",
    nombres_renombre: tuple = ("PLANEADO", "EJECUTADO"),
) -> int:
    import numpy as np

    df = df_origen.copy()
    df[col_planeado]  = pd.to_numeric(df[col_planeado],  errors='coerce').fillna(0)
    df[col_ejecutado] = pd.to_numeric(df[col_ejecutado], errors='coerce').fillna(0)

    if 'MES' not in df.columns or 'AÑO' not in df.columns:
        faltan = [c for c in ('MES', 'AÑO') if c not in df.columns]
        raise ValueError(
            f"calcular_kpi_simple_y_escribir: df_origen no tiene columnas {faltan}. "
            f"El ETL que llama debe propagar MES y AÑO desde el periodo procesado. "
            f"Ruta destino: {ruta_kpi_mes}"
        )
    df['MES'] = pd.to_numeric(df['MES'], errors='coerce').astype('Int64')
    df['AÑO'] = pd.to_numeric(df['AÑO'], errors='coerce').astype('Int64')
    invalidos = df[
        df['MES'].isna() | df['AÑO'].isna()
        | (df['MES'] < 1) | (df['MES'] > 12)
        | (df['AÑO'] < 2020) | (df['AÑO'] > 2099)
    ]
    if not invalidos.empty:
        n_total, n_mal = len(df), len(invalidos)
        muestra = invalidos[['MES', 'AÑO']].head(3).to_dict('records')
        raise ValueError(
            f"calcular_kpi_simple_y_escribir: {n_mal}/{n_total} filas tienen "
            f"MES/AÑO inválidos (esperado MES 1-12, AÑO 2020-2099). "
            f"Muestra: {muestra}. Ruta destino: {ruta_kpi_mes}"
        )
    df['MES'] = df['MES'].astype(int)
    df['AÑO'] = df['AÑO'].astype(int)

    if 'SUPERVISOR_LIDER' not in df.columns:
        df['SUPERVISOR_LIDER'] = ""
    if 'NOMBRE' not in df.columns:
        df['NOMBRE'] = ""

    resumen = (
        df.groupby(['MES', 'AÑO', 'SUPERVISOR_LIDER', 'NOMBRE'], dropna=False)
          .agg(**{col_planeado:  (col_planeado,  'sum'),
                  col_ejecutado: (col_ejecutado, 'sum')})
          .reset_index()
    )

    nom_plan, nom_eje = nombres_renombre
    resumen[nombre_cumplimiento] = np.where(
        resumen[col_planeado] != 0,
        resumen[col_ejecutado] / resumen[col_planeado],
        0,
    )
    resumen = resumen.rename(columns={col_planeado: nom_plan, col_ejecutado: nom_eje})

    cols_final = ['MES', 'AÑO', 'SUPERVISOR_LIDER', 'NOMBRE',
                  nom_plan, nom_eje, nombre_cumplimiento]
    resumen = resumen[cols_final].sort_values(
        ['AÑO', 'MES', 'SUPERVISOR_LIDER', 'NOMBRE']
    ).reset_index(drop=True)

    if not resumen.empty:
        anio_max = int(resumen['AÑO'].max())
        mes_max  = int(resumen[resumen['AÑO'] == anio_max]['MES'].max())
        mask = (resumen['AÑO'] == anio_max) & (resumen['MES'] == mes_max)
        resumen_activo = resumen[mask].copy()
    else:
        resumen_activo = resumen.copy()

    os.makedirs(os.path.dirname(str(ruta_kpi_mes)), exist_ok=True)
    resumen_activo.to_excel(str(ruta_kpi_mes), index=False, engine='openpyxl')

    if os.path.exists(str(ruta_kpi_historico)):
        try:
            hist_prev = pd.read_excel(str(ruta_kpi_historico), engine='openpyxl')
            hist_prev['MES'] = pd.to_numeric(hist_prev.get('MES'), errors='coerce').astype('Int64')
            hist_prev['AÑO'] = pd.to_numeric(hist_prev.get('AÑO'), errors='coerce').astype('Int64')
            valido = (
                hist_prev['MES'].notna() & hist_prev['AÑO'].notna()
                & hist_prev['MES'].between(1, 12)
                & hist_prev['AÑO'].between(2020, 2099)
            )
            n_basura = (~valido).sum()
            if n_basura > 0:
                log.warning(
                    f"Histórico '{os.path.basename(str(ruta_kpi_historico))}': "
                    f"se descartan {n_basura} filas con MES/AÑO inválido."
                )
            hist_prev = hist_prev[valido].copy()
            hist_prev['MES'] = hist_prev['MES'].astype(int)
            hist_prev['AÑO'] = hist_prev['AÑO'].astype(int)
        except Exception:
            hist_prev = pd.DataFrame(columns=cols_final)
    else:
        hist_prev = pd.DataFrame(columns=cols_final)

    if not hist_prev.empty and not resumen.empty:
        periodos_nuevos = set(map(tuple,
            resumen[['MES', 'AÑO']].drop_duplicates().values.tolist()))
        mask_keep = ~hist_prev.apply(
            lambda r: (int(r['MES']), int(r['AÑO'])) in periodos_nuevos,
            axis=1,
        )
        hist_prev = hist_prev[mask_keep]

    hist_final = pd.concat([hist_prev, resumen], ignore_index=True)
    hist_final = hist_final.sort_values(['AÑO', 'MES', 'SUPERVISOR_LIDER', 'NOMBRE'])
    hist_final.to_excel(str(ruta_kpi_historico), index=False, engine='openpyxl')

    return len(resumen_activo)


def id_a_str(serie: pd.Series) -> pd.Series:
    s = pd.to_numeric(serie, errors="coerce")
    salida = s.where(s.notna(), pd.NA).astype("Int64").astype(str)
    return salida.replace("<NA>", "")


# ─────────────────────────────────────────────────────────────────────────────
# DESCARGA DINÁMICA DESDE SHAREPOINT (CLOUD / GITHUB ACTIONS)
# ─────────────────────────────────────────────────────────────────────────────

def _descargar_pt_desde_sharepoint(periodo) -> tuple[str, str]:
    """
    Descarga los archivos de Plan de Trabajo desde SharePoint usando Graph API
    cuando la ruta local no existe (ej. GitHub Actions).
    """
    try:
        import msal
        import paths
        
        tenant_id = os.environ.get("AZURE_TENANT_ID")
        client_id = os.environ.get("AZURE_CLIENT_ID")
        client_secret = os.environ.get("AZURE_CLIENT_SECRET")
        
        if not tenant_id or not client_id or not client_secret:
            log.error("Faltan credenciales de Azure (AZURE_TENANT_ID, CLIENT_ID, CLIENT_SECRET)")
            return "", ""

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
        token_res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" not in token_res:
            log.error(f"No se pudo obtener token de acceso de Azure: {token_res.get('error_description')}")
            return "", ""

        headers = {"Authorization": f"Bearer {token_res['access_token']}"}
        
        # Obtener Site ID
        res_site = requests.get(f"https://graph.microsoft.com/v1.0/sites/root:/sites/{paths.SHAREPOINT_SITE_NAME}", headers=headers).json()
        site_id = res_site.get("id")
        if not site_id:
            log.error(f"No se pudo resolver el site_id para SharePoint: {paths.SHAREPOINT_SITE_NAME}")
            return "", ""

        ruta_formateada = urllib.parse.quote(paths.RUTA_CARPETA_PT)
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_formateada}:/children"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            log.error(f"Error al listar SharePoint en {paths.RUTA_CARPETA_PT} ({response.status_code}): {response.text}")
            return "", ""
            
        archivos = response.json().get("value", [])
        token_periodo = f"{periodo.mes_str} {periodo.anio}".lower() if periodo else ""

        archivos_validos = [
            a for a in archivos 
            if not a["name"].startswith("~$") and "consolidado" not in a["name"].lower()
        ]

        candidatos_dir = [a for a in archivos_validos if "directo" in a["name"].lower()]
        candidatos_ism = [a for a in archivos_validos if "ism" in a["name"].lower()]

        if periodo:
            candidatos_dir = [a for a in candidatos_dir if token_periodo in a["name"].lower()]
            candidatos_ism = [a for a in candidatos_ism if token_periodo in a["name"].lower()]

        url_dir = candidatos_dir[0]["@microsoft.graph.downloadUrl"] if candidatos_dir else None
        url_ism = candidatos_ism[0]["@microsoft.graph.downloadUrl"] if candidatos_ism else None

        os.makedirs("/tmp/pt_temp", exist_ok=True)
        ruta_dir_local = ""
        ruta_ism_local = ""

        if url_dir:
            content = requests.get(url_dir).content
            ruta_dir_local = f"/tmp/pt_temp/{candidatos_dir[0]['name']}"
            with open(ruta_dir_local, "wb") as f:
                f.write(content)
            log.info(f"PT Directo descargado desde SharePoint: {candidatos_dir[0]['name']}")
                
        if url_ism:
            content = requests.get(url_ism).content
            ruta_ism_local = f"/tmp/pt_temp/{candidatos_ism[0]['name']}"
            with open(ruta_ism_local, "wb") as f:
                f.write(content)
            log.info(f"PT ISM descargado desde SharePoint: {candidatos_ism[0]['name']}")

        return ruta_dir_local, ruta_ism_local

    except Exception as e:
        log.error(f"Excepción al descargar PT desde SharePoint: {e}", exc_info=True)
        return "", ""


# ─────────────────────────────────────────────────────────────────────────────
# DESCUBRIMIENTO DINÁMICO DE ARCHIVOS DE PT
# ─────────────────────────────────────────────────────────────────────────────

def descubrir_archivos_pt(directorio_pt: str, periodo=None) -> tuple[str, str]:
    """
    Localiza los archivos de Plan de Trabajo Directo e ISM dentro de
    `directorio_pt`. Si el directorio físico no existe (ej. en GitHub Actions),
    intenta descargarlos automáticamente desde SharePoint.
    """
    if not os.path.isdir(directorio_pt):
        log.warning(
            f"Directorio local de PT no existe: {directorio_pt} "
            f"— intentando descarga remota desde SharePoint..."
        )
        return _descargar_pt_desde_sharepoint(periodo)

    todos = glob.glob(os.path.join(directorio_pt, "*.xlsx"))
    todos = [
        p for p in todos
        if not os.path.basename(p).startswith("~$")
        and "consolidado" not in os.path.basename(p).lower()
    ]

    candidatos_dir = [p for p in todos if "directo" in os.path.basename(p).lower()]
    candidatos_ism = [p for p in todos if "ism"     in os.path.basename(p).lower()]

    if periodo is not None:
        token = f"{periodo.mes_str} {periodo.anio}".lower()
        candidatos_dir = [
            p for p in candidatos_dir
            if token in os.path.basename(p).lower()
        ]
        candidatos_ism = [
            p for p in candidatos_ism
            if token in os.path.basename(p).lower()
        ]
        if len(candidatos_dir) > 1:
            raise RuntimeError(
                f"Múltiples PT Directo coinciden con periodo {token!r} en "
                f"{directorio_pt}: {[os.path.basename(p) for p in candidatos_dir]}"
            )
        if len(candidatos_ism) > 1:
            raise RuntimeError(
                f"Múltiples PT ISM coinciden con periodo {token!r} en "
                f"{directorio_pt}: {[os.path.basename(p) for p in candidatos_ism]}"
            )
        ruta_dir = candidatos_dir[0] if candidatos_dir else ""
        ruta_ism = candidatos_ism[0] if candidatos_ism else ""
    else:
        log.warning(
            "descubrir_archivos_pt llamada sin parámetro `periodo` — "
            "uso el más reciente por mtime (comportamiento legacy)."
        )
        ruta_dir = max(candidatos_dir, key=os.path.getmtime) if candidatos_dir else ""
        ruta_ism = max(candidatos_ism, key=os.path.getmtime) if candidatos_ism else ""

    if not ruta_dir:
        suf = f" para periodo {periodo}" if periodo is not None else ""
        log.warning(f"No se encontró archivo *Directo*.xlsx{suf} en {directorio_pt}")
    else:
        log.info(f"PT Directo detectado: {os.path.basename(ruta_dir)}")

    if not ruta_ism:
        suf = f" para periodo {periodo}" if periodo is not None else ""
        log.warning(f"No se encontró archivo *ISM*.xlsx{suf} en {directorio_pt}")
    else:
        log.info(f"PT ISM detectado: {os.path.basename(ruta_ism)}")

    return ruta_dir, ruta_ism


# ─────────────────────────────────────────────────────────────────────────────

def _leer_archivo_pt(ruta: str, hoja_pt: str, fuente: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not ruta or not os.path.exists(ruta):
        log.critical(
            f"Archivo Plan de Trabajo no encontrado: {ruta} "
            f"— fuente '{fuente}' excluida del pipeline"
        )
        return pd.DataFrame(), pd.DataFrame()

    try:
        with pd.ExcelFile(ruta) as xls:
            df_mod = pd.read_excel(xls, sheet_name="Captura de modulos")
            df_mod.columns = df_mod.columns.str.strip().str.upper()

            if df_mod.empty:
                log.warning(
                    f"[{fuente}] Hoja 'Captura de modulos' vacía en {os.path.basename(ruta)} "
                    f"— los filtros de módulo no se aplicarán para esta fuente"
                )
            elif "ID PDV INVOLVES" not in df_mod.columns:
                log.warning(
                    f"[{fuente}] Columna 'ID PDV INVOLVES' ausente en 'Captura de modulos' "
                    f"({os.path.basename(ruta)}) — filtros de módulo desactivados para esta fuente"
                )
            else:
                df_mod["ID PDV INVOLVES"] = id_a_str(df_mod["ID PDV INVOLVES"])

            df_pt = pd.read_excel(xls, sheet_name=hoja_pt)
            df_pt.columns = df_pt.columns.str.strip()
            df_pt.rename(
                columns={col: COLUMNAS_ESTANDAR_UNIFICADO[col]
                         for col in df_pt.columns if col in COLUMNAS_ESTANDAR_UNIFICADO},
                inplace=True,
            )
            df_pt["FUENTE"] = fuente

            for col in COLUMNAS_NUMERICAS:
                if col in df_pt.columns:
                    df_pt[col] = (
                        df_pt[col].astype(str).str.strip().str.replace(",", ".", regex=False)
                    )
                    df_pt[col] = pd.to_numeric(df_pt[col], errors="coerce")

            if "ID_PDV_INVOLVES" in df_pt.columns:
                df_pt["ID_PDV_INVOLVES"] = id_a_str(df_pt["ID_PDV_INVOLVES"])

            if "FECHA" in df_pt.columns:
                df_pt["FECHA"] = (
                    pd.to_datetime(df_pt["FECHA"], errors="coerce").dt.strftime("%d/%m/%Y")
                )

            for col in COLUMNAS_SUPERSET:
                if col not in df_pt.columns:
                    df_pt[col] = ""

            for col in df_pt.columns:
                if col not in COLUMNAS_NUMERICAS and col != "FECHA":
                    if df_pt[col].dtype == object:
                        df_pt[col] = (
                            df_pt[col].astype(str).str.strip()
                            .replace(["nan", "NaT", "None"], "")
                        )

            df_pt = df_pt[[c for c in COLUMNAS_SUPERSET if c in df_pt.columns]].copy()

    except Exception:
        log.error(
            f"[{fuente}] Excepción al leer {os.path.basename(ruta)} — "
            f"esta fuente se excluye del pipeline",
            exc_info=True,
        )
        return pd.DataFrame(), pd.DataFrame()

    log.info(f"[{fuente}] {len(df_pt):,} filas PT  |  {len(df_mod):,} filas módulos cargadas")
    return df_pt, df_mod


def _filtrar_por_modulo(
    df_pt: pd.DataFrame,
    df_mod: pd.DataFrame,
    fuente: str,
    modulo: str,
) -> pd.DataFrame:
    col = MODULOS[modulo][fuente]

    if col not in df_mod.columns:
        log.warning(
            f"[{fuente}] Columna de módulo '{col}' no encontrada en 'Captura de modulos' "
            f"— módulo '{modulo}' no recibirá filas de esta fuente"
        )
        return pd.DataFrame(columns=df_pt.columns)

    ids_validos = df_mod.loc[df_mod[col] == 1, "ID PDV INVOLVES"].unique()
    df_filtrado = df_pt[df_pt["ID_PDV_INVOLVES"].isin(ids_validos)].copy()

    if df_filtrado.empty:
        log.warning(
            f"[{fuente}] Módulo '{modulo}': el filtro devolvió 0 filas "
            f"(columna '{col}' sin valores = 1, o IDs sin coincidencia en el PT)"
        )
    else:
        log.info(f"[{fuente}] Módulo '{modulo}': {len(df_filtrado):,} filas retenidas")

    return df_filtrado


def _detectar_periodo(df_pt: pd.DataFrame) -> tuple[int, int]:
    from datetime import datetime

    if "FECHA" not in df_pt.columns or df_pt["FECHA"].isna().all():
        log.warning(
            "Columna FECHA ausente o completamente nula en el PT consolidado "
            "— se usará la fecha actual del sistema como fallback"
        )
        ahora = datetime.now()
        return ahora.month, ahora.year

    try:
        fechas = pd.to_datetime(df_pt["FECHA"], dayfirst=True, errors="coerce").dropna()
        if fechas.empty:
            raise ValueError("Todas las fechas son NaT tras el parseo")
        mes  = int(fechas.dt.month.mode()[0])
        anio = int(fechas.dt.year.mode()[0])
        log.info(f"Periodo detectado en el PT: Mes {mes} | Año {anio}")
        return mes, anio
    except Exception as e:
        ahora = datetime.now()
        log.warning(
            f"No se pudo detectar el periodo del PT ({e}) "
            f"— usando fecha del sistema: {ahora.month}/{ahora.year}"
        )
        return ahora.month, ahora.year


# ─────────────────────────────────────────────────────────────────────────────

def cargar_plan_de_trabajo(
    ruta_directo: str,
    ruta_ism: str,
    *,
    mes: int | None = None,
    anio: int | None = None,
) -> dict:
    log.info("━" * 50)
    log.info("Iniciando carga compartida del Plan de Trabajo")
    log.info("━" * 50)

    df_dir_pt, df_dir_mod = _leer_archivo_pt(ruta_directo, "Plan de trabajo",     "DIRECTO")
    df_ism_pt, df_ism_mod = _leer_archivo_pt(ruta_ism,     "Plan de trabajo CIF", "ISM")

    partes = [df for df in [df_dir_pt, df_ism_pt] if not df.empty]
    if not partes:
        log.critical(
            "No se pudo cargar ninguna fuente del Plan de Trabajo "
            "(ni DIRECTO ni ISM) — el pipeline no puede continuar"
        )
        raise RuntimeError(
            "shared_loader: ambas fuentes del PT fallaron. "
            "Revisa las rutas y el acceso a los archivos."
        )

    if len(partes) == 1:
        fuente_faltante = "ISM" if df_dir_pt.empty else "DIRECTO"
        log.warning(
            f"Solo se cargó una fuente del PT ({fuente_faltante} not available) "
            f"— los resultados estarán incompletos"
        )

    df_completo = pd.concat(partes, ignore_index=True)
    log.info(f"PT consolidado: {len(df_completo):,} filas totales")

    if mes is not None and anio is not None:
        mes, anio = int(mes), int(anio)
        log.info(f"Periodo forzado por el caller: Mes {mes} | Año {anio}")
    else:
        mes, anio = _detectar_periodo(df_completo)

    resultado = {
        "cif"          : df_completo,
        "periodo_mes"  : mes,
        "periodo_anio" : anio,
    }

    log.info("Aplicando filtros por módulo")
    for modulo in MODULOS:
        partes_modulo = []
        for fuente, df_pt, df_mod in [
            ("DIRECTO", df_dir_pt, df_dir_mod),
            ("ISM",     df_ism_pt, df_ism_mod),
        ]:
            if df_pt.empty:
                continue
            df_filtrado = _filtrar_por_modulo(df_pt, df_mod, fuente, modulo)
            if not df_filtrado.empty:
                partes_modulo.append(df_filtrado)

        df_modulo = (
            pd.concat(partes_modulo, ignore_index=True)
            if partes_modulo
            else pd.DataFrame(columns=df_completo.columns)
        )

        if df_modulo.empty:
            log.warning(
                f"Módulo '{modulo}': DataFrame final vacío (ambas fuentes sin datos) "
                f"— el ETL correspondiente se ejecutará pero sin filas de PT"
            )

        resultado[modulo] = df_modulo

    log.info("Carga compartida completada")
    return resultado
