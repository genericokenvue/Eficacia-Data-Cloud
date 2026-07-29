"""
run_all.py
──────────
Orquestador del conjunto de ETLs de Eficacia (canal droguerías).

Ver docstring original para descripción de la arquitectura.

Uso
───
    python run_all.py --mes 5 --anio 2026                         # Un mes
    python run_all.py --periodos 03/2026,04/2026,05/2026          # Multi-periodo
    python run_all.py --mes 5 --anio 2026 --solo cif nopresencia  # Subset
    python run_all.py --mes 5 --anio 2026 --workers 3             # Limitar threads
    python run_all.py --mes 5 --anio 2026 --log-dir ./mis_logs    # Carpeta de logs

Sprint 16.1: --mes/--anio o --periodos son REQUERIDOS (consistente con D1).
"""

import sys
import time
import argparse
import traceback
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import pandas as pd

from etl_logger import (
    inicializar_logging, get_logger,
    obtener_eventos_acumulados, ruta_log_actual,
    separador,
)
import paths
from shared_loader import cargar_plan_de_trabajo, descubrir_archivos_pt as _descubrir_pt_con_periodo
import periodo_resolver as pr
import run_log as rl

# ─────────────────────────────────────────────────────────────────────────────
# RUTAS CENTRALIZADAS — todas vienen de paths.py (auto detecta dev/prod)
# ─────────────────────────────────────────────────────────────────────────────
BASE   = str(paths.BASE)
DIR_PT = str(paths.CIF_PT_DIR)

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS DE ETLs
# ─────────────────────────────────────────────────────────────────────────────
import etl_cif
import etl_nopresencia
import etl_precios
import etl_sos
import etl_exhibiciones_pagadas
import etl_exhibiciones_gratis
import etl_ventas
import etl_impactos
import etl_impactos_segmentos


# ─────────────────────────────────────────────────────────────────────────────
# WRAPPERS CON LOGGING
# ─────────────────────────────────────────────────────────────────────────────

def _run_cif(df_pt: pd.DataFrame, mes: int, anio: int) -> None:
    log = get_logger("cif")
    separador(log)
    log.info("ETL CIF — iniciando")

    import etl_cif
    import periodo_resolver as pr

    spec = pr.resolver(int(mes), int(anio))
    
    # Autenticación y obtención de parámetros mediante Graph API igual que los demás módulos
    token = etl_cif.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = etl_cif.obtener_site_id(headers)

    log.info(f"Paso 1 — PT consolidado: {len(df_pt):,} filas")

    # Ejecución secuencial de los pasos internos de CIF usando el spec
    etl_cif.ejecutar_paso_1(spec, headers, site_id)
    etl_cif.ejecutar_paso_2(spec, headers, site_id)
    etl_cif.ejecutar_paso_3(spec, headers, site_id)
    etl_cif.ejecutar_paso_4(spec, headers, site_id)
    etl_cif.ejecutar_paso_5(spec, headers, site_id)
    etl_cif.ejecutar_paso_6(spec, headers, site_id)
    etl_cif.ejecutar_paso_7(spec, headers, site_id)
    etl_cif.generar_resumen_kpis_8(spec, headers, site_id)
    
    log.info("ETL CIF — completado")

def _run_nopresencia(df_pt: pd.DataFrame, mes: int, anio: int) -> None:
    log = get_logger("nopresencia")
    separador(log)
    log.info("ETL No Presencia — iniciando")

    import etl_nopresencia
    import periodo_resolver as pr

    if df_pt.empty:
        log.error(
            "DataFrame de PT vacío para módulo 'no_presencia' "
            "— no hay PDVs con módulo activo. ETL abortado."
        )
        return

    spec = pr.resolver(int(mes), int(anio))
    periodo_final = f"{etl_nopresencia.MESES_ESPANOL[mes]}_{anio}"

    # Autenticación requerida por los pasos nativos de etl_nopresencia
    token = etl_nopresencia.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = etl_nopresencia.obtener_site_id(headers)

    # 1. Consolidar PT (o usar el DataFrame recibido si ya está cargado)
    if df_pt.empty:
        df_pt, periodo_final = etl_nopresencia.ejecutar_paso_1_consolidar_pt_np(spec, headers, site_id)
    else:
        log.info(f"Usando DataFrame de PT proporcionado por el orquestador ({len(df_pt):,} filas)")

    if df_pt.empty:
        log.error("No se pudo obtener el Plan de Trabajo para No Presencia. ETL abortado.")
        return

    # 2. Procesar encuestas de No Presencia
    df_np_proc = etl_nopresencia.procesar_no_presencia(periodo_final, spec, headers, site_id)
    if df_np_proc.empty:
        log.error("No se encontraron archivos de encuesta de No Presencia válidos.")
        return
    log.info("Encuestas de No Presencia procesadas")

    # 3. Procesar PT por semanas
    df_pt_proc = etl_nopresencia.procesar_plan_de_trabajo_semanas(df_pt, periodo_final)
    if df_pt_proc.empty:
        log.error("Fallo al procesar el Plan de Trabajo por semanas — matriz no generada")
        return
    log.info("Plan de Trabajo por semanas procesado")

    # 4. Generar matriz de seguimiento comparativa
    etl_nopresencia.generar_matriz_seguimiento(df_np_proc, df_pt_proc, periodo_final, spec=spec, headers=headers, site_id=site_id)
    log.info("Matriz de seguimiento generada")

    # 5. Análisis de agotados
    etl_nopresencia.generar_analisis_agotados(periodo_final, headers, site_id)
    log.info("Análisis de agotados generado")

    # 6. KPI y carga a Supabase
    etl_nopresencia.generar_resumen_kpi_no_presencia(spec, headers, site_id)
    log.info("KPI No Presencia consolidado y cargado")

    log.info("ETL No Presencia — completado")


def _run_precios(df_pt: pd.DataFrame, mes: int, anio: int) -> None:
    log = get_logger("precios")
    separador(log)
    log.info("ETL Precios — iniciando")

    import os
    import requests
    import msal
    from etl_precios import (
        MESES_ESPANOL,
        generar_reporte_captura_precios, generar_analisis_precios,
        generar_resumen_kpi_precios,
    )

    if df_pt.empty:
        log.error(
            "DataFrame de PT vacío para módulo 'precios' "
            "— no hay PDVs con módulo activo. ETL abortado."
        )
        return

    # Autenticación requerida para SharePoint / Azure
    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        raise ValueError("❌ Faltan credenciales de Azure en el archivo .env")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" not in result:
        raise Exception(f"Error de autenticación en Azure: {result.get('error_description')}")

    headers = {"Authorization": f"Bearer {result['access_token']}"}
    
    # Obtener site_id de SharePoint
    url_site = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{paths.SHAREPOINT_SITE_NAME}"
    res_site = requests.get(url_site, headers=headers).json()
    if "id" not in res_site:
        raise Exception(f"No se encontró el sitio SharePoint '{paths.SHAREPOINT_SITE_NAME}'.")
    site_id = res_site["id"]

    spec = pr.resolver(int(mes), int(anio))
    periodo = f"{MESES_ESPANOL[mes]}_{anio}"

    df_pt_unificado = df_pt.copy()
    df_pt_unificado['ID_PDV_INVOLVES'] = df_pt_unificado['ID_PDV_INVOLVES'].astype(str).str.strip()
    df_pt_unificado = df_pt_unificado.groupby('ID_PDV_INVOLVES').first().reset_index()

    ruta_pt_final = os.path.join(str(paths.CIF_PT_DIR), f"Plan_Trabajo_Precios_{periodo}.xlsx")
    
    # Asegurar que el directorio local exista antes de escribir el archivo
    os.makedirs(os.path.dirname(ruta_pt_final), exist_ok=True)
    
    df_pt_unificado.to_excel(ruta_pt_final, index=False)
    log.info(f"PT filtrado persistido: {len(df_pt_unificado):,} PDVs únicos — periodo {periodo}")

    # Delegamos la ejecución pasando headers y site_id a todo el flujo
    generar_reporte_captura_precios(df_pt_unificado, periodo, spec, headers, site_id)
    log.info("Reporte de captura de precios generado")

    generar_analisis_precios(periodo, spec, headers, site_id)
    log.info("Análisis detallado de precios generado")

    generar_resumen_kpi_precios(spec, headers, site_id)
    log.info("KPI precios consolidado")

    log.info("ETL Precios — completado")

def _run_sos(df_pt: pd.DataFrame, mes: int, anio: int) -> None:
    log = get_logger("sos")
    separador(log)
    log.info("ETL SOS — iniciando")

    import etl_sos
    import periodo_resolver as pr
    from pathlib import Path

    if df_pt.empty:
        log.error(
            "DataFrame de PT vacío para módulo 'sos' "
            "— no hay PDVs con módulo activo. ETL abortado."
        )
        return

    spec = pr.resolver(int(mes), int(anio))
    df_pt_unificado = df_pt.copy()
    df_pt_unificado['ID_PDV_INVOLVES'] = df_pt_unificado['ID_PDV_INVOLVES'].astype(str).str.strip()
    df_pt_unificado = df_pt_unificado.groupby('ID_PDV_INVOLVES').first().reset_index()
    
    ruta_salida = Path("pt_consolidado_sos_final.xlsx")
    df_pt_unificado.to_excel(ruta_salida, index=False)
    log.info(f"PT filtrado persistido: {len(df_pt_unificado):,} PDVs únicos")

    # Autenticación requerida por los pasos nativos de etl_sos
    token = etl_sos.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = etl_sos.obtener_site_id(headers)

    # Delegamos la ejecución secuencial pasando los argumentos que exige el módulo
    etl_sos.ejecutar_paso_2_consolidar_encuestas(spec, headers, site_id)
    log.info("Encuestas SOS consolidadas")

    etl_sos.ejecutar_paso_3_cumplimiento_captura(df_pt_unificado, pd.DataFrame(), spec, headers, site_id)
    log.info("Cumplimiento de captura calculado")

    etl_sos.ejecutar_paso_4_normalizar_target_dinamico(spec, headers, site_id)
    log.info("Target normalizado")

    etl_sos.ejecutar_paso_5_cruce_triple_y_calculo(spec, headers, site_id)
    log.info("Cruce triple y cálculo SOS completado")

    etl_sos.generar_resumen_kpi_sos(spec, headers, site_id)
    log.info("KPI SOS consolidado")

    log.info("ETL SOS — completado")


def _run_exhibiciones_pagadas(df_pt: pd.DataFrame = None, mes: int = None, anio: int = None, spec=None, **_) -> None:
    log = get_logger("exhib_pagadas")
    separador(log)
    log.info("ETL Exhibiciones Pagadas — iniciando")

    import etl_exhibiciones_pagadas
    import periodo_resolver as pr

    # Resolver spec de forma flexible dependiendo de cómo se invoque
    if spec is None:
        if mes is not None and anio is not None:
            spec = pr.resolver(int(mes), int(anio))
        else:
            log.error("spec ausente — exhibiciones pagadas no puede correr sin periodo")
            return

    if spec is None:
        log.error("spec ausente — exhibiciones pagadas no puede correr sin periodo")
        return

    # Autenticación requerida para los servicios en la nube
    token = etl_exhibiciones_pagadas.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = etl_exhibiciones_pagadas.obtener_site_id(headers)

    # Ejecución del proceso principal y generación de KPIs
    etl_exhibiciones_pagadas.ejecutar_proceso(spec, headers, site_id)
    etl_exhibiciones_pagadas.generar_resumen_kpi_exhibiciones_pagadas(spec, headers, site_id)
    
    log.info("ETL Exhibiciones Pagadas — completado")


def _run_exhibiciones_gratis(spec=None, **_) -> None:
    log = get_logger("exhib_gratis")
    separador(log)
    log.info("ETL Exhibiciones Gratis — iniciando")
    if spec is None:
        log.error("spec ausente — exhibiciones gratis no puede correr sin periodo")
        return
    etl_exhibiciones_gratis.run(spec)
    etl_exhibiciones_gratis.generar_resumen_kpi_exhibiciones_gratis(spec)
    log.info("ETL Exhibiciones Gratis — completado")


# ── D&P ──────────────────────────────────────────────────────────────────────

def _run_ventas(**_) -> None:
    log = get_logger("ventas")
    separador(log)
    log.info("ETL Ventas D&P — iniciando")
    n = etl_ventas.run()
    log.info(f"ETL Ventas D&P — completado ({n:,} filas en consolidado)")


def _run_impactos(**_) -> None:
    log = get_logger("impactos")
    separador(log)
    log.info("ETL Impactos D&P — iniciando")
    n = etl_impactos.run()
    log.info(f"ETL Impactos D&P — completado ({n:,} filas en consolidado)")


def _run_impactos_segmentos(**_) -> None:
    log = get_logger("impactos_segmentos")
    separador(log)
    log.info("ETL Impactos por Segmento — iniciando")
    etl_impactos_segmentos.run()
    log.info("ETL Impactos por Segmento — completado")


# ─────────────────────────────────────────────────────────────────────────────
# REGISTRO DE ETLs
# ─────────────────────────────────────────────────────────────────────────────
REGISTRO_ETLS = {
    "cif"                : {"fn": _run_cif,                   "clave_pt": "cif"},
    "nopresencia"        : {"fn": _run_nopresencia,            "clave_pt": "no_presencia"},
    "precios"            : {"fn": _run_precios,                "clave_pt": "precios"},
    "sos"                : {"fn": _run_sos,                    "clave_pt": "sos"},
    "exhib_pagadas"      : {"fn": _run_exhibiciones_pagadas,   "clave_pt": None},
    "exhib_gratis"       : {"fn": _run_exhibiciones_gratis,    "clave_pt": None},
    "ventas"             : {"fn": _run_ventas,                 "clave_pt": None},
    "impactos"           : {"fn": _run_impactos,               "clave_pt": None},
    "impactos_segmentos" : {"fn": _run_impactos_segmentos,     "clave_pt": None},
}


# ─────────────────────────────────────────────────────────────────────────────
# CAPTURA DE stdout/stderr DE LOS ETLs (sus prints internos van al log)
# ─────────────────────────────────────────────────────────────────────────────

class _TeeLogger(io.TextIOBase):
    def __init__(self, nombre_etl: str, stream_original):
        self._log     = get_logger(nombre_etl)
        self._orig    = stream_original
        self._buffer  = ""

    def write(self, texto: str) -> int:
        self._orig.write(texto)
        self._buffer += texto
        while "\n" in self._buffer:
            linea, self._buffer = self._buffer.split("\n", 1)
            linea = linea.strip()
            if not linea:
                continue
            if any(k in linea for k in ["❌", "ERROR", "Error"]):
                self._log.error(f"[stdout] {linea}")
            elif any(k in linea for k in ["⚠", "AUDITORÍA", "WARNING"]):
                self._log.warning(f"[stdout] {linea}")
            elif any(k in linea for k in ["✅", "ÉXITO", "EXITO", "✓"]):
                self._log.info(f"[stdout] {linea}")
            else:
                self._log.debug(f"[stdout] {linea}")
        return len(texto)

    def flush(self):
        self._orig.flush()


# ─────────────────────────────────────────────────────────────────────────────
# EJECUCIÓN DE UN ETL INDIVIDUAL (corre en su thread)
# ─────────────────────────────────────────────────────────────────────────────

def _ejecutar_etl(
    nombre: str,
    fn,
    df_pt,
    spec: pr.PeriodoSpec,
    run_id: str,
) -> tuple[str, float, str | None]:
    log = get_logger(nombre)
    ts_inicio = datetime.now().isoformat(timespec="seconds")
    t0  = time.perf_counter()
    tb  = None

    input_hash, files_input = rl.calcular_input_hash(nombre, spec)

    stdout_orig = sys.stdout
    sys.stdout  = _TeeLogger(nombre, stdout_orig)

    try:
        if df_pt is not None:
            fn(df_pt, spec.mes, spec.anio)
        else:
            fn(spec=spec)
    except Exception:
        tb = traceback.format_exc()
        log.critical(
            f"Excepción no capturada — ETL abortado\n{tb}",
        )
    finally:
        sys.stdout = stdout_orig

    duracion = time.perf_counter() - t0
    ts_fin = datetime.now().isoformat(timespec="seconds")

    output_hash, files_output = rl.calcular_output_hash(nombre)

    rl.log_run(rl.RunEvent(
        run_id=run_id,
        ts_inicio=ts_inicio,
        ts_fin=ts_fin,
        duracion_s=round(duracion, 2),
        mes=spec.mes,
        anio=spec.anio,
        etiqueta=spec.etiqueta,
        etl=nombre,
        status="OK" if tb is None else "FAIL",
        error_msg=(tb.splitlines()[-1] if tb else None),
        input_hash=input_hash,
        output_hash=output_hash,
        files_input=files_input,
        files_output=files_output,
    ))

    if tb is None:
        log.info(f"Duración total: {duracion:.1f}s")
    else:
        log.error(f"Duración hasta fallo: {duracion:.1f}s")

    return nombre, duracion, tb


# ─────────────────────────────────────────────────────────────────────────────
# REPORTE FINAL
# ─────────────────────────────────────────────────────────────────────────────

def _imprimir_reporte(
    resultados: dict,
    etls_solicitados: list[str],
    t_loader: float,
    t_total: float,
    log,
) -> None:
    separador(log, "═")
    log.info("REPORTE DE EJECUCIÓN")
    separador(log, "═")

    log.info(f"  {'ETL':<22} {'Estado':<14} {'Tiempo':>8}")
    log.info(f"  {'─'*22} {'─'*14} {'─'*8}")

    t_etls_serie = 0.0
    for nombre in etls_solicitados:
        if nombre not in resultados:
            log.info(f"  {nombre:<22} {'⚠ OMITIDO':<14} {'—':>8}")
            continue
        r      = resultados[nombre]
        if r.get("skipped"):
            estado = "↷ SKIPPED"
        elif r["error"] is None:
            estado = "✓ OK"
        else:
            estado = "✗ ERROR"
        t_etls_serie += r["tiempo"]
        log.info(f"  {nombre:<22} {estado:<14} {r['tiempo']:>7.1f}s")

    log.info(f"  {'─'*22} {'─'*14} {'─'*8}")
    log.info(f"  {'shared_loader':<22} {'✓ OK':<14} {t_loader:>7.1f}s")
    log.info(f"  {'TOTAL (paralelo)':<22} {'':<14} {t_total:>7.1f}s")
    log.info(f"  {'Sin paralelismo (est.)':<22} {'':<14} {t_loader + t_etls_serie:>7.1f}s")
    ahorro = t_etls_serie - (t_total - t_loader)
    log.info(f"  {'Ahorro estimado':<22} {'':<14} {ahorro:>7.1f}s")

    eventos = obtener_eventos_acumulados()
    criticos  = [e for e in eventos if e["nivel"] == "CRITICAL"]
    errores   = [e for e in eventos if e["nivel"] == "ERROR"]
    warnings  = [e for e in eventos if e["nivel"] == "WARNING"]

    separador(log, "─")
    log.info(
        f"Eventos: {len(criticos)} CRITICAL  |  "
        f"{len(errores)} ERROR  |  {len(warnings)} WARNING"
    )

    if not eventos:
        log.info("Sin alertas — ejecución limpia ✓")
    else:
        for nivel, grupo in [("CRITICAL", criticos), ("ERROR", errores), ("WARNING", warnings)]:
            if not grupo:
                continue
            separador(log, "·")
            log.info(f"  {nivel} ({len(grupo)})")
            separador(log, "·")
            for e in grupo:
                etl_tag = f"[{e['etl'].replace('etl.', '')}]"
                log.info(f"  {e['ts']}  {etl_tag:<18}  {e['mensaje']}")
                if e.get("exc"):
                    for linea in e["exc"].strip().splitlines()[-3:]:
                        log.info(f"    {linea}")

    separador(log, "═")
    ruta = ruta_log_actual()
    if ruta:
        log.info(f"Log completo guardado en: {ruta}")
    log.info(f"Fin: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    separador(log, "═")


# ─────────────────────────────────────────────────────────────────────────────
# ORQUESTADOR PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def orquestar(
    etls_a_ejecutar: list[str],
    spec: pr.PeriodoSpec,
    run_id: str,
    max_workers: int = 6,
    ruta_log: str | None = None,
    skip_if_cached: bool = False,
) -> None:
    log_sys = get_logger("orquestador")
    t_total_inicio = time.perf_counter()

    separador(log_sys, "═")
    log_sys.info(f"ORQUESTADOR ETLs — Eficacia")
    log_sys.info(f"Inicio: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log_sys.info(f"Run ID: {run_id}")
    log_sys.info(f"Periodo: {spec.etiqueta} ({spec.mes:02d}/{spec.anio})")
    log_sys.info(f"ETLs seleccionados: {etls_a_ejecutar}")
    log_sys.info(f"Workers paralelos: {max_workers}")
    log_sys.info(f"Skip si cached: {skip_if_cached}")
    separador(log_sys, "═")

    t0_loader = time.perf_counter()
    ruta_pt_directo, ruta_pt_ism = _descubrir_pt_con_periodo(
        str(paths.CIF_PT_DIR), periodo=spec,
    )
    try:
        resultado_pt = cargar_plan_de_trabajo(
            ruta_pt_directo, ruta_pt_ism,
            mes=spec.mes, anio=spec.anio,
        )
    except RuntimeError as e:
        log_sys.critical(
            f"shared_loader falló con error crítico: {e} "
            f"— el pipeline no puede continuar"
        )
        _imprimir_reporte({}, etls_a_ejecutar, 0.0, 0.0, log_sys)
        return

    t_loader = time.perf_counter() - t0_loader
    mes  = resultado_pt["periodo_mes"]
    anio = resultado_pt["periodo_anio"]
    log_sys.info(f"shared_loader completado en {t_loader:.1f}s")

    tareas = []
    skipped: list[str] = []
    for nombre in etls_a_ejecutar:
        if nombre not in REGISTRO_ETLS:
            log_sys.warning(f"ETL '{nombre}' no reconocido en el registro — se omite")
            continue
        cfg      = REGISTRO_ETLS[nombre]
        clave_pt = cfg["clave_pt"]
        df_pt    = resultado_pt.get(clave_pt) if clave_pt else None

        if skip_if_cached:
            puede_skip, motivo = rl.should_skip(nombre, spec)
            if puede_skip:
                log_sys.info(f"  ↳ ETL '{nombre}' SKIPPED ({motivo})")
                input_hash, files_in = rl.calcular_input_hash(nombre, spec)
                output_hash, files_out = rl.calcular_output_hash(nombre)
                rl.log_run(rl.RunEvent(
                    run_id=run_id,
                    ts_inicio=datetime.now().isoformat(timespec="seconds"),
                    ts_fin=datetime.now().isoformat(timespec="seconds"),
                    duracion_s=0.0,
                    mes=spec.mes, anio=spec.anio, etiqueta=spec.etiqueta,
                    etl=nombre, status="SKIPPED",
                    error_msg=motivo,
                    input_hash=input_hash, output_hash=output_hash,
                    files_input=files_in, files_output=files_out,
                ))
                skipped.append(nombre)
                continue

        tareas.append((nombre, cfg["fn"], df_pt, spec))

    separador(log_sys, "─")
    log_sys.info(f"FASE 1 — Lanzando {len(tareas)} ETLs en paralelo "
                 f"({len(skipped)} skipped)")
    separador(log_sys, "─")

    resultados: dict = {}
    for nom in skipped:
        resultados[nom] = {"tiempo": 0.0, "error": None, "skipped": True}

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futuros = {
            executor.submit(_ejecutar_etl, nombre, fn, df_pt, spec, run_id): nombre
            for nombre, fn, df_pt, spec in tareas
        }
        for futuro in as_completed(futuros):
            nombre, t_etl, error = futuro.result()
            resultados[nombre] = {"tiempo": t_etl, "error": error, "skipped": False}
            estado = "✓ OK" if error is None else "✗ FALLÓ"
            log_sys.info(f"ETL '{nombre}' terminó — {estado} en {t_etl:.1f}s")

    t_total = time.perf_counter() - t_total_inicio

    _imprimir_reporte(resultados, etls_a_ejecutar, t_loader, t_total, log_sys)


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def _parse_args():
    parser = argparse.ArgumentParser(
        description="Orquestador de ETLs — Eficacia canal droguerías",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    pr.cli_add_periodos_arg(parser)
    parser.add_argument(
        "--solo", nargs="+", metavar="ETL",
        help=(
            "Ejecutar solo los ETLs indicados:\n"
            "  cif, nopresencia, precios, sos, exhib_pagadas, exhib_gratis\n"
            "Ejemplo: --solo cif nopresencia"
        ),
    )
    parser.add_argument(
        "--workers", type=int, default=6, metavar="N",
        help="Número máximo de threads paralelos (default: 6)",
    )
    parser.add_argument(
        "--log-dir", type=str, default=None, metavar="DIR",
        help="Carpeta donde guardar el archivo de log (default: ./logs/)",
    )
    parser.add_argument(
        "--skip-if-cached", action="store_true",
        help="Saltar ETLs cuyos inputs no cambiaron desde el último run OK.",
    )
    parser.add_argument(
        "--gold", action="store_true",
        help="Regenerar la capa GOLD al final de la corrida.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    specs = pr.periodos_de_args(args)

    log_dir  = Path(args.log_dir) if args.log_dir else Path(__file__).parent / "logs"
    ts_nombre = datetime.now().strftime("%Y%m%d_%H%M%S")
    etiqueta_periodos = (
        f"{specs[0].anio}{specs[0].mes:02d}" if len(specs) == 1
        else f"{specs[0].anio}{specs[0].mes:02d}_a_{specs[-1].anio}{specs[-1].mes:02d}"
    )
    ruta_log = str(log_dir / f"ejecucion_{etiqueta_periodos}_{ts_nombre}.log")

    inicializar_logging(ruta_log=ruta_log)

    todos     = list(REGISTRO_ETLS.keys())
    seleccion = args.solo if args.solo else todos

    log_root = get_logger("orquestador")
    if len(specs) > 1:
        log_root.info("═" * 70)
        log_root.info(
            f"MULTI-PERIODO: {len(specs)} periodos a procesar — "
            f"{', '.join(s.etiqueta for s in specs)}"
        )
        log_root.info("═" * 70)

    run_id = rl.gen_run_id()
    log_root.info(f"Run ID compartido para esta corrida: {run_id}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            log_root.info("")
            log_root.info(f"▶ PERIODO {i}/{len(specs)}: {spec.etiqueta}")
        orquestar(
            etls_a_ejecutar=seleccion,
            spec=spec,
            run_id=run_id,
            max_workers=args.workers,
            ruta_log=ruta_log,
            skip_if_cached=args.skip_if_cached,
        )

    if args.gold:
        log_root.info("")
        log_root.info("═" * 70)
        log_root.info("FASE FINAL — Regenerando capa GOLD")
        log_root.info("═" * 70)
        try:
            import generar_gold
            generar_gold.generar_gold(filtro_periodos=specs)
            log_root.info("GOLD regenerado OK")
        except Exception as e:
            log_root.error(f"FALLÓ generación GOLD: {e}")
