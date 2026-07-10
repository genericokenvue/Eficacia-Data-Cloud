import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

import pandas as pd
import numpy as np
import os
import glob
from datetime import datetime, timedelta

# =============================================================================
# 1. CONFIGURACIÓN DE RUTAS Y DICCIONARIOS
# =============================================================================

MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

import paths
import periodo_resolver as pr

_DIR_PT = str(paths.CIF_PT_DIR)
RUTA_ORIGEN_NP  = str(paths.NP_BASES)
RUTA_SALIDA_NP  = str(paths.NP_SALIDA)

os.makedirs(RUTA_SALIDA_NP, exist_ok=True)

RUTA_PT_CONSOLIDADO_FINAL = os.path.join(RUTA_ORIGEN_NP, "Plan_Trabajo_NoPresencia_Completo.xlsx")
CLAVE_ARCHIVO_NP = "Respuestas"
CLAVE_ARCHIVO_PT = "Plan de trabajo"

# =============================================================================
# 2. CONFIGURACIÓN DE COLUMNAS Y DICCIONARIOS
# =============================================================================

from shared_loader import (
    COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR,
    COLUMNAS_SUPERSET            as COLUMNAS_FINALES_PT,
    COLUMNAS_NUMERICAS,
)

# =============================================================================
# 3. FUNCIONES DEL BLOQUE 1 (CONSOLIDACIÓN PT) — ¡SIN FILTRO DISCOUNTER!
# =============================================================================

def leer_y_normalizar_pt(ruta, hoja, fuente):
    if not os.path.exists(ruta):
        print(f"⚠️  Archivo no encontrado: {fuente}")
        return pd.DataFrame()
    try:
        OBS_A_ROL = {
            "REPORTA GESTOR":              "GESTOR",
            "REPORTA SUPERVISOR":          "SUPERVISOR",
            "REPORTA GENERADOR DE DEMANDA":"GENERADOR DE DEMANDA",
        }
        with pd.ExcelFile(ruta) as xls:
            df_val = pd.read_excel(xls, sheet_name="Captura de modulos")
            df_val.columns = df_val.columns.str.strip().str.upper()
            col_filtro = "NO PRESENCIA_FINAL" if fuente == "ISM" else "NO PRESENCIA"
            obs_col = next((c for c in df_val.columns if c.startswith('OBSERVAC')), None)
            if obs_col is None:
                print(f"⚠ {fuente}: hoja Captura de modulos sin OBSERVACIÓN")
                return pd.DataFrame()
            
            df_val_filtrado = df_val[df_val[col_filtro] == 1].copy()
            
            # ========================================================================
            # 💥 SE ELIMINÓ EL FILTRO DE "DISCOUNTER" DE LA LECTURA DEL PLAN DE TRABAJO 💥
            # ========================================================================
            
            pdvs_obs = (
                df_val_filtrado[["ID PDV INVOLVES", obs_col]]
                .rename(columns={obs_col: "OBSERVACION"})
            )
            pdvs_obs["OBSERVACION"] = pdvs_obs["OBSERVACION"].astype(str).str.strip().str.upper()
            pdvs_obs["ROL_ESPERADO"] = pdvs_obs["OBSERVACION"].map(OBS_A_ROL).fillna("")

            df = pd.read_excel(xls, sheet_name=hoja)
            df.columns = df.columns.str.strip().str.upper()
            df = df[df["ID PDV INVOLVES"].isin(pdvs_obs["ID PDV INVOLVES"])].copy()

            df.rename(columns={col: COLUMNAS_ESTANDAR[col] for col in df.columns if col in COLUMNAS_ESTANDAR}, inplace=True)
            df = df.merge(pdvs_obs, left_on="ID_PDV_INVOLVES", right_on="ID PDV INVOLVES", how="left")
            df["ROL"] = df["ROL"].astype(str).str.strip().str.upper()
            n_antes = len(df)
            df = df[df["ROL"] == df["ROL_ESPERADO"]].copy()
            print(f"  {fuente}: filtro OBSERVACIÓN → {n_antes} → {len(df)} filas (responsables reales)")
            df["FUENTE"] = fuente

            for col in COLUMNAS_NUMERICAS:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce')

            if "FECHA" in df.columns:
                df["FECHA"] = pd.to_datetime(df["FECHA"], errors='coerce').dt.strftime("%d/%m/%Y")

            for col in COLUMNAS_FINALES_PT:
                if col not in df.columns: df[col] = ""

            df = df[COLUMNAS_FINALES_PT].copy()
            for col in df.columns:
                if col not in COLUMNAS_NUMERICAS and col != "FECHA":
                    df[col] = df[col].astype(str).str.strip().replace(['nan', 'NaT', 'None'], '')

            print(f"✅ [{fuente}] Procesados completos: {len(df)} registros.")
            return df
    except PermissionError as e:
        raise PermissionError(f"No se puede leer {fuente}: {ruta}. Cierra el archivo.") from e
    except Exception as e:
        print(f"❌ Error en {fuente}: {e}")
        raise

# =============================================================================
# 4. FUNCIONES DEL BLOQUE 2 (ENCUESTAS Y MATRIZ EN BASE A 4 ARCHIVOS)
# =============================================================================

def procesar_no_presencia(periodo_nom, spec: pr.PeriodoSpec):
    print("\n--- INICIANDO: Bloque No Presencia (Encuestas Básicas) ---")
    archivo_periodo = pr.np_respuestas(spec)
    print(f"Archivo encuestas principales: {archivo_periodo.name}")

    df = pd.read_excel(archivo_periodo)
    df['Fecha de la encuesta'] = pd.to_datetime(df['Fecha de la encuesta'], dayfirst=True)
    
    fecha_minima = df['Fecha de la encuesta'].min()
    lunes_inicio = fecha_minima - timedelta(days=fecha_minima.weekday())
    semana_base = lunes_inicio.isocalendar()[1]

    def calcular_semana_dinamica(fecha):
        try: return f"SEMANA {(fecha.isocalendar()[1] - semana_base) + 1}"
        except: return "N/A"

    df['SEMANA'] = df['Fecha de la encuesta'].apply(calcular_semana_dinamica)
    df['Fecha de la encuesta'] = df['Fecha de la encuesta'].dt.date

    columnas_requeridas = [
        "ID de la encuesta", "Marca", "Categoría de producto", "Supercategoría", 
        "Línea de producto", "Producto (SKU)", "Código de barras", "ID del PDV", 
        "PDV", "Fecha de la encuesta", "Empleado", "Perfil de acceso", 
        "¿EL PRODUCTO ESTÁ AGOTADO?", "¿CUÁL ES LA CAUSAL DE QUE EL PRODUCTO ESTÁ AGOTADO?", 
        "SEMANA"
    ]
    df_resultado = df[[c for c in columnas_requeridas if c in df.columns]].copy()
    
    ruta_final = os.path.join(RUTA_SALIDA_NP, f"NO_PRESENCIA_PROCESADO_{periodo_nom}.xlsx")
    df_resultado.to_excel(ruta_final, index=False)
    print(f"EXITO: Encuestas creadas en {ruta_final}")
    return ruta_final

def procesar_plan_de_trabajo_semanas(ruta_consolidado, periodo_nom):
    print("\n--- INICIANDO: Bloque Plan de Trabajo (Semanas) ---")
    df = pd.read_excel(ruta_consolidado)
    df['FECHA'] = pd.to_datetime(df['FECHA'], dayfirst=True)
    df = df.dropna(subset=['FECHA']).copy()
    
    fecha_minima = df['FECHA'].min()
    lunes_inicio = fecha_minima - timedelta(days=fecha_minima.weekday())
    semana_base = lunes_inicio.isocalendar()[1]

    df['SEMANA'] = df['FECHA'].apply(lambda x: f"SEMANA {(x.isocalendar()[1] - semana_base) + 1}")
    df['FECHA'] = df['FECHA'].dt.date

    ruta_final = os.path.join(RUTA_SALIDA_NP, f"PLAN_TRABAJO_PROCESADO_{periodo_nom}.xlsx")
    df.to_excel(ruta_final, index=False)
    print(f"EXITO: Plan Procesado creado en {ruta_final}")
    return ruta_final

def generar_matriz_seguimiento(ruta_np, ruta_pt, periodo_nom, spec=None):
    print("\n--- GENERANDO MATRIZ COMPARATIVA FINAL (4 FUENTES DE CAPTURAS CONSOLIDADAS) ---")
    
    # Archivo 1: Encuestas Base No Presencia
    df_np = pd.read_excel(ruta_np)
    lista_capturas = [df_np[['ID del PDV']]]
    
    # Archivos 2, 3 y 4: Discounters dinámicos (ARA, D1, OXXO)
    for marca in ['ARA', 'D1', 'OXXO']:
        patron_busqueda = os.path.join(RUTA_ORIGEN_NP, f"*{marca}*{periodo_nom}*.xlsx")
        archivos_encontrados = glob.glob(patron_busqueda)
        
        if archivos_encontrados:
            ruta_adicional = archivos_encontrados[0]
            print(f"📥 Archivo dinámico detectado e integrando: {os.path.basename(ruta_adicional)}")
            try:
                df_extra = pd.read_excel(ruta_adicional)
                df_extra.columns = df_extra.columns.str.strip().str.upper()
                
                if 'ID DEL PDV' in df_extra.columns:
                    df_extra = df_extra.rename(columns={'ID DEL PDV': 'ID del PDV'})
                    lista_capturas.append(df_extra[['ID del PDV']])
                    print(f"   ✅ Se cargaron {len(df_extra)} filas de capturas de {marca}")
                else:
                    print(f"⚠️ Alerta: El archivo de {marca} no tiene la columna 'ID del PDV'.")
            except Exception as e:
                print(f"❌ No se pudo procesar el archivo adicional de {marca}: {e}")
        else:
            print(f"ℹ️ No se detectó archivo opcional para {marca} en el periodo {periodo_nom}.")

    # Consolidación total de las 4 fuentes de capturas
    df_capturas_totales = pd.concat(lista_capturas, ignore_index=True)

    # Leer del plan COMPLETO (Ya viene unificado desde el main sin filtros de exclusión)
    df_pt = pd.read_excel(RUTA_PT_CONSOLIDADO_FINAL)

    # MES/AÑO desde el spec si está disponible
    if spec is not None:
        df_pt['MES'] = spec.mes
        df_pt['AÑO'] = spec.anio
    else:
        fechas_dt = pd.to_datetime(df_pt['FECHA'], dayfirst=True, errors='coerce')
        df_pt['MES'] = fechas_dt.dt.month.fillna(0).astype(int)
        df_pt['AÑO'] = fechas_dt.dt.year.fillna(0).astype(int)

    # Normalizar IDs de ambos lados antes del merge con shared_loader
    from shared_loader import id_a_str
    df_pt['ID_PDV_INVOLVES'] = id_a_str(df_pt['ID_PDV_INVOLVES'])
    df_pt['NOMBRE']          = df_pt['NOMBRE'].astype(str).str.strip().str.upper()
    
    df_capturas_totales['ID del PDV'] = id_a_str(df_capturas_totales['ID del PDV'])

    # Plan unificado agrupado (Colapsando semanas)
    plan_agg = (
        df_pt.assign(_PLAN=1)
             .groupby(
                 ['ID_PDV_INVOLVES', 'NOMBRE_PDV', 'NOMBRE',
                  'ROL', 'SUPERVISOR_LIDER', 'MES', 'AÑO'],
                 dropna=False, as_index=False,
             )
             .agg(PLANEADO_MES=('_PLAN', 'max'))
    )

    # Capturas totales agrupadas por PDV unificado
    cap_agg = (
        df_capturas_totales.assign(_CAP=1)
             .groupby(['ID del PDV'], dropna=False, as_index=False)
             .agg(CAPTURA_MES=('_CAP', 'max'))
             .rename(columns={'ID del PDV': 'ID_PDV_INVOLVES'})
    )

    # Cruce completo (Left merge)
    matriz = plan_agg.merge(cap_agg, on=['ID_PDV_INVOLVES'], how='left')
    matriz['CAPTURA_MES'] = matriz['CAPTURA_MES'].fillna(0).astype(int)
    matriz['PLANEADO_MES'] = matriz['PLANEADO_MES'].fillna(0).astype(int)
    matriz['%_CUMPLIMIENTO_MES'] = np.where(
        matriz['PLANEADO_MES'] > 0,
        matriz['CAPTURA_MES'] / matriz['PLANEADO_MES'],
        0,
    )

    columnas_finales_reporte = [
        'ID_PDV_INVOLVES', 'NOMBRE_PDV', 'NOMBRE', 'ROL', 'SUPERVISOR_LIDER',
        'PLANEADO_MES', 'CAPTURA_MES', '%_CUMPLIMIENTO_MES', 'MES', 'AÑO'
    ]
    matriz = matriz[columnas_finales_reporte]

    ruta_matriz = os.path.join(RUTA_SALIDA_NP, f"REPORTE_NO_PRESENCIA_{periodo_nom}.xlsx")
    matriz.to_excel(ruta_matriz, index=False)
    print(f"🏁 FINALIZADO: Reporte generado exitosamente en {ruta_matriz}")

# =============================================================================
# 5. FUNCION AGOTADOS Y RESTO DEL SCRIPT IGUAL...
# =============================================================================

def generar_analisis_agotados(ruta_np, periodo_nom):
    print("\n--- GENERANDO ANÁLISIS DE AGOTADOS (INVOLVES) ---")
    df = pd.read_excel(ruta_np)
    df['Fecha de la encuesta'] = pd.to_datetime(df['Fecha de la encuesta'])
    df['MES'] = df['Fecha de la encuesta'].dt.month
    df['AÑO'] = df['Fecha de la encuesta'].dt.year
    df['ES_AGOTADO'] = df['¿EL PRODUCTO ESTÁ AGOTADO?'].apply(lambda x: 1 if str(x).strip().upper() == 'SI' else 0)
    df['PRODUCTOS_MEDIDOS'] = 1

    analisis = df.groupby(['MES', 'AÑO', 'ID del PDV', 'PDV', 'Marca']).agg({
        'PRODUCTOS_MEDIDOS': 'sum',
        'ES_AGOTADO': 'sum'
    }).reset_index()

    analisis.rename(columns={'ES_AGOTADO': 'CANT_AGOTADOS'}, inplace=True)
    analisis['%_AGOTADOS'] = (analisis['CANT_AGOTADOS'] / analisis['PRODUCTOS_MEDIDOS']).fillna(0)

    cols = [c for c in analisis.columns if c not in ['MES', 'AÑO']] + ['MES', 'AÑO']
    analisis = analisis[cols]

    ruta_agotados = os.path.join(RUTA_SALIDA_NP, f"ANALISIS_AGOTADOS_{periodo_nom}.xlsx")
    analisis.to_excel(ruta_agotados, index=False)
    print(f"✅ ÉXITO: Reporte de Agotados generado en {ruta_agotados}")

def generar_resumen_kpi_no_presencia(spec: pr.PeriodoSpec):
    print(f"\n--- PASO KPI (V3): RESUMEN NP por gestor  ({spec.etiqueta}) ---")
    import paths as _paths
    nombre_esperado = f"REPORTE_NO_PRESENCIA_{spec.mes_str_upper}_{spec.anio}.xlsx"
    ruta_np = _paths.NP_SALIDA / nombre_esperado
    if not ruta_np.is_file():
        print(f"❌ No existe el reporte del periodo: {ruta_np}")
        return
    df = pd.read_excel(ruta_np, engine='openpyxl')
    print(f"  Leyendo: {ruta_np.name} ({len(df)} filas)")

    if 'MES' in df.columns and 'AÑO' in df.columns:
        periodos = set(zip(
            pd.to_numeric(df['MES'], errors='coerce').dropna().astype(int),
            pd.to_numeric(df['AÑO'], errors='coerce').dropna().astype(int),
        ))
        if periodos != {(spec.mes, spec.anio)}:
            raise ValueError(f"NP KPI: el reporte tiene periodos {periodos} pero se solicitó {(spec.mes, spec.anio)}.")
    else:
        df['MES'] = spec.mes
        df['AÑO'] = spec.anio

    from shared_loader import calcular_kpi_simple_y_escribir
    n = calcular_kpi_simple_y_escribir(
        df_origen=df,
        col_planeado="PLANEADO_MES",
        col_ejecutado="CAPTURA_MES",
        ruta_kpi_mes=str(_paths.NP_OUT_KPIS),
        ruta_kpi_historico=str(_paths.NP_OUT_KPIS_HISTORICO),
        nombre_cumplimiento="EJECUCION",
        nombres_renombre=("SUMA PLANEADO", "SUMA CAPTURADO"),
    )
    print(f"  ✅ Mes activo: {n} gestores → {_paths.NP_OUT_KPIS.name}")
    print(f"  ✅ Histórico acumulado → {_paths.NP_OUT_KPIS_HISTORICO.name}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL NO PRESENCIA — Eficacia (Sprint 16.1: multi-periodo)")
    parser.add_argument("--solo", nargs="+", choices=["full", "kpi"],
                        help="Por default: corre todo + kpi. Usa --solo kpi para solo KPI.")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    pasos = args.solo or ["full", "kpi"]

    specs = pr.periodos_de_args(args)

    for i, spec in enumerate(specs, 1):
        print("=" * 50)
        print(f" INICIANDO ETL UNIFICADO NO PRESENCIA  ({spec.etiqueta})")
        print("=" * 50)

        try:
            if "full" in pasos:
                from shared_loader import descubrir_archivos_pt
                ruta_directo, ruta_ism = descubrir_archivos_pt(_DIR_PT, periodo=spec)
                if not ruta_directo or not ruta_ism:
                    raise FileNotFoundError(f"NP ({spec.etiqueta}): faltan PT del periodo.")
                
                df_dir = leer_y_normalizar_pt(ruta_directo, "Plan de trabajo", "DIRECTO")
                df_ism = leer_y_normalizar_pt(ruta_ism, "Plan de trabajo CIF", "ISM")

                if not df_dir.empty or not df_ism.empty:
                    df_total_pt = pd.concat([df_dir, df_ism], ignore_index=True)
                    df_total_pt.to_excel(RUTA_PT_CONSOLIDADO_FINAL, index=False)
                    
                    PERIODO_DATA = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
                    ruta_np_proc = procesar_no_presencia(PERIODO_DATA, spec)
                    ruta_pt_proc = procesar_plan_de_trabajo_semanas(RUTA_PT_CONSOLIDADO_FINAL, PERIODO_DATA)

                    if ruta_np_proc and ruta_pt_proc:
                        generar_matriz_seguimiento(ruta_np_proc, ruta_pt_proc, PERIODO_DATA, spec=spec)
                        generar_analisis_agotados(ruta_np_proc, PERIODO_DATA)
            
            if "kpi" in pasos:
                generar_resumen_kpi_no_presencia(spec)

        except Exception as e:
            print(f"\n❌ ERROR en {spec.etiqueta}: {e}")
            raise