import os
import glob
import pandas as pd
import numpy as np

import paths
import periodo_resolver as pr

# ==============================================================================
# 1. CONFIGURACIÓN DE RUTAS Y PALABRAS CLAVE
# ==============================================================================
BASE_DIR   = str(paths.EXHIB_DATA_DIR)   # auto: subcarpeta del mes (prod) o BASES/EXHIBICIONES (dev)
OUTPUT_DIR = str(paths.EXHIB_SALIDA)

KEYS_PLANNING = ["planning"]
KEYS_INVOLVES = ["base", "exhibiciones", "planning"]
KEYS_MAESTRO  = ["plan", "de", "trabajo"]

BASE_NAME_DETALLE  = "Resultado_exhibiciones_pagadas"
BASE_NAME_AGRUPADO = "Rexhibiciones_pagadas_agrupado"

# ==============================================================================
# 2. DEFINICIÓN DE COLUMNAS
# ==============================================================================
COL_ID_INVOLVES  = "ID_PDV_INVOLVES"
COL_PDV_PLANNING = "*PUNTO DE VENTA"
COL_TIPO_PLAN    = "*TIPO - PAGADAS"
COL_MARCA_PLAN   = "*MARCA - PAGADAS"

COL_CANT_PAGADA  = "*PAGADAS - *DIGITE EL NUMERO DE EXHIBICIONES ADICIONALES PARA ESTE TIPO."
COL_CANT_CONTRA  = "*Digite el numero de exhibiciones adicionales para este tipo. - CONTRAPRESTACIÓN"

COL_FECHA_ENCUESTA = "Fecha de la encuesta" 
COL_IMP_OK      = "La Exhibicion esta implementada de acuerdo con el planning?"
COL_CAUSAL      = "Indique las causales:"
COL_TIPO_CONTRA = "Seleccionar el Tipo de la exhibicion - CONTRAPRESTACIÓN"
COL_MAR_CONTRA  = "MARCA - CONTRAPRESTACIÓN"

COL_CANT_PLANEADA = "CANTIDAD_PLANEADA"
COL_CANT_EJECUTADA = "CANTIDAD_EJECUTADA"

# ==============================================================================
# 3. FUNCIONES DE APOYO
# ==============================================================================

import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass


def buscar_archivo_por_keys(lista_keys, ignorar=None, preferir=None):
    """
    Busca el primer .xlsx en BASE_DIR cuyo nombre contiene TODAS las keys
    y NINGUNA de las ignoradas. Si `preferir` está dado, devuelve primero
    cualquier match que también contenga alguna de esas substrings.
    Esto desambigua casos donde aparecen varios candidatos (ej. PT Directo
    vs PT ISM).
    """
    if ignorar is None:
        ignorar = []
    archivos = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
    matches = [
        f for f in archivos
        if all(k in os.path.basename(f).lower() for k in lista_keys)
        and not any(i in os.path.basename(f).lower() for i in ignorar)
    ]
    if not matches:
        return None
    if preferir:
        for f in matches:
            nombre = os.path.basename(f).lower()
            if any(p in nombre for p in preferir):
                return f
    return matches[0]

def obtener_periodo_dinamico(df):
    try:
        f = pd.to_datetime(df[COL_FECHA_ENCUESTA]).dropna().iloc[0]
        meses = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio",
                 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
        return f"{meses[f.month]} {f.year}"
    except:
        return "Periodo_Desconocido"

def limpiar_texto_llave(serie):
    return (serie.astype(str).str.replace(r'\.0$', '', regex=True)
            .str.strip().str.upper()
            .str.replace('Á','A').str.replace('É','E').str.replace('Í','I')
            .str.replace('Ó','O').str.replace('Ú','U')
            .replace('NAN', ''))

# ==============================================================================
# 4. LÓGICA DE NEGOCIO (Planeado vs Ejecutado)
# ==============================================================================

def preparar_datos_agrupado(df):
    # 1. Cantidad Planeada
    df[COL_CANT_PLANEADA] = pd.to_numeric(df[COL_CANT_PAGADA], errors='coerce').fillna(0)
    
    # 2. ASIGNACIÓN DE "NO" A LOS QUE NO CRUZAN
    # Si la columna de implementación está vacía tras el merge, ponemos "No"
    df[COL_IMP_OK] = df[COL_IMP_OK].fillna("No")
    
    # 3. Inicializar campos finales
    df['MARCA_FINAL'] = df[COL_MARCA_PLAN]
    df['TIPO_FINAL']  = df[COL_TIPO_PLAN]
    
    # Por defecto, ejecutada = planeada (solo si cruzó y no es contraprestación)
    df[COL_CANT_EJECUTADA] = df[COL_CANT_PLANEADA]

    # 4. Caso Contraprestación
    causal_std = limpiar_texto_llave(df[COL_CAUSAL])
    es_contra = causal_std.str.contains("CONTRAPRESTACION A LA NEGOCIACION", na=False)

    df.loc[es_contra, 'MARCA_FINAL'] = df.loc[es_contra, COL_MAR_CONTRA]
    df.loc[es_contra, 'TIPO_FINAL']  = df.loc[es_contra, COL_TIPO_CONTRA]
    df.loc[es_contra, COL_CANT_EJECUTADA] = pd.to_numeric(df.loc[es_contra, COL_CANT_CONTRA], errors='coerce').fillna(0)
    df.loc[es_contra, COL_IMP_OK] = "Si"

    # 5. REGLA DE ORO: Si Implementación es No (incluyendo los que no cruzaron), Ejecutado es 0
    imp_final = df[COL_IMP_OK].astype(str).str.strip().str.upper()
    df.loc[imp_final == "NO", COL_CANT_EJECUTADA] = 0
    
    return df

# ==============================================================================
# 5. ORQUESTADOR
# ==============================================================================

def ejecutar_proceso(spec: pr.PeriodoSpec):
    print(f"🚀 Iniciando auditoría exhibiciones pagadas  ({spec.etiqueta})...")

    if not BASE_DIR or not os.path.isdir(BASE_DIR):
        print(f"❌ Error: No se encontró el directorio de exhibiciones: {BASE_DIR}"); return

    # Sprint 15.5.7 — Fix F5: resolución por periodo, no glob ni mtime.
    # 'planning' real (asterisco): PLANNING DE <MES> <AÑO>.xlsx
    # 'involves' (encuestas):       Base Exhibiciones Planning <Mes> <Año>.xlsx
    ruta_p = str(pr.exh_planning_master(spec))
    ruta_i = str(pr.exh_base_planning(spec))

    # PT del periodo: vía periodo_resolver (en CIF/PLAN DE TRABAJO).
    ruta_m = str(pr.cif_pt_directo(spec))

    # Cargas — el PT trae las columnas con espacios ("Nombre Punto de Venta",
    # "ID PDV INVOLVES"); las normalizamos con el rename canónico.
    from shared_loader import COLUMNAS_ESTANDAR_UNIFICADO
    df_m = pd.read_excel(ruta_m, sheet_name="Plan de trabajo")
    df_m.columns = df_m.columns.str.strip()
    df_m.rename(columns={
        c: COLUMNAS_ESTANDAR_UNIFICADO[c]
        for c in df_m.columns if c in COLUMNAS_ESTANDAR_UNIFICADO
    }, inplace=True)
    maestro = df_m[["NOMBRE_PDV", COL_ID_INVOLVES]].drop_duplicates()

    df_p = pd.read_excel(ruta_p, sheet_name=0)
    df_p.columns = [str(c).strip().upper() for c in df_p.columns]
    df_p["_JOIN_"] = df_p[COL_PDV_PLANNING].astype(str).str.strip().str.upper()
    df_detalle = df_p.merge(maestro, left_on="_JOIN_", right_on="NOMBRE_PDV", how="left")
    
    df_inv = pd.read_excel(ruta_i)
    serie_marca_inv = df_inv.iloc[:, 61] # Columna BJ
    df_inv.columns = [str(c).strip() for c in df_inv.columns]
    # Sprint 15.5.7: el periodo es el solicitado, no detectado del archivo.
    periodo = f"{spec.mes_str} {spec.anio}"

    # Llaves
    df_detalle["_LL_"] = (limpiar_texto_llave(df_detalle[COL_ID_INVOLVES]) + "_" + 
                          limpiar_texto_llave(df_detalle[COL_TIPO_PLAN]) + "_" + 
                          limpiar_texto_llave(df_detalle[COL_MARCA_PLAN]))
    
    col_id_inv = "ID del PDV" if "ID del PDV" in df_inv.columns else "PDV"
    df_inv["_LL_"] = (limpiar_texto_llave(df_inv[col_id_inv]) + "_" + 
                      limpiar_texto_llave(df_inv["Seleccionar el Tipo de la exhibicion (Pagadas)"]) + "_" + 
                      limpiar_texto_llave(serie_marca_inv))
    
    # Cruce
    campos_auditoria = [COL_IMP_OK, COL_CAUSAL, COL_TIPO_CONTRA, COL_MAR_CONTRA, COL_CANT_CONTRA]
    df_inv_sub = df_inv[["_LL_"] + [c for c in campos_auditoria if c in df_inv.columns]].drop_duplicates(subset=["_LL_"])
    
    # Unimos y aplicamos lógica
    df_final = df_detalle.merge(df_inv_sub, on="_LL_", how="left")
    df_final = preparar_datos_agrupado(df_final)

    # Agrupado
    dim_agrupar = [COL_ID_INVOLVES, 'TIPO_FINAL', 'MARCA_FINAL', COL_IMP_OK, COL_CAUSAL]
    df_agrupado = df_final.groupby(dim_agrupar, dropna=False)[[COL_CANT_PLANEADA, COL_CANT_EJECUTADA]].sum().reset_index()

    # Guardado
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    
    df_final.drop(columns=["_LL_", "_JOIN_", "NOMBRE_PDV"], errors='ignore').to_excel(
        os.path.join(OUTPUT_DIR, f"{BASE_NAME_DETALLE} {periodo}.xlsx"), index=False)
    
    df_agrupado.to_excel(os.path.join(OUTPUT_DIR, f"{BASE_NAME_AGRUPADO} {periodo}.xlsx"), index=False)

    print(f"✅ ¡Hecho! Los registros que no cruzaron ahora tienen 'No' y ejecución en 0.")

def _calcular_captura_planning(spec: pr.PeriodoSpec) -> "pd.DataFrame":
    """
    Sprint 17 — cumplimiento de captura del módulo PLANNING por gestor.

    Cruza:
      • PLANNING master  (BASES/EXHIBICIONES/PLANNING DE {Mes} {Año}.xlsx)
        — qué (PDV, Empleado) tenía que capturar.
      • Encuestas         (BASES/EXHIBICIONES/Base Exhibiciones Planning ...)
        — qué PDVs respondieron el formulario.

    Devuelve un DataFrame con columnas:
        MES, AÑO, EMPLEADO, CAPTURA_PLANEADA, CAPTURA_EJECUTADA, CUMPLIMIENTO_CAPTURA
    Granularidad: 1 fila por gestor con sus PDVs únicos del planning.
    """
    import paths as _paths
    try:
        ruta_planning  = pr.exh_planning_master(spec)
        ruta_encuestas = pr.exh_base_planning(spec)
    except FileNotFoundError as e:
        print(f"  ⚠️  Captura planning: {e} — KPI capturada omitido")
        return pd.DataFrame(columns=['MES','AÑO','EMPLEADO',
                                     'CAPTURA_PLANEADA','CAPTURA_EJECUTADA',
                                     'CUMPLIMIENTO_CAPTURA'])

    df_plan = pd.read_excel(ruta_planning, engine='openpyxl')
    df_enc  = pd.read_excel(ruta_encuestas, engine='openpyxl')

    col_pdv_p = '*Punto de venta' if '*Punto de venta' in df_plan.columns else (
                '*PUNTO DE VENTA' if '*PUNTO DE VENTA' in df_plan.columns else None)
    col_emp_p = '*Empleado'       if '*Empleado'       in df_plan.columns else (
                '*EMPLEADO'       if '*EMPLEADO'       in df_plan.columns else None)
    if not (col_pdv_p and col_emp_p):
        print("  ❌ Captura planning: PLANNING master sin columnas '*Punto de venta'/'*Empleado'")
        return pd.DataFrame(columns=['MES','AÑO','EMPLEADO',
                                     'CAPTURA_PLANEADA','CAPTURA_EJECUTADA',
                                     'CUMPLIMIENTO_CAPTURA'])
    col_pdv_e = 'PDV' if 'PDV' in df_enc.columns else None
    if not col_pdv_e:
        print("  ❌ Captura planning: encuesta sin columna 'PDV'")
        return pd.DataFrame(columns=['MES','AÑO','EMPLEADO',
                                     'CAPTURA_PLANEADA','CAPTURA_EJECUTADA',
                                     'CUMPLIMIENTO_CAPTURA'])

    def _norm(s): return s.astype(str).str.strip().str.upper()
    df_plan['_PDV_K']  = _norm(df_plan[col_pdv_p])
    df_plan['EMPLEADO'] = _norm(df_plan[col_emp_p])
    df_enc['_PDV_K']   = _norm(df_enc[col_pdv_e])

    pdvs_capturados = set(df_enc['_PDV_K'].dropna().unique())

    # 1 fila por (gestor, PDV) del planning
    asign = df_plan[['EMPLEADO', '_PDV_K']].drop_duplicates()
    asign['CAPTURA_PLANEADA']  = 1
    asign['CAPTURA_EJECUTADA'] = asign['_PDV_K'].isin(pdvs_capturados).astype(int)

    agg = (asign.groupby('EMPLEADO', as_index=False)
                .agg(CAPTURA_PLANEADA=('CAPTURA_PLANEADA', 'sum'),
                     CAPTURA_EJECUTADA=('CAPTURA_EJECUTADA', 'sum')))
    agg['CUMPLIMIENTO_CAPTURA'] = (
        agg['CAPTURA_EJECUTADA'] / agg['CAPTURA_PLANEADA']
    ).where(agg['CAPTURA_PLANEADA'] != 0, 0)
    agg['MES'] = spec.mes
    agg['AÑO'] = spec.anio
    return agg[['MES','AÑO','EMPLEADO',
                'CAPTURA_PLANEADA','CAPTURA_EJECUTADA','CUMPLIMIENTO_CAPTURA']]


def generar_resumen_kpi_exhibiciones_pagadas(spec: pr.PeriodoSpec):
    """Resumen V3: cumplió/no cumplió por empleado, agrupado por mes.

    Sprint 15.5.7: lee el archivo del periodo solicitado (no glob).
    Sprint 17: agrega CAPTURA_PLANEADA / CAPTURA_EJECUTADA / CUMPLIMIENTO_CAPTURA.
    """
    print(f"\n--- KPI (V3): RESUMEN EXHIBICIONES PAGADAS por empleado  ({spec.etiqueta}) ---")
    import paths as _paths
    import sys as _sys
    _SCR = str(__import__('pathlib').Path(__file__).resolve().parent)
    if _SCR not in _sys.path:
        _sys.path.insert(0, _SCR)

    nombre = f"{BASE_NAME_DETALLE} {spec.mes_str} {spec.anio}.xlsx"
    ruta = _paths.EXHIB_SALIDA / nombre
    if not ruta.is_file():
        print(f"❌ No existe: {ruta} — corre primero ejecutar_proceso(spec)")
        return
    df = pd.read_excel(ruta, engine='openpyxl')
    print(f"  Leyendo: {ruta.name} ({len(df)} filas)")

    # Normalización
    df['CANTIDAD_PLANEADA']  = pd.to_numeric(df.get('CANTIDAD_PLANEADA'),  errors='coerce').fillna(0)
    df['CANTIDAD_EJECUTADA'] = pd.to_numeric(df.get('CANTIDAD_EJECUTADA'), errors='coerce').fillna(0)
    df['PLANEADO_REC']  = (df['CANTIDAD_PLANEADA']  > 0).astype(int)
    df['EJECUTADO_REC'] = (df['CANTIDAD_EJECUTADA'] > 0).astype(int)

    # Empleado: la columna se llama '*EMPLEADO' en el reporte detallado
    col_empleado = '*EMPLEADO' if '*EMPLEADO' in df.columns else 'EMPLEADO'
    if col_empleado not in df.columns:
        print(f"❌ Falta la columna {col_empleado}")
        return

    # Sprint 15.5.7: el periodo es el solicitado. No detección por fecha.
    df['MES'] = spec.mes
    df['AÑO'] = spec.anio

    # Agrupar por (MES, AÑO, EMPLEADO)
    resumen = (
        df.groupby(['MES', 'AÑO', col_empleado], dropna=False)
          .agg(PLANEADO=('PLANEADO_REC',  'sum'),
               EJECUTADO=('EJECUTADO_REC', 'sum'))
          .reset_index()
    )
    resumen['CUMPLIMIENTO'] = (
        resumen['EJECUTADO'] / resumen['PLANEADO']
    ).where(resumen['PLANEADO'] != 0, 0)
    resumen = resumen.rename(columns={col_empleado: 'EMPLEADO'})
    resumen['EMPLEADO'] = resumen['EMPLEADO'].astype(str).str.strip().str.upper()

    # Sprint 17 — agregar cumplimiento de captura del PLANNING.
    cap = _calcular_captura_planning(spec)
    if not cap.empty:
        resumen = resumen.merge(cap, on=['MES','AÑO','EMPLEADO'], how='outer')
        # Si quedó EMPLEADO sin ejecución (capturó pero sin exhibición planeada
        # en el detalle), llenar las cols de ejecución con 0.
        for c in ('PLANEADO', 'EJECUTADO', 'CUMPLIMIENTO'):
            if c in resumen.columns:
                resumen[c] = resumen[c].fillna(0)
        for c in ('CAPTURA_PLANEADA', 'CAPTURA_EJECUTADA', 'CUMPLIMIENTO_CAPTURA'):
            if c in resumen.columns:
                resumen[c] = resumen[c].fillna(0)
        resumen['MES'] = resumen['MES'].fillna(spec.mes).astype(int)
        resumen['AÑO'] = resumen['AÑO'].fillna(spec.anio).astype(int)

    cols_orden = ['MES','AÑO','EMPLEADO',
                  'PLANEADO','EJECUTADO','CUMPLIMIENTO',
                  'CAPTURA_PLANEADA','CAPTURA_EJECUTADA','CUMPLIMIENTO_CAPTURA']
    cols_orden = [c for c in cols_orden if c in resumen.columns]
    resumen = resumen[cols_orden].sort_values(['AÑO', 'MES', 'EMPLEADO']).reset_index(drop=True)

    # Mes activo
    if not resumen.empty:
        anio_max = int(resumen['AÑO'].max())
        mes_max  = int(resumen[resumen['AÑO'] == anio_max]['MES'].max())
        mask = (resumen['AÑO'] == anio_max) & (resumen['MES'] == mes_max)
        resumen_activo = resumen[mask].copy()
    else:
        resumen_activo = resumen.copy()

    os.makedirs(os.path.dirname(str(_paths.EXHIB_PAG_OUT_KPIS)), exist_ok=True)
    resumen_activo.to_excel(str(_paths.EXHIB_PAG_OUT_KPIS), index=False, engine='openpyxl')
    print(f"  ✅ Mes activo: {len(resumen_activo)} empleados → {_paths.EXHIB_PAG_OUT_KPIS.name}")

    # Histórico con upsert por (MES, AÑO, EMPLEADO)
    ruta_hist = str(_paths.EXHIB_PAG_OUT_KPIS_HISTORICO)
    cols_final = list(resumen.columns)
    if os.path.exists(ruta_hist):
        try:
            hist_prev = pd.read_excel(ruta_hist, engine='openpyxl')
            hist_prev['MES'] = pd.to_numeric(hist_prev.get('MES'), errors='coerce').fillna(0).astype(int)
            hist_prev['AÑO'] = pd.to_numeric(hist_prev.get('AÑO'), errors='coerce').fillna(0).astype(int)
        except Exception:
            hist_prev = pd.DataFrame(columns=cols_final)
    else:
        hist_prev = pd.DataFrame(columns=cols_final)

    # Sprint 15.9 (final fix) — upsert por periodo completo.
    if not hist_prev.empty and not resumen.empty:
        periodos_nuevos = set(map(tuple,
            resumen[['MES', 'AÑO']].drop_duplicates().values.tolist()))
        mask_keep = ~hist_prev.apply(
            lambda r: (int(r['MES']), int(r['AÑO'])) in periodos_nuevos,
            axis=1,
        )
        hist_prev = hist_prev[mask_keep]

    hist_final = pd.concat([hist_prev, resumen], ignore_index=True)
    hist_final = hist_final.sort_values(['AÑO', 'MES', 'EMPLEADO'])
    hist_final.to_excel(ruta_hist, index=False, engine='openpyxl')
    print(f"  ✅ Histórico → {_paths.EXHIB_PAG_OUT_KPIS_HISTORICO.name}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL Exhibiciones Pagadas — Sprint 16.1: multi-periodo")
    parser.add_argument("--solo", nargs="+", choices=["full", "kpi"],
                        help="Default: corre todo + kpi.")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    pasos = args.solo or ["full", "kpi"]
    specs = pr.periodos_de_args(args)

    if len(specs) > 1:
        print(f"🎯 ETL Exh Pagadas — multi-periodo: {len(specs)} meses → "
              f"{', '.join(s.etiqueta for s in specs)}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL Exh Pagadas — procesando periodo {spec.etiqueta} ({spec})")

        try:
            if "full" in pasos: ejecutar_proceso(spec)
            if "kpi"  in pasos: generar_resumen_kpi_exhibiciones_pagadas(spec)
        except Exception as e:
            print(f"\n❌ ERROR en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise