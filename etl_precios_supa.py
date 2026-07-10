import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

import pandas as pd
import os
import glob
import re
from datetime import datetime

# --- LIBRERÍAS REQUERIDAS PARA SUPABASE ---
from supabase import create_client

# =============================================================================
# 1. CONFIGURACIÓN GLOBAL
# =============================================================================

LISTA_CATEGORIAS_PRECIOS = [
    "PROTECCION FEMENINA", "JABONES DE TOCADOR", "ASEO DEL BEBE",
    "CREMAS CORPORALES", "CUIDADO FACIAL", "ENJUAGUE BUCAL", "CREMAS DENTALES"
]

MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# --- RUTAS resueltas vía paths.py ---
import paths
import periodo_resolver as pr

_DIR_PT = str(paths.CIF_PT_DIR)

RUTA_ORIGEN_PRECIOS = str(paths.PR_BASES)
RUTA_SALIDA_PRECIOS = str(paths.PR_SALIDA)

if not os.path.exists(RUTA_SALIDA_PRECIOS):
    os.makedirs(RUTA_SALIDA_PRECIOS)

# Patrón relajado para soportar "Respuestas de encuestas..." y "Respuestas Precios"
CLAVE_ARCHIVO_ENCUESTA = "Respuestas"

# --- CONEXIÓN DE SUPABASE ---
# (Se adaptará de forma automática en GitHub Actions al leer de los Secrets)
SUPABASE_URL = os.environ.get("SUPABASE_URL") or "https://scrhnipyveihqntyrykn.supabase.co"
SUPABASE_KEY = os.environ.get("SUPABASE_KEY") or "sb_secret_DrK4g3O4E4qBWuTPlHI3zg_J6cnHvup"
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# =============================================================================
# 2. FUNCIONES DE PROCESAMIENTO
# =============================================================================

def ejecutar_paso_1_consolidar_pt(spec: pr.PeriodoSpec):
    print(f"\n--- PASO 1: Consolidando Plan de Trabajo  ({spec.etiqueta}) ---")

    from shared_loader import COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR

    # Sprint 15.9 (Fix H4 final) — agregamos OBSERVACION para filtrar por
    # responsable real. La columna OBSERVACIÓN de "Captura de modulos" indica
    # qué ROL debe capturar el módulo en ese PDV (REPORTA GESTOR / REPORTA
    # SUPERVISOR / REPORTA GENERADOR DE DEMANDA). Sin este filtro, líderes
    # como INGRID quedaban incluidos por estar en el plan, aunque su rol no
    # sea capturar Precios.
    COLUMNAS_FINALES_PT = [
        "ID_PDV_INVOLVES", "NOMBRE_PDV", "VENTAS_PROMEDIO_MES", "ACRONIMO",
        "CEDULA", "NOMBRE", "COD_MERCADERISTA", "ROL", "SUPERVISOR_LIDER", "FUENTE",
        "OBSERVACION",
    ]

    # Mapeo OBSERVACIÓN → ROL esperado del responsable de captura
    OBS_A_ROL = {
        "REPORTA GESTOR":              "GESTOR",
        "REPORTA SUPERVISOR":          "SUPERVISOR",
        "REPORTA GENERADOR DE DEMANDA":"GENERADOR DE DEMANDA",
    }

    def leer_y_normalizar(ruta, hoja, fuente):
        if not os.path.exists(ruta): return pd.DataFrame()
        try:
            with pd.ExcelFile(ruta) as xls:
                df_val = pd.read_excel(xls, sheet_name="Captura de modulos")
                df_val.columns = df_val.columns.str.strip().str.upper()
                col_filtro = "PRECIOS_FINAL" if fuente == "ISM" else "PRECIOS"
                if col_filtro not in df_val.columns: return pd.DataFrame()

                # PDVs target Precios + su observación (rol responsable)
                obs_col = next((c for c in df_val.columns if c.startswith('OBSERVAC')), None)
                if obs_col is None:
                    print(f"⚠ {fuente}: hoja Captura de modulos sin columna OBSERVACIÓN")
                    return pd.DataFrame()
                # Sprint 15.9 (D10) — excluir PDVs subcanal DISCOUNTER (no se miden)
                df_val_filtrado = df_val[df_val[col_filtro] == 1].copy()
                if "SUB CANAL" in df_val_filtrado.columns:
                    n_antes_dc = len(df_val_filtrado)
                    df_val_filtrado = df_val_filtrado[
                        df_val_filtrado["SUB CANAL"].astype(str).str.strip().str.upper() != "DISCOUNTER"
                    ]
                    n_dc = n_antes_dc - len(df_val_filtrado)
                    if n_dc > 0:
                        print(f"  {fuente}: filtro DISCOUNTER → {n_dc} PDVs excluidos")
                pdvs_obs = (
                    df_val_filtrado[["ID PDV INVOLVES", obs_col]]
                    .rename(columns={obs_col: "OBSERVACION"})
                )
                pdvs_obs["OBSERVACION"] = pdvs_obs["OBSERVACION"].astype(str).str.strip().str.upper()
                pdvs_obs["ROL_ESPERADO"] = pdvs_obs["OBSERVACION"].map(OBS_A_ROL).fillna("")

                df = pd.read_excel(xls, sheet_name=hoja)
                df.columns = df.columns.str.strip().str.upper()

                # Filtro 1: PDVs target del módulo
                df = df[df["ID PDV INVOLVES"].isin(pdvs_obs["ID PDV INVOLVES"])].copy()
                df.rename(columns={col: COLUMNAS_ESTANDAR[col] for col in df.columns if col in COLUMNAS_ESTANDAR}, inplace=True)

                # Anexar OBSERVACIÓN + ROL_ESPERADO por PDV
                df = df.merge(pdvs_obs, left_on="ID_PDV_INVOLVES", right_on="ID PDV INVOLVES", how="left")

                # Filtro 2 (Sprint 15.9 fix H4 final): solo personas cuyo ROL
                # coincide con el responsable indicado en OBSERVACIÓN del PDV.
                df["ROL"] = df["ROL"].astype(str).str.strip().str.upper()
                mask_responsable = df["ROL"] == df["ROL_ESPERADO"]
                n_antes = len(df)
                df = df[mask_responsable].copy()
                print(f"  {fuente}: filtro OBSERVACIÓN → {n_antes} → {len(df)} filas (responsables reales)")

                df["FUENTE"] = fuente
                df["OBSERVACION"] = df["OBSERVACION"]

                for col in COLUMNAS_FINALES_PT:
                    if col not in df.columns: df[col] = ""

                return df[COLUMNAS_FINALES_PT]
        except PermissionError as e:
            raise PermissionError(
                f"No se puede leer {fuente} (archivo abierto en Excel?): {ruta}. "
                f"Cierra el archivo y vuelve a correr."
            ) from e
        except Exception as e:
            print(f"❌ Error leyendo {fuente}: {e}")
            raise

    from shared_loader import descubrir_archivos_pt
    ruta_directo, ruta_ism = descubrir_archivos_pt(_DIR_PT, periodo=spec)
    if not ruta_directo or not ruta_ism:
        raise FileNotFoundError(
            f"Precios ({spec.etiqueta}): faltan PT del periodo. "
            f"Directo={ruta_directo or 'N/D'} | ISM={ruta_ism or 'N/D'}"
        )
    df_dir = leer_y_normalizar(ruta_directo, "Plan de trabajo", "DIRECTO")
    df_ism = leer_y_normalizar(ruta_ism, "Plan de trabajo CIF", "ISM")

    # Sprint 15.5.5: periodo viene del spec, no se detecta del archivo.
    periodo = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"

    if not df_dir.empty or not df_ism.empty:
        df_total = pd.concat([df_dir, df_ism], ignore_index=True)
        df_total['ID_PDV_INVOLVES'] = df_total['ID_PDV_INVOLVES'].astype(str).str.strip()

        # Sprint 15.7 (Fix H2) — D7: mantenemos UNA fila por (PDV, persona).
        # Antes: groupby('ID_PDV_INVOLVES').first() dejaba sólo UNA persona
        # por PDV. Cuando varias personas tenían el mismo PDV (típico cuando
        # el PDV lo visita un gestor + su líder/supervisor) el groupby
        # descartaba todas menos una. Eso provocaba que gestores como MARIA
        # ANGELICA CASALLAS quedaran fuera de Precios aunque sí tenían PDVs
        # target.
        #
        # Ahora: dedup por (PDV, NOMBRE) — cada combinación cuenta una vez.
        df_total_unificado = df_total.drop_duplicates(
            subset=['ID_PDV_INVOLVES', 'NOMBRE'],
        ).reset_index(drop=True)

        ruta_pt_final = os.path.join(RUTA_ORIGEN_PRECIOS, f"Plan_Trabajo_Precios_{periodo}.xlsx")
        df_total_unificado.to_excel(ruta_pt_final, index=False)

        print(
            f"✅ EXITO: Plan unificado guardado en: Plan_Trabajo_Precios_{periodo}.xlsx "
            f"({len(df_total_unificado)} filas, {df_total_unificado['ID_PDV_INVOLVES'].nunique()} PDVs únicos)"
        )
        return df_total_unificado, periodo
    
    return pd.DataFrame(), periodo

def generar_reporte_captura_precios(df_pt, periodo_nom, spec: pr.PeriodoSpec):
    print(f"\n--- PASO 2: Cruzando Capturas por Categorías ({periodo_nom}) ---")
    # Sprint 15.5.5: encuestas del periodo solicitado, no glob por mtime.
    archivo_encuesta = pr.precios_respuestas(spec)
    print(f"   Encuestas: {archivo_encuesta.name}")
    df_enc = pd.read_excel(archivo_encuesta)

    def extraer_cat(texto):
        if pd.isna(texto): return None
        t = re.sub(r'[^A-Z0-9\s]', ' ', str(texto).upper())
        t = " ".join(t.split())
        for cat in LISTA_CATEGORIAS_PRECIOS:
            if cat in t: return cat
        return None

    df_enc['CATEGORIA_LIMPIA'] = df_enc['Rótulo de la encuesta'].apply(extraer_cat)
    df_enc_valida = df_enc.dropna(subset=['CATEGORIA_LIMPIA']).copy()
    
    # Normalizar ambos IDs con el helper canónico
    from shared_loader import id_a_str
    df_pt['ID_PDV_INVOLVES'] = id_a_str(df_pt['ID_PDV_INVOLVES'])
    df_enc_valida['ID del PDV'] = id_a_str(df_enc_valida['ID del PDV'])

    # --- NUEVA LÓGICA: IDENTIFICAR PDVs QUE PERTENECEN A MAYORISTAS ---
    # Convertimos a string y mayúsculas para buscar coincidencias de forma segura
    df_enc_valida['ROTULO_UPPER'] = df_enc_valida['Rótulo de la encuesta'].astype(str).str.upper()
    
    # Agrupamos por PDV para mapear si alguna de sus encuestas contiene la palabra "MAYORISTA"
    pdvs_mayoristas = df_enc_valida.groupby('ID del PDV')['ROTULO_UPPER'].apply(
        lambda x: any('MAYORISTA' in r for r in x)
    ).to_dict()

    # Agrupar las categorías únicas capturadas por cada Punto de Venta
    dict_enc = df_enc_valida.groupby('ID del PDV')['CATEGORIA_LIMPIA'].unique().to_dict()
    set_oficial_estandar = set(LISTA_CATEGORIAS_PRECIOS)
    res = []

    for id_pdv in df_pt['ID_PDV_INVOLVES'].unique():
        cats = dict_enc.get(id_pdv, [])
        set_cats = set(cats)
        
        # Evaluar si el PDV actual fue identificado como Mayorista
        es_mayorista = pdvs_mayoristas.get(id_pdv, False)
        
        # Definir la meta de categorías según el tipo de canal
        # Si es mayorista la meta es 1 categoría capturada; de lo contrario, se requieren las 7 estándar
        categorias_requeridas = 1 if es_mayorista else len(set_oficial_estandar)
        
        # Calcular faltantes con respecto al catálogo oficial completo (solo aplica lógico para los regulares)
        if es_mayorista:
            # Para mayoristas, si ya capturó al menos 1, no tiene faltantes críticas para su canal
            faltantes = [] if len(set_cats) >= 1 else ["AL MENOS 1 CATEGORÍA"]
            conteo = len(set_cats)
            msg_faltantes = "NINGUNA" if len(set_cats) >= 1 else "REQUIERE 1 CATEGORÍA"
        else:
            faltantes = sorted(list(set_oficial_estandar - set_cats))
            conteo = len(set_cats)
            msg_faltantes = ", ".join(faltantes) if faltantes else "NINGUNA"
            
        if conteo == 0: 
            msg_faltantes = f"SIN CAPTURAS (FALTAN {categorias_requeridas})"

        res.append({
            'ID_PDV_INVOLVES': id_pdv, 
            'CAPTURA_PLANEADA': 1, 
            'CONTEO_CATEGORIAS': conteo, 
            'CATEGORIAS_FALTANTES': msg_faltantes,
            # Se marca ejecutado (1) si cumple la meta de su canal (>=1 para mayorista, ==7 para regular)
            'CAPTURA_EJECUTADA': 1 if conteo >= categorias_requeridas else 0
        })
    
    df_final = pd.merge(df_pt, pd.DataFrame(res), on='ID_PDV_INVOLVES', how='left')
    df_final['CAPTURA_PLANEADA'] = df_final['CAPTURA_PLANEADA'].fillna(1).astype(int)
    df_final['CAPTURA_EJECUTADA'] = df_final['CAPTURA_EJECUTADA'].fillna(0).astype(int)

    # Sprint 15.5.5 — Fix F3: propagar MES/AÑO al reporte intermedio
    df_final['MES'] = spec.mes
    df_final['AÑO'] = spec.anio

    cols_out = list(df_pt.columns) + [
        'CAPTURA_PLANEADA', 'CONTEO_CATEGORIAS', 'CATEGORIAS_FALTANTES',
        'CAPTURA_EJECUTADA', 'MES', 'AÑO',
    ]
    ruta_matriz = os.path.join(RUTA_SALIDA_PRECIOS, f"REPORTE_CAPTURA_PRECIOS_{periodo_nom}.xlsx")
    df_final[cols_out].to_excel(ruta_matriz, index=False)
    print(f"✅ Reporte capturas generado.")

def generar_analisis_precios(periodo_nom, spec: pr.PeriodoSpec):
    print(f"\n--- PASO 3: Generando Análisis Detallado ({periodo_nom}) ---")
    archivo_encuesta = pr.precios_respuestas(spec)
    df = pd.read_excel(archivo_encuesta)
    mapeo_binario = {'SI': 1, 'SÍ': 1, 'NO': 0, 'no': 0, 'si': 1, 'sí': 1}

    if "Producto (SKU)" in df.columns:
        split_sku = df["Producto (SKU)"].astype(str).str.split(' - ', n=1, expand=True)
        df['CODIGO_SKU'] = split_sku[0].str.strip()
        df['NOMBRE_PRODUCTO'] = split_sku[1].str.strip() if split_sku.shape[1] > 1 else ""

    df['PRESENCIA'] = df["El producto está presente en el PDV?"].astype(str).str.strip().str.upper().map(mapeo_binario).fillna(0).astype(int)
    df['PRECIO_REGULAR'] = pd.to_numeric(df["Digite Precio Regular"], errors='coerce').fillna(0)
    df['LABEL_UBICADO'] = df["El producto cuenta con el Label (Precio impreso) ubicado?"].astype(str).str.strip().str.upper().map(mapeo_binario).fillna(0).astype(int)
    df['HAY_PROMO'] = df["Hay Promociones o Descuentos?"].astype(str).str.strip().str.upper().map(mapeo_binario).fillna(0).astype(int)
    df['PRECIO_PROMO'] = pd.to_numeric(df["Digite Precio Promoción:"], errors='coerce').fillna(0)
    df['TIPO_PROMO'] = df["Seleccione la promoción:"].fillna("SIN PROMOCION")

    cols_finales = ["ID del PDV", "PDV", "Fecha de la encuesta", "Empleado", "Marca", 
                    "CODIGO_SKU", "NOMBRE_PRODUCTO", "PRESENCIA", "PRECIO_REGULAR", 
                    "LABEL_UBICADO", "HAY_PROMO", "PRECIO_PROMO", "TIPO_PROMO"]
    
    ruta_analisis = os.path.join(RUTA_SALIDA_PRECIOS, f"ANALISIS_PRECIOS_{periodo_nom}.xlsx")
    df[cols_finales].to_excel(ruta_analisis, index=False)
    print(f"✅ Análisis detallado guardado.")

# =============================================================================
# 3. EJECUCIÓN PRINCIPAL Y INTEGRACIÓN CON CLOUD
# =============================================================================

def generar_resumen_kpi_precios(spec: pr.PeriodoSpec):
    """Lee el reporte de captura del periodo solicitado y produce
    PRECIOS_KPIS.xlsx + histórico acumulado (V3) e impacta la nube en Supabase.
    """
    print(f"\n--- PASO KPI (V3): RESUMEN PRECIOS por gestor  ({spec.etiqueta}) ---")
    import paths as _paths
    periodo_nom = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"
    ruta = _paths.PR_SALIDA / f"REPORTE_CAPTURA_PRECIOS_{periodo_nom}.xlsx"
    if not ruta.is_file():
        print(f"❌ No existe el reporte del periodo: {ruta}")
        return
    df = pd.read_excel(ruta, engine='openpyxl')
    print(f"  Leyendo: {ruta.name} ({len(df)} filas)")

    from shared_loader import calcular_kpi_simple_y_escribir
    n = calcular_kpi_simple_y_escribir(
        df_origen=df,
        col_planeado="CAPTURA_PLANEADA",
        col_ejecutado="CAPTURA_EJECUTADA",
        ruta_kpi_mes=str(_paths.PR_OUT_KPIS),
        ruta_kpi_historico=str(_paths.PR_OUT_KPIS_HISTORICO),
        nombre_cumplimiento="CUMPLIMIENTO",
        nombres_renombre=("PLANEADO", "EJECUTADO"),
    )
    print(f"  ✅ Mes activo: {n} gestores → {_paths.PR_OUT_KPIS.name}")
    print(f"  ✅ Histórico acumulado → {_paths.PR_OUT_KPIS_HISTORICO.name}")

    # --- SECCIÓN DE CARGA AUTOMÁTICA A SUPABASE ---
    try:
        if os.path.exists(str(_paths.PR_OUT_KPIS)):
            print(f"  🚀 Preparando cargue del KPI de Precios consolidado a Supabase...")
            df_kpi = pd.read_excel(str(_paths.PR_OUT_KPIS))
            
            # Normalizamos nombres de columnas a minúsculas para hacer match con SQL
            df_kpi.columns = df_kpi.columns.str.strip().str.lower()
            
            # Reemplazo seguro de la 'ñ' en el dataframe para encajar con el campo de Postgres
            if 'año' in df_kpi.columns:
                df_kpi = df_kpi.rename(columns={'año': 'anio'})
            
            # AJUSTE FIJO SPRINT 16.1: Reemplazo seguro de NaN matemáticos a None usando where
            df_kpi_limpio = df_kpi.where(pd.notnull(df_kpi), None)
            registros_json = df_kpi_limpio.to_dict(orient="records")
            
            # Envío vía Upsert: inserta o actualiza basándose en la llave compuesta (nombre, mes, anio)
            supabase.table("precios_kpis").upsert(registros_json).execute()
            print(f"  ✅ ÉXITO CLOUD: {len(registros_json)} filas subidas de forma exitosa a Supabase.")
    except Exception as ex_cloud:
        print(f"  ⚠ ADVERTENCIA EN CARGA CLOUD (Los archivos Excel locales se generaron bien): {ex_cloud}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL PRECIOS — Eficacia (Sprint 16.1: multi-periodo)")
    parser.add_argument("--solo", nargs="+", choices=["full", "kpi"],
                        help="Por default: corre todo + kpi. Usa --solo kpi para solo KPI.")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    pasos = args.solo or ["full", "kpi"]
    specs = pr.periodos_de_args(args)

    if len(specs) > 1:
        print(f"🎯 ETL PRECIOS — multi-periodo: {len(specs)} meses → "
              f"{', '.join(s.etiqueta for s in specs)}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL PRECIOS — procesando periodo {spec.etiqueta} ({spec})")

        df_pt = pd.DataFrame()
        periodo_final = f"{MESES_ESPANOL[spec.mes]}_{spec.anio}"

        try:
            if "full" in pasos:
                df_pt, periodo_final = ejecutar_paso_1_consolidar_pt(spec)

                if df_pt.empty:
                    ruta_pt_periodo = os.path.join(
                        RUTA_ORIGEN_PRECIOS,
                        f"Plan_Trabajo_Precios_{periodo_final}.xlsx",
                    )
                    if os.path.isfile(ruta_pt_periodo):
                        df_pt = pd.read_excel(ruta_pt_periodo)
                        print(f"📂 Cargado PT del periodo: {os.path.basename(ruta_pt_periodo)}")

                if not df_pt.empty:
                    generar_reporte_captura_precios(df_pt, periodo_final, spec)
                generar_analisis_precios(periodo_final, spec)

            if "kpi" in pasos:
                generar_resumen_kpi_precios(spec)

            print(f"\n🏁 PROCESO TERMINADO  ({spec.etiqueta})")

        except Exception as e:
            print(f"\n❌ ERROR en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise