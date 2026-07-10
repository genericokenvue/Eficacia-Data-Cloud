import sys
# Forzar UTF-8 en stdout (Windows cp1252 explota con emojis 📅 ⚠️ ❌)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

import pandas as pd
import numpy as np
import sqlite3
import os
from datetime import datetime, timedelta

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE RUTAS  ── todas resueltas en paths.py
# ─────────────────────────────────────────────
import paths
import periodo_resolver as pr

_DIR_PT = str(paths.CIF_PT_DIR)

# Archivos de salida intermedios (CSVs)
RUTA_CONSOLIDADO = str(paths.CIF_CONSOLIDADO_CSV)
RUTA_AGRUPADO    = str(paths.CIF_AGRUPADO_CSV)

# Fuentes externas que NO dependen del periodo (catálogos estables)
RUTA_COLABORADORES   = str(paths.CIF_COLABORADORES)
RUTA_OUTPUT_INVOLVES = str(paths.CIF_INVOLVES_PROCESADO)
RUTA_BASE_VENTAS     = str(paths.CIF_BASE_VENTAS)
RUTA_CAUSALES        = str(paths.CIF_CAUSALES)
# RUTA_INVOLVES — eliminada. Sprint 15.5.3: el archivo Involves es por mes,
# se resuelve dentro de los pasos 3 y 4 vía pr.cif_involves(spec).

# Aliases legacy — apuntan al MISMO archivo que RUTA_CONSOLIDADO.
RUTA_PLAN_CONSOLIDADO = RUTA_CONSOLIDADO
RUTA_PLAN_ORIGINAL    = RUTA_CONSOLIDADO

# Bloque 5 — derivadas de Involves
ruta_entrada = RUTA_OUTPUT_INVOLVES  # archivo generado en el Paso 4
ruta_salida  = RUTA_OUTPUT_INVOLVES.replace(".csv", "_agrupado_final.csv")

# Bloque 7 — salida final
RUTA_PLAN_AGRUPADO    = RUTA_AGRUPADO  # alias legacy
RUTA_AGRUPADO_VISITAS = ruta_salida    # mismo archivo que paso 5 produce
RUTA_FINAL_CIF_EXCEL  = str(paths.CIF_OUT_FINAL)


# ─────────────────────────────────────────────
# MAPEO COLUMNAS  ── importado desde shared_loader (fuente canónica)
# ─────────────────────────────────────────────
from shared_loader import (
    COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR,
    COLUMNAS_SUPERSET            as COLUMNAS_FINALES,
    COLUMNAS_NUMERICAS,
)

# ─────────────────────────────────────────────
# BLOQUE 1: CONSOLIDACIÓN PLAN DE TRABAJO DIRECTO Y PROXIMITY
# ─────────────────────────────────────────────
def leer_y_normalizar(ruta, hoja, fuente):
    if not os.path.exists(ruta):
        print(f"⚠️  Archivo no encontrado: {fuente}")
        return pd.DataFrame()

    df = pd.read_excel(ruta, sheet_name=hoja)
    df.columns = df.columns.str.strip()
    df.rename(columns={col: COLUMNAS_ESTANDAR[col] for col in df.columns if col in COLUMNAS_ESTANDAR}, inplace=True)
    df["FUENTE"] = fuente

    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors='coerce').dt.strftime("%d/%m/%Y")

    for col in COLUMNAS_NUMERICAS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df[[c for c in COLUMNAS_FINALES if c in df.columns]].copy()
    
    for col in df.columns:
        if col not in COLUMNAS_NUMERICAS:
            df[col] = df[col].astype(str).str.strip().replace(['nan', 'NaT', 'None'], '')

    print(f"   [{fuente}] {len(df)} filas procesadas.")
    return df

def ejecutar_paso_1(spec: pr.PeriodoSpec):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 1: CONSOLIDACIÓN  ({spec.etiqueta})")
    print("="*40)
    from shared_loader import descubrir_archivos_pt
    ruta_directo, ruta_ism = descubrir_archivos_pt(_DIR_PT, periodo=spec)
    if not ruta_directo or not ruta_ism:
        raise FileNotFoundError(
            f"Paso 1 ({spec.etiqueta}): faltan PT del periodo. "
            f"Directo={ruta_directo or 'N/D'} | ISM={ruta_ism or 'N/D'}"
        )
    df_dir = leer_y_normalizar(ruta_directo, "Plan de trabajo", "DIRECTO")
    df_ism = leer_y_normalizar(ruta_ism, "Plan de trabajo CIF", "ISM")
    
    df_total = pd.concat([df_dir, df_ism], ignore_index=True)
    
    for col in COLUMNAS_FINALES:
        if col not in df_total.columns:
            df_total[col] = ""
            
    df_total[COLUMNAS_FINALES].to_csv(RUTA_CONSOLIDADO, index=False, encoding="utf-8-sig", sep=";", decimal=",")
    print(f"\n✅ ÉXITO: Consolidado guardado en: {RUTA_CONSOLIDADO}")

# ─────────────────────────────────────────────
# BLOQUE 2: AGRUPACIÓN PLAN DE TRABAJO
# ─────────────────────────────────────────────
def ejecutar_paso_2():
    print("\n" + "="*40)
    print(" INICIANDO PASO 2: AGRUPACIÓN SQL")
    print("="*40)
    
    if not os.path.exists(RUTA_CONSOLIDADO):
        print("❌ ERROR: No existe el archivo consolidado. Ejecuta el Paso 1 primero.")
        return

    df = pd.read_csv(RUTA_CONSOLIDADO, sep=";", decimal=",", encoding="utf-8-sig", 
                     dtype={'ID_PDV_INVOLVES': str, 'CEDULA': str})

    conexion = sqlite3.connect(':memory:')
    df.to_sql('plan_consolidado', conexion, index=False, if_exists='replace')

    query = """
    SELECT 
        ID_PDV_INVOLVES, ACRONIMO,
        MAX(NOMBRE_PDV) AS NOMBRE_PDV,
        MAX(VENTAS_PROMEDIO_MES) AS VENTAS_PROMEDIO_MES,
        MAX(CEDULA) AS CEDULA,
        MAX(NOMBRE) AS NOMBRE,
        MAX(COD_MERCADERISTA) AS COD_MERCADERISTA,
        COUNT(FECHA) AS CANTIDAD_VISITAS,
        SUM(TIEMPO_SERVICIO) AS TIEMPO_SERVICIO,
        (SUM(TIEMPO_SERVICIO) / 60.0) AS TIEMPO_SERVICIO_HORAS,
        MAX(FREC_SEMANAL) AS FREC_SEMANAL,
        MAX(FREC_MENSUAL) AS FREC_MENSUAL,
        MAX(HORAS_X_VISITA) AS HORAS_X_VISITA,
        MAX(HORAS_MES) AS HORAS_MES,
        MAX(ROL) AS ROL,
        MAX(SUPERVISOR_LIDER) AS SUPERVISOR_LIDER,
        MAX(FUENTE) AS FUENTE
    FROM plan_consolidado
    GROUP BY ACRONIMO, ID_PDV_INVOLVES
    """

    df_agrupado = pd.read_sql_query(query, conexion)
    conexion.close()

    df_agrupado.to_csv(RUTA_AGRUPADO, index=False, sep=";", decimal=",", encoding="utf-8-sig")
    print(f"✅ ÉXITO: Agrupado guardado en: {RUTA_AGRUPADO}")
    print("\nResumen agrupado (Primeros 5):")
    print(df_agrupado[['ACRONIMO', 'ID_PDV_INVOLVES', 'TIEMPO_SERVICIO_HORAS']].head())





# ─────────────────────────────────────────────
# BLOQUE 3: INVOLVES VISITAS REALIZADAS Y VALIDAS
# ─────────────────────────────────────────────
def ejecutar_paso_3(spec: pr.PeriodoSpec):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 3: INVOLVES (CRUCE PRIORIZANDO BASE VENTAS)  ({spec.etiqueta})")
    print("="*40)

    # Sprint 15.5.3 — Fix F1: Involves se resuelve por periodo
    # (informe-gerencial-visitas_YYYYMM.xlsx), no archivo sin sufijo.
    RUTA_INVOLVES = str(pr.cif_involves(spec))

    # Validar existencia de archivos
    rutas_necesarias = [RUTA_INVOLVES, RUTA_COLABORADORES, RUTA_PLAN_CONSOLIDADO, RUTA_BASE_VENTAS]
    for r in rutas_necesarias:
        if not os.path.exists(r):
            print(f"❌ ERROR: No se encontró el archivo en: {r}")
            return

    # 1. CARGA DE MAESTROS CON LIMPIEZA DE IDs
    df_maestro_plan = pd.read_csv(RUTA_PLAN_CONSOLIDADO, sep=";", decimal=",", encoding="utf-8-sig")
    df_maestro_plan['ID_PDV_INVOLVES'] = pd.to_numeric(df_maestro_plan['ID_PDV_INVOLVES'], errors='coerce').fillna(0).astype(int).astype(str)
    df_maestro_plan = df_maestro_plan[['NOMBRE_PDV', 'ID_PDV_INVOLVES']].drop_duplicates(subset=['NOMBRE_PDV'])
    df_maestro_plan.rename(columns={'NOMBRE_PDV': 'P_NOMBRE', 'ID_PDV_INVOLVES': 'ID_PLAN'}, inplace=True)

    df_maestro_ventas = pd.read_excel(RUTA_BASE_VENTAS)
    df_maestro_ventas['ID'] = pd.to_numeric(df_maestro_ventas['ID'], errors='coerce').fillna(0).astype(int).astype(str)
    df_maestro_ventas = df_maestro_ventas[['Nombre del PDV', 'ID']].drop_duplicates(subset=['Nombre del PDV'])
    df_maestro_ventas.rename(columns={'Nombre del PDV': 'V_NOMBRE', 'ID': 'ID_VENTAS'}, inplace=True)

    # 2. CARGAR COLABORADORES
    df_colab = pd.read_excel(RUTA_COLABORADORES)[['Nombre del empleado', 'Usuario', 'Perfil de acceso']].drop_duplicates(subset=['Nombre del empleado'])
    df_colab.rename(columns={'Nombre del empleado': 'EMPLEADO_KEY', 'Usuario': 'USUARIO_INVOLVES', 'Perfil de acceso': 'PERFIL_ACCESO'}, inplace=True)

    # 3. Cargar Involves y cruzar con Colaboradores
    df_inv_raw = pd.read_excel(RUTA_INVOLVES, sheet_name="Sheet0")
    df = df_inv_raw.merge(df_colab, left_on='Empleado', right_on='EMPLEADO_KEY', how='left')
    
    # 4. Seleccionar y renombrar campos
    columnas_interes = {
        'Punto de venta': 'PUNTO_DE_VENTA',
        'Empleado': 'EMPLEADO',
        'USUARIO_INVOLVES': 'USUARIO_INVOLVES', 
        'PERFIL_ACCESO': 'PERFIL_ACCESO',       
        'Fecha de la visita': 'FECHA_VISITA',
        'Tipo de check-in': 'TIPO_CHECKIN',
        'Primer check-in manual': 'CHECKIN_MANUAL',
        'Último check-out manual': 'CHECKOUT_MANUAL',
        'Check-in automático': 'CHECKIN_AUTO',
        'Check-out automático': 'CHECKOUT_AUTO',
        'Justificación': 'JUSTIFICACION',
        'Justificación auditada el': 'JUSTIFICADO_EL',
        'Quién justificó': 'QUIEN_JUSTIFICO'
    }
    
    df = df[[c for c in columnas_interes.keys() if c in df.columns]].rename(columns=columnas_interes)
    df = df[df['TIPO_CHECKIN'].isin(['Manual', 'Manual y GPS'])].copy()

    # 5. CRUCES DE ID_PDV
    df = df.merge(df_maestro_ventas, left_on='PUNTO_DE_VENTA', right_on='V_NOMBRE', how='left')
    df = df.merge(df_maestro_plan, left_on='PUNTO_DE_VENTA', right_on='P_NOMBRE', how='left')
    df['ID_PDV_INVOLVES'] = df['ID_VENTAS'].fillna(df['ID_PLAN'])

    # 6. Cálculos de Tiempo
    time_cols = ['CHECKIN_MANUAL', 'CHECKOUT_MANUAL', 'CHECKIN_AUTO', 'CHECKOUT_AUTO']
    for col in time_cols:
        df[col] = pd.to_datetime(df[col], format='%H:%M:%S', errors='coerce')

    # ── VECTORIZADO: reemplaza apply(calcular_minutos, axis=1) ── 35x más rápido ──
    # Condición 1: sin CHECKOUT_MANUAL → 0
    sin_checkout = df['CHECKOUT_MANUAL'].isna()

    # Condición 2: cualquier checkout a las 23:59 (visita no cerrada) → 0
    es_2359 = (
        (df['CHECKOUT_MANUAL'].notna() & df['CHECKOUT_MANUAL'].dt.hour.eq(23) & df['CHECKOUT_MANUAL'].dt.minute.eq(59)) |
        (df['CHECKOUT_AUTO'].notna()   & df['CHECKOUT_AUTO'].dt.hour.eq(23)   & df['CHECKOUT_AUTO'].dt.minute.eq(59))
    )

    # Deltas base en minutos (clip evita negativos sin condicional explícito)
    t_manual = (df['CHECKOUT_MANUAL'] - df['CHECKIN_MANUAL']).dt.total_seconds().div(60).clip(lower=0)
    t_auto   = (df['CHECKOUT_AUTO']   - df['CHECKIN_MANUAL']).dt.total_seconds().div(60).clip(lower=0)

    # Si hay CHECKOUT_AUTO y |dif| > 60 min vs manual → usar t_auto; si no → t_manual
    hay_auto        = df['CHECKOUT_AUTO'].notna()
    dif_auto_manual = (df['CHECKOUT_AUTO'] - df['CHECKOUT_MANUAL']).dt.total_seconds().div(60)
    usar_auto       = hay_auto & (dif_auto_manual.abs() > 60)
    tiempo_base     = np.where(usar_auto, t_auto, t_manual)

    # Aplicar invalidaciones y guardar resultado
    df['TIEMPO_SERVICIO_MIN'] = np.where(sin_checkout | es_2359, 0.0, tiempo_base).clip(min=0)
    df['VISITA_OK'] = df['TIEMPO_SERVICIO_MIN'] >= 10

    # 7. Limpieza y Formato Final
    for col in time_cols:
        df[col] = df[col].dt.strftime('%H:%M:%S')

    columnas_a_quitar = ['EMPLEADO_KEY', 'V_NOMBRE', 'ID_VENTAS', 'P_NOMBRE', 'ID_PLAN']
    cols_finales = ['ID_PDV_INVOLVES'] + [c for c in df.columns if c not in columnas_a_quitar and c != 'ID_PDV_INVOLVES']
    
    df_detalle_final = df[cols_finales].drop_duplicates()

    df_detalle_final.to_csv(RUTA_OUTPUT_INVOLVES, index=False, sep=";", decimal=",", encoding="utf-8-sig")
    print(f"✅ Paso 3 finalizado. Condición de Checkout Manual obligatoria aplicada.")


# ─────────────────────────────────────────────
# BLOQUE 4: INVOLVES VISITAS JUSTIFICADAS
# ─────────────────────────────────────────────

def ejecutar_paso_4(spec: pr.PeriodoSpec):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 4: JUSTIFICACIONES (CRUCE PRIORIZANDO BASE VENTAS)  ({spec.etiqueta})")
    print("="*40)

    # Sprint 15.5.3 — Fix F1: Involves resuelto por periodo.
    RUTA_INVOLVES = str(pr.cif_involves(spec))

    # 1. CARGAR DICCIONARIO DE CAUSALES
    if not os.path.exists(RUTA_CAUSALES):
        print(f"❌ ERROR: No se encontró el archivo de causales")
        return

    df_causales_maestro = pd.read_excel(RUTA_CAUSALES)
    CAUSALES_CONFIG = dict(zip(
        df_causales_maestro['Motivo para no realizar la visita'].str.strip().str.upper(),
        df_causales_maestro['Validez']
    ))

    # 2. CARGAR ORIGEN INVOLVES (Solo los "Sin check-in")
    df_inv_raw = pd.read_excel(RUTA_INVOLVES, sheet_name="Sheet0")
    df_just = df_inv_raw[
        (df_inv_raw['Tipo de check-in'] == 'Sin check-in') & 
        (df_inv_raw['Motivo para no realizar la visita'].notnull())
    ].copy()

    if df_just.empty:
        print("⚠️ No hay motivos de inasistencia para procesar.")
        return

    # 3. RENOMBRAR Y PREPARAR COLUMNAS
    columnas_just = {
        'Punto de venta': 'PUNTO_DE_VENTA', 
        'Empleado': 'EMPLEADO',
        'Fecha de la visita': 'FECHA_VISITA', 
        'Tipo de check-in': 'TIPO_CHECKIN',
        'Motivo para no realizar la visita': 'MOTIVO_REAL',
        'Justificación': 'OBSERVACION_JUSTIFICACION', 
        'Justificación auditada el': 'JUSTIFICADO_EL',
        'Quién justificó': 'QUIEN_JUSTIFICO'
    }
    df_just = df_just[[c for c in columnas_just.keys() if c in df_just.columns]].rename(columns=columnas_just)

    # 4. CARGAR MAESTROS DE CRUCE (ID y Datos Personales)
    # Maestro 1: Plan de Trabajo (Respaldo) - Limpiamos el ID_PDV_INVOLVES
    df_plan = pd.read_csv(RUTA_PLAN_CONSOLIDADO, sep=";", decimal=",", encoding="utf-8-sig")
    df_plan['ID_PDV_INVOLVES'] = pd.to_numeric(df_plan['ID_PDV_INVOLVES'], errors='coerce').fillna(0).astype(int).astype(str)
    
    df_maestro_info = df_plan[['NOMBRE', 'NOMBRE_PDV', 'ID_PDV_INVOLVES', 'CEDULA', 'ROL', 'HORAS_X_VISITA']].drop_duplicates()
    df_maestro_validador = df_plan[['NOMBRE', 'CEDULA']].drop_duplicates(subset=['NOMBRE'])

    # Maestro 2: Base Ventas (Prioridad) - Limpiamos el ID
    df_maestro_ventas = pd.read_excel(RUTA_BASE_VENTAS)
    df_maestro_ventas['ID'] = pd.to_numeric(df_maestro_ventas['ID'], errors='coerce').fillna(0).astype(int).astype(str)
    df_maestro_ventas = df_maestro_ventas[['Nombre del PDV', 'ID']].drop_duplicates(subset=['Nombre del PDV'])

    # 5. CRUCE PRINCIPAL (Identidad y Datos del Plan)
    df_just = df_just.merge(
        df_maestro_info, 
        left_on=['EMPLEADO', 'PUNTO_DE_VENTA'], 
        right_on=['NOMBRE', 'NOMBRE_PDV'], 
        how='left'
    )

    # 6. CRUCE DE ID PRIORITARIO (Lógica Inversa: Ventas -> Plan)
    df_just = df_just.merge(df_maestro_ventas, left_on='PUNTO_DE_VENTA', right_on='Nombre del PDV', how='left')
    
    # Si ID de Ventas (ID) es nulo, toma el del Plan (ID_PDV_INVOLVES)
    df_just['ID_FINAL'] = df_just['ID'].fillna(df_just['ID_PDV_INVOLVES'])

    # 7. CRUCE SECUNDARIO (Cédula del validador)
    df_just = df_just.merge(
        df_maestro_validador, 
        left_on='QUIEN_JUSTIFICO', 
        right_on='NOMBRE', 
        how='left', 
        suffixes=('', '_VAL')
    )
    
    # Renombrar y limpiar
    df_just.rename(columns={'CEDULA_VAL': 'CEDULA_QUIEN_JUSTIFICO', 'ID_FINAL': 'ID_PDV_RESULTADO'}, inplace=True)
    cols_drop = ['NOMBRE_x', 'NOMBRE_y', 'NOMBRE_PDV', 'NOMBRE', 'ID', 'ID_PDV_INVOLVES', 'Nombre del PDV']
    df_just.drop(columns=[c for c in cols_drop if c in df_just.columns], inplace=True, errors='ignore')
    df_just.rename(columns={'ID_PDV_RESULTADO': 'ID_PDV_INVOLVES'}, inplace=True)

    # 8. CALCULAR VALIDEZ Y TIEMPO DINÁMICO
    df_just['VALIDEZ_JUSTIFICACION'] = df_just['MOTIVO_REAL'].apply(
        lambda x: CAUSALES_CONFIG.get(str(x).strip().upper(), 0)
    )
    
    # ── VECTORIZADO: reemplaza apply() de 2 columnas ── 7.8x más rápido ──
    # VALIDEZ_JUSTIFICACION ya es columna entera en este punto
    df_just['TIEMPO_SERVICIO_MIN'] = np.where(
        df_just['VALIDEZ_JUSTIFICACION'] == 1,
        df_just['HORAS_X_VISITA'] * 60,
        0
    )
    df_just['VISITA_OK'] = df_just['TIEMPO_SERVICIO_MIN'] > 0

    # 9. INTEGRACIÓN FINAL
    if os.path.exists(RUTA_OUTPUT_INVOLVES):
        df_paso3 = pd.read_csv(RUTA_OUTPUT_INVOLVES, sep=";", decimal=",", dtype={'ID_PDV_INVOLVES': str})
        
        df_final_total = pd.concat([df_paso3, df_just], ignore_index=True, sort=False)
        
        columnas_a_excluir = ['VALIDEZ_JUSTIFICACION', 'HORAS_X_VISITA']
        cols_finales = [c for c in df_final_total.columns if c not in columnas_a_excluir]
        
        df_final_total[cols_finales].drop_duplicates().to_csv(RUTA_OUTPUT_INVOLVES, index=False, sep=";", decimal=",", encoding="utf-8-sig")

        print(f"✅ Paso 4 finalizado: IDs unificados (Ventas > Plan) y justificaciones procesadas.")

# ─────────────────────────────────────────────
# BLOQUE 5: INVOLVES AGRUPAR ARCHIVO INVOLVES
# ─────────────────────────────────────────────

def ejecutar_paso_5():
    print("\n" + "="*40)
    print(" INICIANDO PASO 5: CONSOLIDACIÓN (CORRECCIÓN DE ID FORMATO .0)")
    print("="*40)
    

    if not os.path.exists(ruta_entrada):
        print(f"❌ ERROR: No se encuentra el archivo.")
        return

    try:
        # 1. Cargar datos
        df = pd.read_csv(ruta_entrada, sep=";", decimal=",", encoding="utf-8-sig")

        # --- CORRECCIÓN DEL ID 10005.0 ---
        # Primero llenamos nulos con 0, convertimos a entero para quitar el .0 y luego a texto
        df['ID_PDV_INVOLVES'] = pd.to_numeric(df['ID_PDV_INVOLVES'], errors='coerce').fillna(0).astype(int).astype(str)
        # ---------------------------------

        df['EMPLEADO'] = df['EMPLEADO'].astype(str).str.strip()
        df['PERFIL_ACCESO'] = df['PERFIL_ACCESO'].fillna('').astype(str).str.upper().str.strip()

        # 2. Identificar al GESTOR TITULAR (Dueño real del punto)
        mask_titular = (df['PERFIL_ACCESO'] == 'GESTOR') & (~df['EMPLEADO'].str.contains('SUPERNUMERARIO', case=False, na=False))
        df_titulares = df[mask_titular].sort_values('TIEMPO_SERVICIO_MIN', ascending=False)
        mapa_titulares = df_titulares.drop_duplicates('ID_PDV_INVOLVES').set_index('ID_PDV_INVOLVES')['EMPLEADO'].to_dict()

        # 3. Unificación de identidad — vectorizado con map() + np.where ── 235x más rápido
        # Máscara: filas que son apoyo (BACKUP o SUPERNUMERARIO)
        es_backup         = df['PERFIL_ACCESO'] == 'BACKUP'
        es_supernumerario = df['EMPLEADO'].str.contains('SUPERNUMERARIO', case=False, na=False)
        es_apoyo          = es_backup | es_supernumerario

        # Buscar titular del PDV en el mapa construido arriba
        titular_mapeado = df['ID_PDV_INVOLVES'].map(mapa_titulares)  # NaN si no hay titular
        tiene_titular   = titular_mapeado.notna()

        # Sustituir solo donde es apoyo Y existe un titular conocido
        reemplazar = es_apoyo & tiene_titular
        df['EMPLEADO_UNIFICADO'] = np.where(reemplazar, titular_mapeado, df['EMPLEADO'])
        df['PERFIL_UNIFICADO']   = np.where(reemplazar, 'GESTOR',        df['PERFIL_ACCESO'])

        # 4. Agrupación Final
        df_final = df.groupby(['ID_PDV_INVOLVES', 'EMPLEADO_UNIFICADO', 'PERFIL_UNIFICADO']).agg({
            'PUNTO_DE_VENTA': 'first',
            'USUARIO_INVOLVES': 'first',
            'TIEMPO_SERVICIO_MIN': 'sum',
            'VISITA_OK': 'sum'
        }).reset_index()

        # 5. Formatear salida
        df_final = df_final[[
            'ID_PDV_INVOLVES', 'PUNTO_DE_VENTA', 'EMPLEADO_UNIFICADO', 
            'USUARIO_INVOLVES', 'PERFIL_UNIFICADO', 'TIEMPO_SERVICIO_MIN', 'VISITA_OK'
        ]]
        
        df_final.columns = [
            'ID_PDV_INVOLVES', 'PUNTO_DE_VENTA', 'EMPLEADO', 
            'USUARIO_INVOLVES', 'PERFIL_ACCESO', 'SUMA_TIEMPO_SERVICIO_MIN', 'CONTEO_VISITAS_OK'
        ]

        # Guardar (Aseguramos que no escriba nulos como .0)
        df_final.to_csv(ruta_salida, index=False, sep=";", decimal=",", encoding="utf-8-sig")
        
        print(f"✅ ¡Proceso Limpio! IDs corregidos y apoyos sumados al Gestor.")

    except Exception as e:
        print(f"❌ Error: {e}")

# ─────────────────────────────────────────────
# BLOQUE 6: ARCHIVO ALERTA CUANDO EL TIEMPO DE SERVICIO MAYOR A 540 MIN
# ─────────────────────────────────────────────

def ejecutar_paso_6():
    print("\n" + "="*40)
    print(" INICIANDO PASO 6: GENERACIÓN DE ALERTAS (RUTA CIF SALIDA)")
    print("="*40)

    # 1. Ruta de salida (resuelta vía paths.py)
    RUTA_ALERTA_CIF = str(paths.CIF_OUT_ALERTAS)

    # 2. Validar existencia del archivo base procesado
    if not os.path.exists(RUTA_OUTPUT_INVOLVES):
        print(f"❌ ERROR: No se encontró el archivo base en: {RUTA_OUTPUT_INVOLVES}")
        return

    # 3. Cargar data procesada
    df = pd.read_csv(RUTA_OUTPUT_INVOLVES, sep=";", decimal=",", encoding="utf-8-sig")

    # 4. Filtrar registros con más de 540 minutos (9 horas)
    df['TIEMPO_NUM'] = pd.to_numeric(df['TIEMPO_SERVICIO_MIN'], errors='coerce').fillna(0)
    df_alertas = df[df['TIEMPO_NUM'] > 540].copy()

    # 5. Limpiar campos y guardar si hay alertas
    if not df_alertas.empty:
        columnas_a_quitar = [
            'JUSTIFICACION', 'JUSTIFICADO_EL', 'QUIEN_JUSTIFICO', 
            'MOTIVO_REAL', 'OBSERVACION_JUSTIFICACION', 'CEDULA', 
            'ROL', 'NOMBRE_VAL', 'CEDULA_QUIEN_JUSTIFICO'
        ]
        
        # Eliminar columnas no deseadas
        df_alertas.drop(columns=[c for c in columnas_a_quitar if c in df_alertas.columns], inplace=True)
        
        # Eliminar columna auxiliar de cálculo
        if 'TIEMPO_NUM' in df_alertas.columns:
            df_alertas.drop(columns=['TIEMPO_NUM'], inplace=True)

        # 6. Crear carpeta de salida si no existe y guardar
        try:
            os.makedirs(os.path.dirname(RUTA_ALERTA_CIF), exist_ok=True)
            df_alertas.to_csv(RUTA_ALERTA_CIF, index=False, sep=";", decimal=",", encoding="utf-8-sig")
            
            print(f"⚠️ AUDITORÍA: Se detectaron {len(df_alertas)} registros con tiempos > 9h.")
            print(f"📁 Archivo de alertas generado en: {RUTA_ALERTA_CIF}")
        except Exception as e:
            print(f"❌ ERROR al guardar alertas: {e}")
    else:
        print("✅ No se detectaron tiempos superiores a 9 horas.")

    print("="*40)


# ─────────────────────────────────────────────
# BLOQUE 7: ARCHIVO FINAL SALIDA CIF
# ─────────────────────────────────────────────

import pandas as pd
import os

def ejecutar_paso_7(spec: pr.PeriodoSpec):
    print("\n" + "="*40)
    print(f" INICIANDO PASO 7: ESTRUCTURACIÓN FINAL  ({spec.etiqueta})")
    print("="*40)

    # Sprint 15.5.3 — Fix F2: el mes/año son del periodo solicitado, no
    # detectados por mode/max sobre las fechas del archivo.
    val_mes  = spec.mes
    val_anio = spec.anio

    # Sanity check: si el archivo trae fechas mayoritariamente de OTRO mes,
    # avisar — puede indicar que las BASES no corresponden al periodo pedido.
    try:
        df_ref = pd.read_csv(
            RUTA_PLAN_ORIGINAL, sep=";", decimal=",",
            encoding="utf-8-sig", usecols=['FECHA'],
        )
        df_ref['FECHA_DT'] = pd.to_datetime(df_ref['FECHA'], dayfirst=True, errors='coerce')
        df_ref = df_ref.dropna(subset=['FECHA_DT'])
        if not df_ref.empty:
            mes_archivo  = int(df_ref['FECHA_DT'].dt.month.mode()[0])
            anio_archivo = int(df_ref['FECHA_DT'].dt.year.mode()[0])
            if (mes_archivo, anio_archivo) != (val_mes, val_anio):
                print(
                    f"⚠️ ALERTA DE CONSISTENCIA: el PT consolidado tiene fechas "
                    f"predominantes de {mes_archivo:02d}/{anio_archivo}, pero "
                    f"se está procesando como {val_mes:02d}/{val_anio}. "
                    f"Verifica que las BASES correspondan al periodo solicitado."
                )
            else:
                print(f"📅 Periodo confirmado: {val_mes:02d}/{val_anio}")
    except Exception as e:
        print(f"⚠️ No pude validar fechas del archivo: {e}. Continúo con {val_mes:02d}/{val_anio}.")

    # 3. Carga y Cruce
    df_plan = pd.read_csv(RUTA_PLAN_AGRUPADO, sep=";", decimal=",", encoding="utf-8-sig")
    df_visitas = pd.read_csv(RUTA_AGRUPADO_VISITAS, sep=";", decimal=",", encoding="utf-8-sig")

    # Limpieza de Llaves numéricas
    for df_temp, col in [(df_visitas, 'ID_PDV_INVOLVES'), (df_visitas, 'USUARIO_INVOLVES'),
                         (df_plan, 'ID_PDV_INVOLVES'), (df_plan, 'CEDULA')]:
        if col in df_temp.columns:
            df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0).astype(int).astype(str).str.strip()

    # ── Sprint 15.9 (Fix H4 — opción C) — NOMBRE primario + CEDULA fallback ──
    # Problema raíz: las CEDULAS del plan no siempre coinciden con USUARIO de
    # Colaboradores. En particular hay casos de cédulas LITERALMENTE
    # INTERCAMBIADAS entre dos personas (NICOLL ↔ CAMILA), lo que hace que
    # el cruce por CEDULA asigne visitas cruzadas. También nombres con
    # espacios múltiples (ANGELICA REYES  MORENO) fallaban en el match.
    #
    # Estrategia:
    #   1. NOMBRE primario — normalizado quitando espacios múltiples,
    #      es la llave más estable porque "EMPLEADO" en Involves se
    #      escribió por la misma persona en Colaboradores.
    #   2. CEDULA fallback — para los casos donde el NOMBRE difiere entre
    #      plan y visitas (raro, pero posible por tildes o cambios menores).
    import re as _re
    def _norm_nombre(s):
        return _re.sub(r'\s+', ' ', str(s)).strip().upper()

    df_plan['_NOMBRE_KEY']    = df_plan['NOMBRE'].map(_norm_nombre)
    df_visitas['_NOMBRE_KEY'] = df_visitas['EMPLEADO'].map(_norm_nombre)

    # 1) Consolidar visitas por (PDV, NOMBRE_KEY) sumando TIEMPO y VISITAS.
    # Sprint 15.10 — Fix bug duplicados Paso 5:
    # El Paso 5 agrupa por (PDV, EMPLEADO, PERFIL_UNIFICADO). Las visitas
    # justificadas ("Sin check-in") no tienen PERFIL_ACCESO, lo que genera
    # filas separadas de las realizadas para la misma persona. Con un simple
    # drop_duplicates(first), pandas ordena alfabéticamente y retiene la fila
    # vacía (PERFIL="") con TIEMPO=0, descartando las realizadas (PERFIL=GESTOR).
    # → 226 casos afectados, ~1,587 hrs perdidas en Mayo 2026.
    # Solución: groupby + sum, así realizadas + justificadas válidas suman
    # correctamente y las inválidas (TIEMPO=0) no afectan el total.
    def _first_non_null(s):
        s = s.dropna()
        return s.iloc[0] if len(s) else None
    df_visitas_unicas_nombre = df_visitas.groupby(
        ['ID_PDV_INVOLVES', '_NOMBRE_KEY'], as_index=False, dropna=False
    ).agg(
        PUNTO_DE_VENTA=('PUNTO_DE_VENTA', 'first'),
        USUARIO_INVOLVES=('USUARIO_INVOLVES', _first_non_null),
        PERFIL_ACCESO=('PERFIL_ACCESO', _first_non_null),
        SUMA_TIEMPO_SERVICIO_MIN=('SUMA_TIEMPO_SERVICIO_MIN', 'sum'),
        CONTEO_VISITAS_OK=('CONTEO_VISITAS_OK', 'sum'),
    )

    # Cruce primario por NOMBRE normalizado
    df_via_nombre = df_plan.merge(
        df_visitas_unicas_nombre,
        on=['ID_PDV_INVOLVES', '_NOMBRE_KEY'],
        how='left',
        suffixes=('', '_VIS'),
    )

    # 2) Filas del plan que NO matchearon — intentar por CEDULA
    cols_visitas_solo = [c for c in df_visitas.columns
                         if c not in ('ID_PDV_INVOLVES', '_NOMBRE_KEY')]
    mask_no_match = df_via_nombre['SUMA_TIEMPO_SERVICIO_MIN'].isna()

    if mask_no_match.any():
        # Sprint 15.9 (final fix) — usar df_via_nombre[mask_no_match] (no df_plan)
        # porque después del primer merge, df_via_nombre puede tener filas
        # adicionales si hubo duplicados (aunque ahora protegido con drop_duplicates).
        # También aseguramos unicidad de visitas por (PDV, CEDULA).
        # Sprint 15.10 — mismo fix que arriba pero por CEDULA (USUARIO_INVOLVES).
        df_visitas_unicas_cedula = df_visitas.groupby(
            ['ID_PDV_INVOLVES', 'USUARIO_INVOLVES'], as_index=False, dropna=False
        ).agg(
            PUNTO_DE_VENTA=('PUNTO_DE_VENTA', 'first'),
            EMPLEADO=('EMPLEADO', _first_non_null),
            _NOMBRE_KEY=('_NOMBRE_KEY', _first_non_null),
            PERFIL_ACCESO=('PERFIL_ACCESO', _first_non_null),
            SUMA_TIEMPO_SERVICIO_MIN=('SUMA_TIEMPO_SERVICIO_MIN', 'sum'),
            CONTEO_VISITAS_OK=('CONTEO_VISITAS_OK', 'sum'),
        )
        plan_rescate = df_via_nombre.loc[mask_no_match, df_plan.columns.tolist()].copy()
        rescate = plan_rescate.merge(
            df_visitas_unicas_cedula,
            left_on=['ID_PDV_INVOLVES', 'CEDULA'],
            right_on=['ID_PDV_INVOLVES', 'USUARIO_INVOLVES'],
            how='left',
            suffixes=('', '_VIS'),
        )
        # Re-inyectar: usar .reset_index porque los índices de df_via_nombre
        # y rescate son distintos.
        rescate_reset = rescate.reset_index(drop=True)
        mask_indices = df_via_nombre.index[mask_no_match]
        for col in cols_visitas_solo:
            if col in rescate_reset.columns:
                df_via_nombre.loc[mask_indices, col] = rescate_reset[col].values
        n_rescatadas = (~rescate_reset['SUMA_TIEMPO_SERVICIO_MIN'].isna()).sum()
        print(f"  🔧 H4 fallback: {n_rescatadas:,} filas del plan rescatadas via CEDULA.")

    df_final = df_via_nombre
    df_final.drop(columns=['_NOMBRE_KEY'], inplace=True, errors='ignore')

    # 4. Cálculos y Renombramiento
    df_final['SUMA_TIEMPO_SERVICIO_MIN'] = df_final['SUMA_TIEMPO_SERVICIO_MIN'].fillna(0)
    df_final['CONTEO_VISITAS_OK'] = df_final['CONTEO_VISITAS_OK'].fillna(0)
    df_final['SUMA_TIEMPO_SERVICIO_HORAS_REAL'] = df_final['SUMA_TIEMPO_SERVICIO_MIN'] / 60

    columnas_nuevas = {
        'TIEMPO_SERVICIO': 'TIEMPO_SERVICIO_PLANEADO',
        'TIEMPO_SERVICIO_HORAS': 'TIEMPO_SERVICIO_HORAS_PLANEADO',
        'FREC_SEMANAL': 'FREC_SEMANAL_PLANEADO',
        'FREC_MENSUAL': 'FREC_MENSUAL_PLANEADO',
        'HORAS_X_VISITA': 'HORAS_X_VISITA_PLANEADO',
        'HORAS_MES': 'HORAS_MES_PLANEADO',
        'SUMA_TIEMPO_SERVICIO_MIN': 'SUMA_TIEMPO_SERVICIO_MIN_REAL',
        'CONTEO_VISITAS_OK': 'VISITAS_REAL'
    }
    df_final.rename(columns=columnas_nuevas, inplace=True)

    # ── 4b. KPIs base por fila (V3 — Sprint 7) ────────────────────────
    # Estos KPIs los consume después generar_resumen_kpis_8() y también el
    # detector de anomalías (Fase 3). Por fila (PDV × persona).
    #
    # COBERTURA       = 1 si VISITAS_REAL >= 1, sino 0
    # INTENSIDAD      = horas_real / HORAS_MES_PLANEADO  (todos los roles)
    # INTENSIDAD GEST = horas_real / TIEMPO_SERVICIO_HORAS_PLANEADO  (solo gestores)
    # FRECUENCIA      = visitas_real / (FREC_SEMANAL × FREC_MENSUAL)  (todos)
    # FRECUENCIA GEST = visitas_real / CANTIDAD_VISITAS  (solo gestores)
    df_final['COBERTURA'] = (df_final['VISITAS_REAL'] >= 1).astype(int)

    df_final['INTENSIDAD'] = np.where(
        df_final['HORAS_MES_PLANEADO'].astype(float) > 0,
        df_final['SUMA_TIEMPO_SERVICIO_HORAS_REAL'] / df_final['HORAS_MES_PLANEADO'],
        0.0,
    )

    mask_gestor = df_final['ROL'].astype(str).str.upper().str.contains('GESTOR', na=False)
    df_final['INTENSIDAD GESTORES'] = 0.0
    df_final.loc[mask_gestor, 'INTENSIDAD GESTORES'] = np.where(
        df_final.loc[mask_gestor, 'TIEMPO_SERVICIO_HORAS_PLANEADO'].astype(float) > 0,
        df_final.loc[mask_gestor, 'SUMA_TIEMPO_SERVICIO_HORAS_REAL']
            / df_final.loc[mask_gestor, 'TIEMPO_SERVICIO_HORAS_PLANEADO'],
        0.0,
    )

    den_frec = (
        df_final['FREC_SEMANAL_PLANEADO'].astype(float)
        * df_final['FREC_MENSUAL_PLANEADO'].astype(float)
    )
    df_final['FRECUENCIA'] = np.where(
        den_frec > 0, df_final['VISITAS_REAL'] / den_frec, 0.0
    )

    df_final['FRECUENCIA GESTORES'] = 0.0
    df_final.loc[mask_gestor, 'FRECUENCIA GESTORES'] = np.where(
        df_final.loc[mask_gestor, 'CANTIDAD_VISITAS'].astype(float) > 0,
        df_final.loc[mask_gestor, 'VISITAS_REAL']
            / df_final.loc[mask_gestor, 'CANTIDAD_VISITAS'],
        0.0,
    )

    # Asignar valores numéricos capturados
    df_final['MES'] = val_mes
    df_final['AÑO'] = val_anio

    # 5. Guardado Final en XLSX
    columnas_ordenadas = [
        'ID_PDV_INVOLVES', 'ACRONIMO', 'NOMBRE_PDV', 'VENTAS_PROMEDIO_MES',
        'CEDULA', 'NOMBRE', 'COD_MERCADERISTA', 'CANTIDAD_VISITAS',
        'TIEMPO_SERVICIO_PLANEADO', 'TIEMPO_SERVICIO_HORAS_PLANEADO',
        'FREC_SEMANAL_PLANEADO', 'FREC_MENSUAL_PLANEADO',
        'HORAS_X_VISITA_PLANEADO', 'HORAS_MES_PLANEADO',
        'SUMA_TIEMPO_SERVICIO_MIN_REAL', 'SUMA_TIEMPO_SERVICIO_HORAS_REAL',
        'VISITAS_REAL',
        # KPIs V3 (Sprint 7)
        'COBERTURA', 'INTENSIDAD', 'INTENSIDAD GESTORES',
        'FRECUENCIA', 'FRECUENCIA GESTORES',
        'ROL', 'SUPERVISOR_LIDER', 'FUENTE', 'MES', 'AÑO',
    ]

    df_resultado = df_final[[c for c in columnas_ordenadas if c in df_final.columns]].copy()

    # D8=B: escribir AMBOS archivos para mantener compatibilidad
    os.makedirs(os.path.dirname(RUTA_FINAL_CIF_EXCEL), exist_ok=True)

    # Sprint 17.20 — upsert por (MES, AÑO): preservar periodos previos
    # para que GOLD/BI mantengan el detalle multi-periodo cuando se corre
    # solo un mes. Antes el archivo se sobrescribia completo.
    if os.path.exists(str(paths.CIF_OUT_CIF)):
        try:
            df_prev = pd.read_excel(str(paths.CIF_OUT_CIF))
            if "MES" in df_prev.columns and "AÑO" in df_prev.columns:
                periodos_actuales = set(
                    (int(m), int(a)) for m, a in zip(df_resultado["MES"], df_resultado["AÑO"])
                    if pd.notna(m) and pd.notna(a)
                )
                df_prev["_mes_int"]  = pd.to_numeric(df_prev["MES"],  errors="coerce")
                df_prev["_anio_int"] = pd.to_numeric(df_prev["AÑO"], errors="coerce")
                mask_keep = ~df_prev.apply(
                    lambda r: (r["_mes_int"], r["_anio_int"]) in periodos_actuales,
                    axis=1,
                )
                df_prev = df_prev[mask_keep].drop(columns=["_mes_int", "_anio_int"])
                df_resultado = pd.concat([df_prev, df_resultado], ignore_index=True, sort=False)
            else:
                print("⚠️  CIF.xlsx previo sin MES/AÑO — descartado en upsert.")
        except Exception as e:
            print(f"⚠️  No se pudo leer CIF.xlsx previo ({e}); se sobrescribe.")

    df_resultado.to_excel(str(paths.CIF_OUT_CIF), index=False, engine='openpyxl')
    df_resultado.to_excel(str(paths.CIF_OUT_FINAL), index=False, engine='openpyxl')

    print(f"✅ ¡LISTO! Excel generado para el periodo {val_mes}/{val_anio}.")
    print(f"📁 {paths.CIF_OUT_CIF}")
    print(f"📁 {paths.CIF_OUT_FINAL}  (alias legacy)")
    print("="*40)


# ─────────────────────────────────────────────
# BLOQUE 8 (V3): RESUMEN KPIS POR GESTOR
# ─────────────────────────────────────────────
# Lee CIF.xlsx (detalle PDV×persona), agrupa por (MES, AÑO, NOMBRE),
# calcula CUMPLIMIENTO COBERTURA/INTENSIDAD/FRECUENCIA ponderado por
# VENTAS_PROMEDIO_MES, y produce:
#   • KPIS_CIF.xlsx              — solo el periodo más reciente
#   • KPIS_CIF_HISTORICO.xlsx    — acumulado con upsert por (MES,AÑO,NOMBRE)
# Fórmula final: TOTAL = 0.5·COB + 0.2·INT + 0.3·FREC  (cambio confirmado)

def generar_resumen_kpis_8(spec: pr.PeriodoSpec):
    print("\n" + "=" * 40)
    print(f" INICIANDO PASO 8: RESUMEN KPIS POR GESTOR (V3)  ({spec.etiqueta})")
    print("=" * 40)

    file_in = str(paths.CIF_OUT_CIF)
    if not os.path.exists(file_in):
        print(f"❌ No existe el detalle CIF: {file_in}")
        return

    print(f"  Leyendo: {file_in}")
    df = pd.read_excel(file_in, engine='openpyxl')

    # Sprint 17.20 — CIF.xlsx ahora hace upsert por (MES, AÑO) y puede
    # contener varios periodos. Filtramos al periodo solicitado en lugar
    # de exigir que solo este (la antigua validación se rompía tras el upsert).
    if 'MES' in df.columns and 'AÑO' in df.columns:
        m_int = pd.to_numeric(df['MES'], errors='coerce')
        a_int = pd.to_numeric(df['AÑO'], errors='coerce')
        mask_periodo = (m_int == spec.mes) & (a_int == spec.anio)
        if not mask_periodo.any():
            raise ValueError(
                f"Paso 8: el detalle CIF.xlsx no contiene filas para el "
                f"periodo {spec.mes}/{spec.anio}. Re-ejecuta el paso 7."
            )
        df = df[mask_periodo].copy()

    def _limpiar(serie: pd.Series) -> pd.Series:
        s = serie.astype(str).str.strip().str.replace('%', '').str.replace(',', '.')
        n = pd.to_numeric(s, errors='coerce')
        # Heurística: si los datos vienen como 80 en vez de 0.8, dividir por 100
        if n.dropna().mean() > 1.5:
            n = n / 100
        return n.fillna(0)

    df['VENTAS_PROMEDIO_MES'] = pd.to_numeric(
        df['VENTAS_PROMEDIO_MES'].astype(str).str.replace(',', '.'),
        errors='coerce',
    ).fillna(0)
    df['COBERTURA']  = _limpiar(df['COBERTURA'])
    df['INTENSIDAD'] = _limpiar(df['INTENSIDAD'])
    df['FRECUENCIA'] = _limpiar(df['FRECUENCIA'])
    df['MES']  = pd.to_numeric(df.get('MES'),  errors='coerce').fillna(0).astype(int)
    df['AÑO']  = pd.to_numeric(df.get('AÑO'),  errors='coerce').fillna(0).astype(int)

    # Productos ponderados por venta (numeradores)
    df['_prod_cob'] = df['VENTAS_PROMEDIO_MES'] * df['COBERTURA']
    df['_prod_int'] = df['VENTAS_PROMEDIO_MES'] * df['INTENSIDAD']
    df['_prod_fre'] = df['VENTAS_PROMEDIO_MES'] * df['FRECUENCIA']

    # Agrupar por (MES, AÑO, NOMBRE) — un gestor = una fila de cumplimiento.
    # Sprint 15.10 — Fix opción C: el groupby anterior incluía SUPERVISOR_LIDER,
    # lo que hacía que un gestor asignado a varios supervisores apareciera
    # múltiples veces. Caso JENNY × PDVs sin supervisor producía una fila
    # fantasma con TOTAL=0.50. SUPERVISOR_LIDER se mantiene como informativo
    # (moda, ignorando "0" si hay alternativa válida).
    def _supervisor_dominante(s):
        s = s.dropna().astype(str).str.strip()
        s = s[(s != '') & (s != '0')]
        if s.empty:
            return ''
        return s.mode().iloc[0]

    resumen = (
        df.groupby(['MES', 'AÑO', 'NOMBRE'], dropna=False)
          .agg(SUPERVISOR_LIDER=('SUPERVISOR_LIDER', _supervisor_dominante),
               VENTAS_PROMEDIO_MES=('VENTAS_PROMEDIO_MES', 'sum'),
               prod_cob=('_prod_cob', 'sum'),
               prod_int=('_prod_int', 'sum'),
               prod_fre=('_prod_fre', 'sum'))
          .reset_index()
    )

    resumen['CUMPLIMIENTO COBERTURA']  = np.where(
        resumen['VENTAS_PROMEDIO_MES'] != 0,
        resumen['prod_cob'] / resumen['VENTAS_PROMEDIO_MES'], 0)
    resumen['CUMPLIMIENTO INTENSIDAD'] = np.where(
        resumen['VENTAS_PROMEDIO_MES'] != 0,
        resumen['prod_int'] / resumen['VENTAS_PROMEDIO_MES'], 0)
    resumen['CUMPLIMIENTO FRECUENCIA'] = np.where(
        resumen['VENTAS_PROMEDIO_MES'] != 0,
        resumen['prod_fre'] / resumen['VENTAS_PROMEDIO_MES'], 0)

    # Fórmula confirmada: 0.5·COB + 0.2·INT + 0.3·FREC
    resumen['TOTAL'] = (
        resumen['CUMPLIMIENTO COBERTURA']  * 0.5
        + resumen['CUMPLIMIENTO INTENSIDAD'] * 0.2
        + resumen['CUMPLIMIENTO FRECUENCIA'] * 0.3
    )

    cols_final = ['MES', 'AÑO', 'SUPERVISOR_LIDER', 'NOMBRE',
                  'VENTAS_PROMEDIO_MES',
                  'CUMPLIMIENTO COBERTURA', 'CUMPLIMIENTO INTENSIDAD',
                  'CUMPLIMIENTO FRECUENCIA', 'TOTAL']
    resumen_final = resumen[cols_final].sort_values(
        ['AÑO', 'MES', 'SUPERVISOR_LIDER', 'NOMBRE']
    ).reset_index(drop=True)

    # ── Output 1: mes activo (último periodo presente en el archivo) ──
    if not resumen_final.empty:
        anio_max = int(resumen_final['AÑO'].max())
        mes_max  = int(resumen_final[resumen_final['AÑO'] == anio_max]['MES'].max())
        del_periodo_activo = (
            (resumen_final['AÑO'] == anio_max) & (resumen_final['MES'] == mes_max)
        )
        resumen_activo = resumen_final[del_periodo_activo].copy()
    else:
        anio_max, mes_max = 0, 0
        resumen_activo = resumen_final.copy()

    os.makedirs(os.path.dirname(str(paths.CIF_OUT_KPIS)), exist_ok=True)
    resumen_activo.to_excel(str(paths.CIF_OUT_KPIS), index=False, engine='openpyxl')
    print(f"  ✅ Mes activo  ({mes_max:02d}/{anio_max}): "
          f"{len(resumen_activo)} gestores → {paths.CIF_OUT_KPIS.name}")

    # ── Output 2: histórico acumulado con upsert por (MES, AÑO, NOMBRE) ──
    ruta_hist = str(paths.CIF_OUT_KPIS_HISTORICO)
    if os.path.exists(ruta_hist):
        try:
            hist_prev = pd.read_excel(ruta_hist, engine='openpyxl')
            hist_prev['MES'] = pd.to_numeric(hist_prev.get('MES'), errors='coerce').fillna(0).astype(int)
            hist_prev['AÑO'] = pd.to_numeric(hist_prev.get('AÑO'), errors='coerce').fillna(0).astype(int)
        except Exception:
            hist_prev = pd.DataFrame(columns=cols_final)
    else:
        hist_prev = pd.DataFrame(columns=cols_final)

    # Sprint 15.9 (final fix) — upsert por PERIODO completo (MES, AÑO),
    # no por (MES, AÑO, NOMBRE). Si una persona desaparece del run nuevo
    # (ej. el filtro la excluye), su fila vieja debe limpiarse del histórico.
    if not hist_prev.empty and not resumen_final.empty:
        periodos_nuevos = set(map(tuple,
            resumen_final[['MES', 'AÑO']].drop_duplicates().values.tolist()))
        mask_keep = ~hist_prev.apply(
            lambda r: (int(r['MES']), int(r['AÑO'])) in periodos_nuevos,
            axis=1,
        )
        hist_prev = hist_prev[mask_keep]

    hist_final = pd.concat([hist_prev, resumen_final], ignore_index=True)
    hist_final = hist_final.sort_values(['AÑO', 'MES', 'SUPERVISOR_LIDER', 'NOMBRE'])
    hist_final.to_excel(ruta_hist, index=False, engine='openpyxl')
    print(f"  ✅ Histórico: {len(hist_final)} filas totales → {paths.CIF_OUT_KPIS_HISTORICO.name}")
    print("=" * 40)

# ─────────────────────────────────────────────
# CONTROL DE EJECUCIÓN PARA PRUEBAS
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="ETL CIF — Eficacia (Sprint 16.1: multi-periodo)")
    parser.add_argument("--solo", nargs="+",
                        choices=["consolidacion", "agrupacion", "visitas", "visitas2",
                                  "agrupacion_visitas", "alerta", "final", "kpis"],
                        help="Correr solo los pasos indicados (default: final + kpis)")
    parser.add_argument("--full", action="store_true",
                        help="Correr la cadena completa (paso 1 → paso 8)")
    # Sprint 16.1 — D1 multi-periodo: --mes/--anio para un mes, --periodos para varios.
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()

    specs = pr.periodos_de_args(args)

    if args.full:
        pasos = ["consolidacion", "agrupacion", "visitas", "visitas2",
                  "agrupacion_visitas", "alerta", "final", "kpis"]
    else:
        pasos = args.solo or ["final", "kpis"]

    if len(specs) > 1:
        print(f"\n🎯 ETL CIF — multi-periodo: {len(specs)} meses → "
              f"{', '.join(s.etiqueta for s in specs)}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL CIF — procesando periodo {spec.etiqueta} ({spec})")

        try:
            if "consolidacion" in pasos:        ejecutar_paso_1(spec)
            if "agrupacion" in pasos:           ejecutar_paso_2()
            if "visitas" in pasos:              ejecutar_paso_3(spec)
            if "visitas2" in pasos:             ejecutar_paso_4(spec)
            if "agrupacion_visitas" in pasos:   ejecutar_paso_5()
            if "alerta" in pasos:               ejecutar_paso_6()
            if "final" in pasos:                ejecutar_paso_7(spec)
            if "kpis" in pasos:                 generar_resumen_kpis_8(spec)

            print(f"\n🏁 PROCESO TERMINADO  ({spec.etiqueta}).")

        except Exception as e:
            print(f"\n❌ OCURRIÓ UN ERROR CRÍTICO en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise