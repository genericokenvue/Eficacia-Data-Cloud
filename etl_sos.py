import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

import pandas as pd
import os
import glob
import re
import unicodedata
from datetime import datetime

# =============================================================================
# 1. CONFIGURACIÓN DE RUTAS REALES Y DICCIONARIOS
# =============================================================================

MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# --- RUTAS resueltas vía paths.py ---
import paths
import periodo_resolver as pr

_DIR_PT = str(paths.CIF_PT_DIR)

RUTA_ORIGEN_SOS = str(paths.SOS_BASES)
RUTA_SALIDA_SOS = str(paths.SOS_SALIDA)
os.makedirs(RUTA_SALIDA_SOS, exist_ok=True)

# Patrón relajado (matchea "Respuestas de encuestas..." y "Respuestas SOS"/"Respuestas SOS Baby")
CLAVE_ARCHIVO_NP = "Respuestas"
CLAVE_ARCHIVO_TARGET = "Colombia-sos-target"

# --- NOMBRES DE ARCHIVOS DE PROCESO ---
NOMBRE_SALIDA_ENCUESTA = "Encuesta Sos Consolidada.xlsx"
RUTA_PT_CONSOLIDADO_FINAL = os.path.join(RUTA_ORIGEN_SOS, "Plan_Trabajo_SOS_Completo.xlsx")
RUTA_ENCUESTA_BASE = os.path.join(RUTA_ORIGEN_SOS, NOMBRE_SALIDA_ENCUESTA)
RUTA_TARGET_SALIDA_BASE = os.path.join(RUTA_ORIGEN_SOS, "Colombia_sos_target_Normalizado.xlsx")

# --- ARCHIVO DE SALIDA FINAL ---
RUTA_REPORTE_FINAL_SOS = os.path.join(RUTA_SALIDA_SOS, "Reporte_SOS_Final_Calculado.xlsx")

# --- CONFIGURACIÓN DE CUMPLIMIENTO (17 CATEGORÍAS ÚNICAS) ---
LISTA_17_CATEGORIAS = [
    "SHAMPOO BEBE EN ADULTOS", "SHAMPOO BEBE", "JABONES SOLIDOS BEBE", 
    "JABONES LIQUIDOS BEBE", "CREMAS CORPORALES BEBE", "ASEO DEL BEBE", 
    "TOALLAS", "TAMPONES", "PROTECTORES", "PROTECCION SOLAR", 
    "JABONES SOLIDOS ADULTOS", "JABONES LIQUIDOS ADULTOS", "FACIALES", 
    "ENJUAGUES BUCALES MASIVOS", "ENJUAGUES BUCALES ESPECIALIZADOS", 
    "ENJUAGUES BUCALES", "CREMAS CORPORALES ADULTO"
]

# =============================================================================
# 2. FUNCIONES DE APOYO
# =============================================================================

def eliminar_tildes(texto):
    if pd.isna(texto): return texto
    texto = str(texto)
    s = ''.join(c for c in unicodedata.normalize('NFD', texto)
                if unicodedata.category(c) != 'Mn')
    return s.upper().strip()

# =============================================================================
# 3. FUNCIONES DE PROCESAMIENTO
# =============================================================================

def ejecutar_paso_1_consolidar_pt(spec: pr.PeriodoSpec):
    print(f"\n--- PASO 1: Consolidando Plan de Trabajo  ({spec.etiqueta}) ---")
    from shared_loader import COLUMNAS_ESTANDAR_UNIFICADO as COLUMNAS_ESTANDAR

    # Sprint 15.9 — incluir OBSERVACIÓN para filtrar por responsable real.
    COLUMNAS_FINALES_PT = ["ID_PDV_INVOLVES", "NOMBRE_PDV", "VENTAS_PROMEDIO_MES", "ACRONIMO",
                           "CEDULA", "NOMBRE", "COD_MERCADERISTA", "ROL", "SUPERVISOR_LIDER",
                           "FUENTE", "OBSERVACION"]
    OBS_A_ROL = {
        "REPORTA GESTOR":              "GESTOR",
        "REPORTA SUPERVISOR":          "SUPERVISOR",
        "REPORTA GENERADOR DE DEMANDA":"GENERADOR DE DEMANDA",
    }

    def leer_y_normalizar(ruta, hoja, fuente):
        if not os.path.exists(ruta): return pd.DataFrame()
        with pd.ExcelFile(ruta) as xls:
            df_val = pd.read_excel(xls, sheet_name="Captura de modulos")
            df_val.columns = df_val.columns.str.strip().str.upper()
            col_filtro = "ESPACIOS_FINAL" if fuente == "ISM" else "ESPACIOS"
            obs_col = next((c for c in df_val.columns if c.startswith('OBSERVAC')), None)
            if obs_col is None:
                print(f"⚠ {fuente}: hoja Captura de modulos sin OBSERVACIÓN")
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
            df = df[df["ID PDV INVOLVES"].isin(pdvs_obs["ID PDV INVOLVES"])].copy()
            df.rename(columns={col: COLUMNAS_ESTANDAR[col] for col in df.columns if col in COLUMNAS_ESTANDAR}, inplace=True)
            df = df.merge(pdvs_obs, left_on="ID_PDV_INVOLVES", right_on="ID PDV INVOLVES", how="left")
            df["ROL"] = df["ROL"].astype(str).str.strip().str.upper()
            n_antes = len(df)
            df = df[df["ROL"] == df["ROL_ESPERADO"]].copy()
            print(f"  {fuente}: filtro OBSERVACIÓN → {n_antes} → {len(df)} filas (responsables reales)")
            df["FUENTE"] = fuente
            for col in COLUMNAS_FINALES_PT:
                if col not in df.columns: df[col] = ""
            return df[COLUMNAS_FINALES_PT]

    from shared_loader import descubrir_archivos_pt
    ruta_directo, ruta_ism = descubrir_archivos_pt(_DIR_PT, periodo=spec)
    if not ruta_directo or not ruta_ism:
        raise FileNotFoundError(
            f"SOS ({spec.etiqueta}): faltan PT del periodo. "
            f"Directo={ruta_directo or 'N/D'} | ISM={ruta_ism or 'N/D'}"
        )
    df_dir = leer_y_normalizar(ruta_directo, "Plan de trabajo", "DIRECTO")
    df_ism = leer_y_normalizar(ruta_ism, "Plan de trabajo CIF", "ISM")
    if not df_dir.empty or not df_ism.empty:
        df_total = pd.concat([df_dir, df_ism], ignore_index=True)
        df_total['ID_PDV_INVOLVES'] = df_total['ID_PDV_INVOLVES'].astype(str).str.strip()

        # Sprint 15.7 (Fix H2) — D7: mantener una fila por (PDV, persona).
        # Antes: groupby('ID_PDV_INVOLVES').first() descartaba todas las
        # personas menos una por PDV — gestores quedaban fuera de SOS
        # cuando compartían PDV con su líder o supervisor.
        df_total = df_total.drop_duplicates(
            subset=['ID_PDV_INVOLVES', 'NOMBRE'],
        ).reset_index(drop=True)

        df_total.to_excel(RUTA_PT_CONSOLIDADO_FINAL, index=False)
        print(
            f"✅ EXITO: Plan unificado guardado "
            f"({len(df_total)} filas, {df_total['ID_PDV_INVOLVES'].nunique()} PDVs únicos)."
        )

def ejecutar_paso_2_consolidar_encuestas(spec: pr.PeriodoSpec):
    print(f"\n--- PASO 2: Consolidando Encuestas SOS  ({spec.etiqueta}) ---")
    # Sprint 15.5.6: encuestas del periodo solicitado (Normal + Baby), no glob.
    archivos = []
    ruta_normal = pr.sos_respuestas(spec)
    archivos.append(ruta_normal)
    print(f"  ✓ {ruta_normal.name}")
    try:
        ruta_baby = pr.sos_respuestas_baby(spec)
        archivos.append(ruta_baby)
        print(f"  ✓ {ruta_baby.name}")
    except FileNotFoundError:
        # Baby es opcional: no todos los meses la tienen.
        print(f"  (sin Respuestas SOS Baby para {spec.etiqueta})")
    lista_dfs = [pd.read_excel(a) for a in archivos]
    pd.concat(lista_dfs, ignore_index=True).to_excel(RUTA_ENCUESTA_BASE, index=False)
    print(f"✅ EXITO: Encuestas unidas ({len(archivos)} archivo(s)).")

def ejecutar_paso_3_cumplimiento_captura(spec: pr.PeriodoSpec):
    print(f"\n--- PASO 3: Generando Reporte de Cumplimiento  ({spec.etiqueta}) ---")
    if not os.path.exists(RUTA_PT_CONSOLIDADO_FINAL) or not os.path.exists(RUTA_ENCUESTA_BASE): return
    
    df_pt = pd.read_excel(RUTA_PT_CONSOLIDADO_FINAL)
    df_enc = pd.read_excel(RUTA_ENCUESTA_BASE)

    def extraer_categoria_estricta(texto):
        if pd.isna(texto): return None
        t = re.sub(r'[^A-Z0-9\s]', ' ', str(texto).upper())
        t = " ".join(t.split())
        if "SHAMPOO BEBE EN ADULTOS" in t: return "SHAMPOO BEBE EN ADULTOS"
        if re.search(r'\bSHAMPOO BEBE\b', t): return "SHAMPOO BEBE"
        for cat in LISTA_17_CATEGORIAS:
            if cat not in ["SHAMPOO BEBE", "SHAMPOO BEBE EN ADULTOS"] and cat in t: return cat
        return None

    # Procesamiento de encuestas
    df_enc['CATEGORIA_LIMPIA'] = df_enc['Rótulo de la encuesta'].apply(extraer_categoria_estricta)
    df_enc_valida = df_enc.dropna(subset=['CATEGORIA_LIMPIA']).copy()

    # Normalizar ambos IDs con el helper canónico (evita trailing ".0")
    from shared_loader import id_a_str
    df_pt['ID_PDV_INVOLVES'] = id_a_str(df_pt['ID_PDV_INVOLVES'])
    df_enc_valida['ID del PDV'] = id_a_str(df_enc_valida['ID del PDV'])

    # Cálculo de cumplimiento por PDV
    dict_encuesta = df_enc_valida.groupby('ID del PDV')['CATEGORIA_LIMPIA'].unique().to_dict()
    set_oficial = set(LISTA_17_CATEGORIAS)
    res = []
    
    for id_pdv in df_pt['ID_PDV_INVOLVES'].unique():
        cats = dict_encuesta.get(id_pdv, [])
        faltantes = sorted(list(set_oficial - set(cats))) if len(cats) < 17 else []
        res.append({
            'ID_PDV_INVOLVES': id_pdv, 
            'CONTEO_CATEGORIAS': len(cats), 
            'CATEGORIAS_FALTANTES': ", ".join(faltantes)
        })
    
    # Unión de datos
    df_final = pd.merge(df_pt, pd.DataFrame(res), on='ID_PDV_INVOLVES', how='left')
    
    # --- AGREGAR CAPTURA PLANEADA DESPUÉS DE FUENTE ---
    if 'FUENTE' in df_final.columns:
        idx_fuente = df_final.columns.get_loc('FUENTE')
        df_final.insert(idx_fuente + 1, 'CAPTURA_PLANEADA', 1)
    else:
        df_final['CAPTURA_PLANEADA'] = 1

    # Cálculo de indicadores
    df_final['CAPTURA_EJECUTADA'] = df_final['CONTEO_CATEGORIAS'].apply(lambda x: 1 if x == 17 else 0)
    df_final.loc[df_final['CONTEO_CATEGORIAS'] == 0, 'CATEGORIAS_FALTANTES'] = "SIN CAPTURAS (FALTAN 17)"

    # Sprint 15.5.6: propagar MES/AÑO al reporte intermedio (mismo Fix F3 que Precios).
    df_final['MES'] = spec.mes
    df_final['AÑO'] = spec.anio

    # --- REORDENAR PARA QUE LAS FECHAS QUEDEN AL FINAL ---
    cols_fecha = ["MES_AÑO", "PLAN_DE_TRABAJO"]
    # Filtramos columnas que no sean las de fecha para ponerlas primero
    otras_cols = [c for c in df_final.columns if c not in cols_fecha]
    # Reconstruimos la lista de columnas asegurando que las de fecha existan y vayan al final
    orden_final = otras_cols + [c for c in cols_fecha if c in df_final.columns]

    df_final = df_final[orden_final]

    df_final.to_excel(os.path.join(RUTA_SALIDA_SOS, "Cumplimiento_Captura_SOS.xlsx"), index=False)
    print(f"✅ EXITO: Reporte de capturas generado con Captura Planeada, Fechas y periodo {spec.etiqueta}.")

def ejecutar_paso_4_normalizar_target_dinamico(spec: pr.PeriodoSpec):
    print(f"\n--- PASO 4: Normalizando Target  ({spec.etiqueta}) ---")
    # Sprint 15.5.6: target del periodo solicitado.
    ruta_target_actual = str(pr.sos_target(spec))
    print(f"📂 Archivo detectado: {os.path.basename(ruta_target_actual)}")

    df_target = pd.read_excel(ruta_target_actual)
    df_target.columns = [str(c).strip().upper() for c in df_target.columns]
    
    col_id_target = "ID INVOLVES"
    if col_id_target not in df_target.columns:
        print(f"❌ Error: No encontré '{col_id_target}'.")
        return

    if not os.path.exists(RUTA_ENCUESTA_BASE):
        print("❌ Error: No existe la encuesta consolidada.")
        return
    df_encuesta = pd.read_excel(RUTA_ENCUESTA_BASE)

    col_cat = next((c for c in df_target.columns if 'CATEGOR' in c), None)
    if col_cat:
        df_target[col_cat] = df_target[col_cat].astype(str).str.upper().replace(
            "ENJUAGUES BUCALES TOTALES", "ENJUAGUES BUCALES"
        )
    
    for col in [col_cat, 'MARCA', 'NOMBRE DEL PDV']:
        if col in df_target.columns:
            df_target[col] = df_target[col].apply(eliminar_tildes)

    from shared_loader import id_a_str
    df_encuesta['ID del PDV'] = id_a_str(df_encuesta['ID del PDV'])
    df_target[col_id_target] = id_a_str(df_target[col_id_target])
    
    if 'Marca' in df_encuesta.columns:
        df_encuesta['Marca_Enc'] = df_encuesta['Marca'].apply(eliminar_tildes)
        ref_marcas = df_encuesta[['ID del PDV', 'Marca_Enc']].drop_duplicates().rename(columns={'Marca_Enc': 'Marca_Oficial'})
        df_target = pd.merge(df_target, ref_marcas, left_on=[col_id_target, 'MARCA'], 
                             right_on=['ID del PDV', 'Marca_Oficial'], how='left')
        df_target['MARCA'] = df_target['Marca_Oficial'].fillna(df_target['MARCA'])
        df_target.drop(columns=['ID del PDV', 'Marca_Oficial'], inplace=True, errors='ignore')

    df_target.to_excel(RUTA_TARGET_SALIDA_BASE, index=False)
    print(f"✅ EXITO: Target normalizado guardado en: {RUTA_TARGET_SALIDA_BASE}")

def ejecutar_paso_5_cruce_triple_y_calculo(spec: pr.PeriodoSpec):
    print(f"\n--- PASO 5: Generando Reporte Final (Cruce Triple y Cálculos)  ({spec.etiqueta}) ---")

    # Sprint 15.5.6 — Fix F4: short-circuit. Si fallaron pasos previos, abortar.
    # Antes el paso 6 (KPI) seguía corriendo sobre Cumplimiento_Captura_SOS.xlsx
    # viejo de Abril, produciendo datos falsamente fechados.
    faltantes = []
    if not os.path.exists(RUTA_TARGET_SALIDA_BASE): faltantes.append("Target Normalizado (paso 4)")
    if not os.path.exists(RUTA_ENCUESTA_BASE):      faltantes.append("Encuesta Consolidada (paso 2)")
    if faltantes:
        raise FileNotFoundError(
            f"Paso 5 SOS ({spec.etiqueta}): faltan archivos requeridos: "
            f"{faltantes}. Corre los pasos previos antes."
        )

    # Cargar bases
    df_target = pd.read_excel(RUTA_TARGET_SALIDA_BASE)
    df_enc = pd.read_excel(RUTA_ENCUESTA_BASE)

    # 1. Preparar Encuesta (Cruce Triple)
    def extraer_cat_cruce(texto):
        t = eliminar_tildes(texto)
        if "SHAMPOO BEBE EN ADULTOS" in t: return "SHAMPOO BEBE EN ADULTOS"
        if "SHAMPOO BEBE" in t: return "SHAMPOO BEBE"
        for c in LISTA_17_CATEGORIAS:
            if c in t: return c
        return t

    # Normalizamos los campos de cruce en la encuesta
    from shared_loader import id_a_str
    df_enc['CAT_CRUCE'] = df_enc['Rótulo de la encuesta'].apply(extraer_cat_cruce)
    df_enc['ID_CRUCE'] = id_a_str(df_enc['ID del PDV'])
    df_enc['MARCA_CRUCE'] = df_enc['Marca'].apply(eliminar_tildes)

    col_universo = "¿Cual es el universo en cms de la categoria?"
    col_marca_cms = "¿Cuántos cms tiene la marca?"
    
    # Agrupamos encuesta para evitar duplicados en el cruce
    # Usamos 'PDV' que es el nombre de la columna en tu archivo
    df_enc_resumen = df_enc.groupby(['ID_CRUCE', 'CAT_CRUCE', 'MARCA_CRUCE']).agg({
        col_universo: 'max',
        col_marca_cms: 'max',
        'PDV': 'first'
    }).reset_index()

    # 2. Preparar Target
    df_target['ID INVOLVES'] = id_a_str(df_target['ID INVOLVES'])
    col_cat_t = next((c for c in df_target.columns if 'CATEGOR' in c), None)

    # 3. REALIZAR EL CRUCE (Merge)
    df_final = pd.merge(
        df_target,
        df_enc_resumen,
        left_on=['ID INVOLVES', col_cat_t, 'MARCA'],
        right_on=['ID_CRUCE', 'CAT_CRUCE', 'MARCA_CRUCE'],
        how='left'
    )

    # 4. Cálculos de SOS
    df_final[col_universo] = pd.to_numeric(df_final[col_universo], errors='coerce').fillna(0)
    df_final[col_marca_cms] = pd.to_numeric(df_final[col_marca_cms], errors='coerce').fillna(0)

    # Participación SOS (Marca / Universo)
    df_final['PARTICIPACION_SOS'] = df_final.apply(
        lambda x: x[col_marca_cms] / x[col_universo] if x[col_universo] > 0 else 0, axis=1
    )

    # 5. REORDENAMIENTO Y FILTRADO FINAL (Solo los campos solicitados)
    # Renombramos 'PDV' a 'PUNTO DE VENTA' y la columna de categoría del target a 'CATEGORÍA DE PRODUCTO'
    df_final = df_final.rename(columns={
        'PDV': 'PUNTO DE VENTA',
        col_cat_t: 'CATEGORÍA DE PRODUCTO'
    })

    # Lista de columnas en el orden exacto solicitado
    columnas_ordenadas = [
        "ID INVOLVES", 
        "PUNTO DE VENTA", 
        "CATEGORÍA DE PRODUCTO", 
        "MARCA", 
        "SUBCANAL", 
        "TARGET", 
        "¿Cual es el universo en cms de la categoria?", 
        "¿Cuántos cms tiene la marca?", 
        "PARTICIPACION_SOS"
    ]

    # Validar que todas las columnas existan antes de filtrar para evitar errores
    columnas_existentes = [c for c in columnas_ordenadas if c in df_final.columns]
    df_final = df_final[columnas_existentes].copy()

    # Sprint 16.3 — propagar MES/AÑO al detalle para multi-periodo y BI.
    df_final["MES"] = spec.mes
    df_final["AÑO"] = spec.anio

    # Sprint 17.21 — el DataFrame tiene columnas con nombre duplicado
    # ("PUNTO DE VENTA" aparece 2 veces). pd.read_excel auto-renombra a
    # "PUNTO DE VENTA.1" al leer, pero df_final viene con el duplicado
    # crudo, lo que rompe pd.concat con InvalidIndexError. Renombrar
    # duplicados con sufijo .N — replica la convencion de pandas.
    def _dedupe_columns(cols):
        seen = {}
        out = []
        for c in cols:
            if c in seen:
                seen[c] += 1
                out.append(f"{c}.{seen[c]}")
            else:
                seen[c] = 0
                out.append(c)
        return out
    df_final.columns = _dedupe_columns(list(df_final.columns))

    # 6. Guardado Final — upsert por PERIODO sobre el archivo histórico.
    if os.path.exists(RUTA_REPORTE_FINAL_SOS):
        try:
            df_prev = pd.read_excel(RUTA_REPORTE_FINAL_SOS)
            # Defensive: el archivo tiene columnas con nombre duplicado
            # (e.g. PUNTO DE VENTA y PUNTO DE VENTA.1). pandas falla con
            # "Reindexing only valid with uniquely valued Index objects"
            # al filtrar. Quitar duplicados (mantener primera).
            df_prev = df_prev.loc[:, ~df_prev.columns.duplicated()]
            if "MES" in df_prev.columns and "AÑO" in df_prev.columns:
                mask_keep = ~(
                    (pd.to_numeric(df_prev["MES"], errors="coerce") == spec.mes)
                    & (pd.to_numeric(df_prev["AÑO"], errors="coerce") == spec.anio)
                )
                df_prev = df_prev[mask_keep]
            else:
                # Archivo legacy sin MES/AÑO — lo descartamos: no podemos saber
                # a qué periodo pertenece y mantenerlo contaminaría el detalle.
                print("⚠️  Reporte SOS previo no tenía MES/AÑO — descartado en upsert.")
                df_prev = pd.DataFrame(columns=df_final.columns)
            df_final = pd.concat([df_prev, df_final.reset_index(drop=True)], ignore_index=True, sort=False)
        except Exception as e:
            import traceback
            print(f"⚠️  Error en upsert SOS previo: {type(e).__name__}: {e}")
            traceback.print_exc()
            print(f"   df_prev.shape (si existe): {getattr(df_prev, 'shape', None)}")
            print(f"   df_final.shape: {df_final.shape}")
            print(f"   df_final.columns dup: {df_final.columns[df_final.columns.duplicated()].tolist()}")
            print(f"   df_final.index unique: {df_final.index.is_unique}")
            print(f"⚠️  PRESERVANDO archivo previo — NO se sobrescribe.")
            return   # NO escribir el archivo, evitar perder periodos previos

    df_final.to_excel(RUTA_REPORTE_FINAL_SOS, index=False)
    print(f"✅ EXITO: Reporte Final SOS ({spec.etiqueta}) generado en: {RUTA_REPORTE_FINAL_SOS}")

# =============================================================================
# 3.B PASO 6 (V3) — RESUMEN KPI por gestor
# =============================================================================
def generar_resumen_kpi_sos(spec: pr.PeriodoSpec):
    """Lee Cumplimiento_Captura_SOS.xlsx y produce SOS_KPIS.xlsx (mes activo)
    + SOS_KPIS_HISTORICO.xlsx (upsert acumulado). V3 spec.

    Sprint 15.5.6: valida que el cumplimiento corresponde al periodo solicitado
    (el paso 3 ya propaga MES/AÑO). Si no coincide, aborta — el archivo es de
    una corrida vieja.
    """
    print(f"\n--- PASO 6 (V3): RESUMEN KPI SOS por gestor  ({spec.etiqueta}) ---")
    ruta_cumpl = os.path.join(RUTA_SALIDA_SOS, "Cumplimiento_Captura_SOS.xlsx")
    if not os.path.exists(ruta_cumpl):
        print(f"❌ No existe: {ruta_cumpl} — corre primero ejecutar_paso_3_cumplimiento_captura()")
        return
    df = pd.read_excel(ruta_cumpl, engine='openpyxl')
    print(f"  Leyendo: {os.path.basename(ruta_cumpl)} ({len(df)} filas)")

    # Validar que el cumplimiento corresponde al periodo solicitado.
    if 'MES' in df.columns and 'AÑO' in df.columns:
        periodos = set(zip(
            pd.to_numeric(df['MES'], errors='coerce').dropna().astype(int),
            pd.to_numeric(df['AÑO'], errors='coerce').dropna().astype(int),
        ))
        if periodos and periodos != {(spec.mes, spec.anio)}:
            raise ValueError(
                f"SOS KPI: Cumplimiento_Captura_SOS.xlsx tiene periodos "
                f"{periodos} pero se pidió {(spec.mes, spec.anio)}. "
                f"El archivo es de una corrida vieja — re-ejecuta el paso 3."
            )
    else:
        # Archivo sin MES/AÑO → asignar al periodo solicitado (compat hacia atrás)
        df['MES'] = spec.mes
        df['AÑO'] = spec.anio

    from shared_loader import calcular_kpi_simple_y_escribir
    n = calcular_kpi_simple_y_escribir(
        df_origen=df,
        col_planeado="CAPTURA_PLANEADA",
        col_ejecutado="CAPTURA_EJECUTADA",
        ruta_kpi_mes=str(paths.SOS_OUT_KPIS),
        ruta_kpi_historico=str(paths.SOS_OUT_KPIS_HISTORICO),
        nombre_cumplimiento="CUMPLIMIENTO",
        nombres_renombre=("PLANEADO", "EJECUTADO"),
    )
    print(f"  ✅ Mes activo: {n} gestores → {paths.SOS_OUT_KPIS.name}")
    print(f"  ✅ Histórico acumulado → {paths.SOS_OUT_KPIS_HISTORICO.name}")


# =============================================================================
# 4. EJECUCIÓN PRINCIPAL
# =============================================================================

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL SOS — Eficacia (Sprint 16.1: multi-periodo)")
    parser.add_argument("--solo", nargs="+",
                        choices=["pt", "enc", "cumpl", "target", "final", "kpi"],
                        help="Correr solo los pasos indicados (default: final + kpi)")
    parser.add_argument("--full", action="store_true",
                        help="Correr la cadena completa (paso 1 → paso 6)")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    specs = pr.periodos_de_args(args)

    if args.full:
        pasos = ["pt", "enc", "cumpl", "target", "final", "kpi"]
    else:
        pasos = args.solo or ["final", "kpi"]

    if len(specs) > 1:
        print(f"🎯 ETL SOS — multi-periodo: {len(specs)} meses → "
              f"{', '.join(s.etiqueta for s in specs)}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL SOS — procesando periodo {spec.etiqueta} ({spec})")

        try:
            if "pt"     in pasos: ejecutar_paso_1_consolidar_pt(spec)
            if "enc"    in pasos: ejecutar_paso_2_consolidar_encuestas(spec)
            if "cumpl"  in pasos: ejecutar_paso_3_cumplimiento_captura(spec)
            if "target" in pasos: ejecutar_paso_4_normalizar_target_dinamico(spec)
            if "final"  in pasos: ejecutar_paso_5_cruce_triple_y_calculo(spec)
            if "kpi"    in pasos: generar_resumen_kpi_sos(spec)
            print(f"\n🏁 PROCESO TERMINADO  ({spec.etiqueta}).")
        except Exception as e:
            print(f"\n❌ OCURRIÓ UN ERROR CRÍTICO en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise