"""
aplicar_correcciones.py
───────────────────────
Lee los casos APROBADOS desde revision.sqlite y aplica las correcciones
a los archivos fuente de Involves / D&P. Cada modificación al source
queda registrada en Historial_Correcciones (auditoría completa).

Matriz decisión → acción
─────────────────────────
  Correcto / Ruta_ampliada / Desabastecimiento / Error_sistema
       → no modifica source; solo registra en Historial
  Doble_registro          → elimina la(s) fila(s) del source
  Error_digitacion        → sobreescribe el valor en source
  Cambio_real_PDV         → sobreescribe el valor en source

SOS especial
────────────
Cuando VALOR_CORRECTO viene en formato `universo=X|marca=Y`, el universo
se replica a TODAS las filas con el mismo ID_ENCUESTA en el source SOS;
la marca solo modifica la fila puntual.

Idempotencia
────────────
Casos ya en ESTADO=APLICADO se saltan. El campo APLICADO_EN registra
fecha/hora del primer apply.

Backups
───────
Antes de tocar cualquier archivo fuente, se hace una copia .bak con
timestamp en ALERTAS/REVISION/backups/. Si todo sale bien, se conserva
como punto de restauración.

Modo dry-run
────────────
  python aplicar_correcciones.py --dry-run   ← muestra qué haría sin
                                              modificar nada
"""

from __future__ import annotations

import os
import sys
import json
import shutil
import sqlite3
import argparse
import datetime as dt
from pathlib import Path
from typing import Any

# UTF-8 stdout en Windows
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import paths


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────
DIR_REV     = paths.ALERTAS_DIR / "REVISION"
DB_PATH     = DIR_REV / "revision.sqlite"
DIR_BACKUPS = DIR_REV / "backups"
DIR_BACKUPS.mkdir(parents=True, exist_ok=True)

# Decisiones que NO requieren modificar el source
DECISIONES_SIN_ACCION = {"Correcto", "Ruta_ampliada", "Desabastecimiento", "Error_sistema"}
DECISIONES_ELIMINAR  = {"Doble_registro"}
DECISIONES_SOBRESCRIBIR = {"Error_digitacion", "Cambio_real_PDV"}

# Sources por módulo (rutas relativas a BASES o SALIDA)
def _sources_por_modulo() -> dict[str, Path]:
    return {
        "SOS":     paths.SOS_BASES / "Encuesta Sos Consolidada.xlsx",
        "PRECIOS": _ultimo_archivo(paths.PR_SALIDA, "ANALISIS_PRECIOS_*.xlsx"),
        "NO_PRESENCIA": _ultimo_archivo(paths.NP_SALIDA, "ANALISIS_AGOTADOS_*.xlsx"),
        # CIF: las anomalías son agregados — no hay un campo individual a sobreescribir
        # por caso. Se registra en Historial pero no se modifica el archivo fuente.
        # Si en el futuro se quisiera, sería el informe-gerencial-visitas.xlsx.
    }


def _ultimo_archivo(directorio: Path, patron: str) -> Path:
    candidatos = sorted(directorio.glob(patron))
    return candidatos[-1] if candidatos else directorio / patron


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────────────────

def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(str(DB_PATH))
    c.row_factory = sqlite3.Row
    return c


def _backup(ruta: Path) -> Path:
    """Copia ruta a backups/<nombre>_<ts>.bak.<ext>. Devuelve la ruta del backup."""
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"{ruta.stem}_{ts}.bak{ruta.suffix}"
    destino = DIR_BACKUPS / nombre
    shutil.copy2(ruta, destino)
    return destino


def _parse_valor_correcto_sos(s: str) -> dict[str, Any]:
    """`universo=80|marca=25` → {'universo': 80.0, 'marca': 25.0}."""
    if not s:
        return {}
    out = {}
    for par in str(s).split("|"):
        if "=" not in par:
            continue
        k, v = par.split("=", 1)
        k = k.strip().lower()
        try:
            out[k] = float(v.strip())
        except (ValueError, TypeError):
            out[k] = v.strip()
    return out


def _parse_metadata(s: str) -> dict[str, Any]:
    if not s:
        return {}
    try:
        return json.loads(s)
    except (ValueError, TypeError):
        return {}


# ─────────────────────────────────────────────────────────────────────────────
# APLICADORES POR MÓDULO
# ─────────────────────────────────────────────────────────────────────────────

def aplicar_sos(caso: dict, source: Path, dry_run: bool) -> tuple[int, str]:
    """
    Aplica la corrección al archivo crudo de SOS.

    Si VALOR_CORRECTO viene como `universo=X|marca=Y`:
      • universo → reemplaza la columna '¿Cual es el universo en cms de la
        categoria?' en TODAS las filas con el mismo ID_ENCUESTA
      • marca → reemplaza la columna '¿Cuántos cms tiene la marca?' solo
        en la fila correspondiente a (ID_ENCUESTA, MARCA, CATEGORIA)

    Retorna (filas_modificadas, mensaje).
    """
    import pandas as pd

    meta = _parse_metadata(caso.get("METADATA", ""))
    id_enc = meta.get("id_encuesta", "")
    categ  = meta.get("categoria", "")
    marca  = (caso.get("MERCADERISTA_O_MARCA", "") or "").strip().upper()
    valores = _parse_valor_correcto_sos(caso.get("VALOR_CORRECTO", ""))

    if not id_enc:
        return 0, "Sin ID_ENCUESTA en METADATA — no se puede ubicar"
    if not valores:
        return 0, f"VALOR_CORRECTO vacío o malformado: {caso.get('VALOR_CORRECTO')!r}"

    df = pd.read_excel(source)
    col_univ  = "¿Cual es el universo en cms de la categoria?"
    col_marca = "¿Cuántos cms tiene la marca?"
    col_id    = "ID de la encuesta"
    col_cat   = "Categoría de producto"
    col_mk    = "Marca"

    if col_id not in df.columns:
        return 0, f"Source SOS sin columna '{col_id}'"

    df[col_id] = df[col_id].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    id_enc_str = str(id_enc).strip().replace(".0", "")

    mask_enc = df[col_id] == id_enc_str
    n_total = int(mask_enc.sum())
    if n_total == 0:
        return 0, f"No se encontraron filas con ID_ENCUESTA={id_enc_str}"

    modif = 0

    # 1) Universo → replicar a TODAS las filas del ID_ENCUESTA
    if "universo" in valores:
        if not dry_run:
            df.loc[mask_enc, col_univ] = valores["universo"]
        modif += n_total

    # 2) Marca → solo la fila puntual (mismo ID_ENCUESTA + MARCA + CATEGORIA)
    if "marca" in valores:
        df_marca = df[col_mk].astype(str).str.strip().str.upper()
        df_cat   = df[col_cat].astype(str).str.strip().str.upper()
        mask_pun = mask_enc & (df_marca == marca) & (df_cat == categ)
        n_pun = int(mask_pun.sum())
        if n_pun == 0:
            return modif, (
                f"Universo {'aplicado' if not dry_run else 'a aplicar'} "
                f"en {n_total} filas; pero no se halló fila puntual para "
                f"MARCA={marca} CATEG={categ}"
            )
        if not dry_run:
            df.loc[mask_pun, col_marca] = valores["marca"]
        modif += n_pun

    if not dry_run and modif > 0:
        df.to_excel(source, index=False)

    return modif, (
        f"Universo×{n_total if 'universo' in valores else 0}, "
        f"marca×{1 if 'marca' in valores else 0} "
        f"(ID_ENCUESTA={id_enc_str})"
    )


def aplicar_precios(caso: dict, source: Path, dry_run: bool) -> tuple[int, str]:
    """
    Modifica PRECIO_REGULAR (o PRECIO_PROMO si la regla era R10) en el
    archivo ANALISIS_PRECIOS para la fila identificada por (id_pdv, sku).
    """
    import pandas as pd

    meta = _parse_metadata(caso.get("METADATA", ""))
    id_pdv = str(meta.get("id_pdv", "")).strip()
    sku    = str(meta.get("sku", "")).strip().upper()
    regla  = meta.get("regla", "")
    valor  = caso.get("VALOR_CORRECTO", "")

    if not id_pdv or not sku:
        return 0, f"Falta id_pdv o sku en METADATA"

    try:
        valor_num = float(str(valor).replace("$", "").replace(",", "").strip())
    except (ValueError, TypeError):
        return 0, f"VALOR_CORRECTO no numérico: {valor!r}"

    df = pd.read_excel(source)
    df["ID del PDV"] = df["ID del PDV"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df["CODIGO_SKU"] = df.get("CODIGO_SKU", pd.Series("", index=df.index)).astype(str).str.strip().str.upper()

    mask = (df["ID del PDV"] == id_pdv) & (df["CODIGO_SKU"] == sku)
    if mask.sum() == 0:
        return 0, f"No se halló fila para id_pdv={id_pdv} sku={sku}"

    # R10 → promo; R11/R12 → regular
    col = "PRECIO_PROMO" if regla == "R10" else "PRECIO_REGULAR"
    if not dry_run:
        df.loc[mask, col] = valor_num
        df.to_excel(source, index=False)
    return int(mask.sum()), f"{col}={valor_num} en {int(mask.sum())} fila(s)"


def aplicar_np(caso: dict, source: Path, dry_run: bool) -> tuple[int, str]:
    """
    Para NP la corrección típica es eliminar el conteo erróneo (R4) o
    confirmar desabastecimiento (R3 — sin acción).
    """
    import pandas as pd

    meta = _parse_metadata(caso.get("METADATA", ""))
    decision = caso.get("DECISION", "")

    if decision in DECISIONES_ELIMINAR:
        id_pdv = str(meta.get("id_pdv", "")).strip()
        marca  = (caso.get("MERCADERISTA_O_MARCA", "") or "").strip().upper()
        if not id_pdv:
            return 0, "Falta id_pdv en METADATA"
        df = pd.read_excel(source)
        df["ID del PDV"] = df["ID del PDV"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        df["Marca_N"]    = df["Marca"].astype(str).str.strip().str.upper()
        mask = (df["ID del PDV"] == id_pdv) & (df["Marca_N"] == marca)
        n = int(mask.sum())
        if n == 0:
            return 0, f"No se halló fila NP para id_pdv={id_pdv} marca={marca}"
        if not dry_run:
            df = df[~mask].drop(columns=["Marca_N"])
            df.to_excel(source, index=False)
        return n, f"Eliminada(s) {n} fila(s) NP duplicada(s)"

    return 0, "Sin acción sobre el source (decisión informativa)"


def aplicar_cif(caso: dict, source: Path, dry_run: bool) -> tuple[int, str]:
    """
    CIF: las anomalías son agregados por gestor (sobrecumplimiento de
    visitas, tiempo > 9h). No hay un campo individual a sobreescribir
    en el Plan de trabajo. Las decisiones quedan en Historial como
    justificación pero el archivo fuente no se modifica.
    """
    return 0, "CIF — solo se registra en Historial (no se modifica source)"


def aplicar_exhibiciones(caso: dict, source: Path, dry_run: bool) -> tuple[int, str]:
    """Exhibiciones — agregados; igual que CIF, solo registro."""
    return 0, "EXHIBICIONES — solo se registra en Historial"


APLICADORES = {
    "SOS":                  aplicar_sos,
    "PRECIOS":              aplicar_precios,
    "NO_PRESENCIA":         aplicar_np,
    "CIF":                  aplicar_cif,
    "EXHIBICIONES_PAGADAS": aplicar_exhibiciones,
    "EXHIBICIONES_GRATIS":  aplicar_exhibiciones,
}


# ─────────────────────────────────────────────────────────────────────────────
# HISTORIAL — insertar registros de auditoría
# ─────────────────────────────────────────────────────────────────────────────

def _registrar_en_historial(conn: sqlite3.Connection, caso: dict,
                              n_filas: int, detalle: str) -> str:
    """Inserta uno o más registros en historial. Devuelve los ID_CORRECCION."""
    meta = _parse_metadata(caso.get("METADATA", ""))
    id_corr = f"HC-{caso['ID_CASO']}-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}"
    fila = {
        "ID_CORRECCION":  id_corr,
        "ID_PDV":         str(meta.get("id_pdv", "")),
        "NOMBRE_PDV":     caso.get("PDVS_AFECTADOS", "").split("—")[0].strip() or caso.get("PDVS_AFECTADOS", ""),
        "MES":            str(caso.get("MES", "")),
        "ANIO":           str(caso.get("ANIO", "")),
        "MODULO":         caso.get("MODULO", ""),
        "CAUSA":          caso.get("CAUSA", ""),
        "DECISION":       caso.get("DECISION", ""),
        "VALOR_ORIGINAL": caso.get("VALOR_ORIGINAL", ""),
        "VALOR_CORRECTO": caso.get("VALOR_CORRECTO", ""),
        "APROBADO_POR":   caso.get("APROBADO_POR", ""),
        "FECHA":          dt.datetime.now().isoformat(timespec="seconds"),
        "ID_CASO_ORIGEN": caso.get("ID_CASO", ""),
    }
    cols = list(fila.keys())
    vals = list(fila.values())
    placeholders = ",".join("?" * len(cols))
    cols_q = ",".join(f'"{c}"' for c in cols)
    conn.execute(
        f"INSERT OR REPLACE INTO historial ({cols_q}) VALUES ({placeholders})",
        vals,
    )
    return id_corr


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Aplicar correcciones aprobadas")
    parser.add_argument("--dry-run", action="store_true",
                        help="Muestra qué haría sin tocar archivos ni BD")
    parser.add_argument("--mes",  type=int, default=None)
    parser.add_argument("--anio", type=int, default=None)
    parser.add_argument("--solo", nargs="+",
                        help="Aplicar solo a estos ID_CASO específicos")
    args = parser.parse_args()

    print("═" * 70)
    print(f"  APLICAR CORRECCIONES — Eficacia")
    if args.dry_run:
        print("  ⚠️  DRY-RUN — no se modifica ningún archivo ni la BD")
    print(f"  DB: {DB_PATH}")
    print("═" * 70)

    if not DB_PATH.exists():
        print(f"❌ No existe la BD: {DB_PATH}")
        print("   Corre primero detector_anomalias.py + local_server.py")
        sys.exit(1)

    # Cargar fuentes por módulo
    sources = _sources_por_modulo()
    print(f"\n  Fuentes resueltas:")
    for m, p in sources.items():
        existe = "✓" if p.exists() else "✗ (faltante)"
        print(f"    {m:<22} {existe}  {p}")

    # Cargar aprobados de la BD
    with _conn() as conn:
        sql = "SELECT * FROM casos WHERE UPPER(ESTADO)='APROBADO'"
        params: list = []
        if args.mes:
            sql += " AND CAST(MES AS INTEGER) = ?"
            params.append(args.mes)
        if args.anio:
            sql += " AND CAST(ANIO AS INTEGER) = ?"
            params.append(args.anio)
        if args.solo:
            placeholders = ",".join("?" * len(args.solo))
            sql += f" AND ID_CASO IN ({placeholders})"
            params.extend(args.solo)

        aprobados = conn.execute(sql, params).fetchall()

    if not aprobados:
        print("\n  ℹ  No hay casos APROBADO para aplicar.")
        return

    print(f"\n  Casos APROBADOS a procesar: {len(aprobados)}")

    n_aplicados = 0
    n_sin_accion = 0
    n_errores = 0
    backups_hechos: dict[Path, Path] = {}

    with _conn() as conn:
        for r in aprobados:
            caso = dict(r)
            id_caso = caso.get("ID_CASO", "")
            modulo  = (caso.get("MODULO", "") or "").upper()
            decision = caso.get("DECISION", "")

            print(f"\n  ▶ {id_caso} ({modulo}) → {decision}")

            # 1) Decisiones sin acción al source — solo Historial
            if decision in DECISIONES_SIN_ACCION:
                if not args.dry_run:
                    id_h = _registrar_en_historial(conn, caso, 0, "Sin modificación al source")
                    conn.execute("UPDATE casos SET APLICADO_EN=? WHERE ID_CASO=?",
                                  (dt.datetime.now().isoformat(timespec="seconds"), id_caso))
                    conn.commit()
                print(f"     ✓ Registrado en Historial (sin tocar source)")
                n_sin_accion += 1
                continue

            # 2) Decisiones que modifican el source
            source = sources.get(modulo)
            aplicador = APLICADORES.get(modulo)
            if not aplicador or not source:
                print(f"     ⚠️  Módulo no soportado o sin source: {modulo}")
                n_errores += 1
                continue
            if not source.exists():
                print(f"     ⚠️  Source no existe: {source}")
                n_errores += 1
                continue

            # Backup (una vez por archivo)
            if not args.dry_run and source not in backups_hechos:
                bk = _backup(source)
                backups_hechos[source] = bk
                print(f"     ▸ Backup: {bk.name}")

            try:
                n_filas, detalle = aplicador(caso, source, args.dry_run)
                print(f"     ✓ {detalle}")
                if not args.dry_run:
                    _registrar_en_historial(conn, caso, n_filas, detalle)
                    conn.execute("UPDATE casos SET APLICADO_EN=? WHERE ID_CASO=?",
                                  (dt.datetime.now().isoformat(timespec="seconds"), id_caso))
                    conn.commit()
                n_aplicados += 1
            except Exception as e:
                print(f"     ❌ Error: {e}")
                n_errores += 1

    print("\n" + "═" * 70)
    print(f"  RESUMEN")
    print(f"    Aplicados al source:    {n_aplicados}")
    print(f"    Solo Historial:         {n_sin_accion}")
    print(f"    Errores:                {n_errores}")
    if backups_hechos:
        print(f"    Backups creados:        {len(backups_hechos)} en {DIR_BACKUPS}")
    print("═" * 70)


if __name__ == "__main__":
    main()
