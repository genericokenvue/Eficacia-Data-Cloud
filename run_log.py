"""
run_log.py — Sprint 16.2
─────────────────────────
Registro de ejecuciones del pipeline ETL en formato JSON Lines (append-only).
Permite trazabilidad histórica y skip de runs ya completados cuando los
inputs no cambiaron (idempotencia opcional).

Cada línea del run_log representa un evento (etl, periodo) con:
  - run_id      : identificador del run del orquestador (varios ETLs comparten)
  - ts_inicio   : ISO 8601 inicio del ETL
  - ts_fin      : ISO 8601 fin del ETL
  - duracion_s  : segundos
  - mes, anio   : periodo procesado
  - etiqueta    : "Mayo_2026"
  - etl         : nombre del ETL (cif, nopresencia, ...)
  - status      : OK | FAIL | SKIPPED
  - error_msg   : str|None
  - input_hash  : MD5 concatenado del contenido de inputs (None si no aplica)
  - output_hash : MD5 concatenado del contenido de outputs (None si no aplica)
  - files_input : list[str] (nombre + tamaño)
  - files_output: list[str]

Uso CLI:
    python run_log.py --list                       # últimos 20 runs
    python run_log.py --list --periodo 05/2026     # filtrar
    python run_log.py --list --etl cif --limit 50
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import secrets
import sys
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

import paths
import periodo_resolver as pr

try:
    from supabase import create_client
    _SUPABASE_URL = os.environ.get("SUPABASE_URL")
    _SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
    supabase = create_client(_SUPABASE_URL, _SUPABASE_KEY) if (_SUPABASE_URL and _SUPABASE_KEY) else None
except Exception:
    supabase = None

TABLA_RUN_LOG = "etl_run_log"


# ─────────────────────────────────────────────────────────────────────────────
# Localización del run_log
# ─────────────────────────────────────────────────────────────────────────────

LOGS_DIR     = Path(__file__).parent / "logs"
RUN_LOG_PATH = LOGS_DIR / "run_log.jsonl"


def _ensure_logs_dir() -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)


# ─────────────────────────────────────────────────────────────────────────────
# Mapeos ETL → inputs/outputs resolvibles por periodo
# ─────────────────────────────────────────────────────────────────────────────

# Cada entrada es una lista de callables (spec → Path). Si alguno falla con
# FileNotFoundError, el hash incluye el marcador "MISSING:<basename>".
INPUT_RESOLVERS: dict[str, list[Callable[[pr.PeriodoSpec], Path]]] = {
    "cif": [
        pr.cif_pt_directo,
        pr.cif_pt_ism,
        pr.cif_involves,
    ],
    "nopresencia": [
        pr.cif_pt_directo,
        pr.cif_pt_ism,
        pr.np_respuestas,
    ],
    "precios": [
        pr.cif_pt_directo,
        pr.cif_pt_ism,
        pr.precios_respuestas,
    ],
    "sos": [
        pr.cif_pt_directo,
        pr.cif_pt_ism,
        pr.sos_respuestas,
        pr.sos_respuestas_baby,
        pr.sos_target,
    ],
    "exhib_pagadas": [
        pr.exh_planning_master,
        pr.exh_base_planning,
    ],
    "exhib_gratis": [
        pr.exh_planning_master,
        pr.exh_base_informar,
        pr.exh_visitas,
    ],
    # ventas/impactos/impactos_segmentos son multi-mes — no soportan skip-if-cached
    # en este sprint. Su input_hash queda en None.
    "ventas": [],
    "impactos": [],
    "impactos_segmentos": [],
}

# Outputs por ETL (no dependen del periodo en sus paths — el histórico
# se filtra/upserta internamente por periodo).
OUTPUT_PATHS: dict[str, list[Path]] = {
    "cif":                [paths.CIF_OUT_CIF, paths.CIF_OUT_KPIS_HISTORICO],
    "nopresencia":        [paths.NP_OUT_KPIS, paths.NP_OUT_KPIS_HISTORICO],
    "precios":            [paths.PR_OUT_KPIS, paths.PR_OUT_KPIS_HISTORICO],
    "sos":                [paths.SOS_OUT_KPIS, paths.SOS_OUT_KPIS_HISTORICO],
    "exhib_pagadas":      [paths.EXHIB_PAG_OUT_KPIS, paths.EXHIB_PAG_OUT_KPIS_HISTORICO],
    "exhib_gratis":       [paths.EXHIB_GRA_OUT_KPIS, paths.EXHIB_GRA_OUT_KPIS_HISTORICO],
    "ventas":             [paths.DYP_OUT_VENTAS],
    "impactos":           [paths.DYP_OUT_IMPACTOS],
    "impactos_segmentos": [paths.DYP_OUT_SEGMENTOS],
}


def inputs_de(etl: str, spec: pr.PeriodoSpec) -> list[Path | None]:
    """
    Devuelve los Path resueltos por periodo para un ETL.
    Cada elemento es Path si el resolver tuvo éxito, o None si falló
    (típicamente FileNotFoundError porque el archivo no existe aún).
    """
    resolvers = INPUT_RESOLVERS.get(etl, [])
    out: list[Path | None] = []
    for fn in resolvers:
        try:
            out.append(fn(spec))
        except (FileNotFoundError, ValueError):
            out.append(None)
    return out


def outputs_de(etl: str) -> list[Path]:
    return OUTPUT_PATHS.get(etl, [])


# ─────────────────────────────────────────────────────────────────────────────
# Hashing
# ─────────────────────────────────────────────────────────────────────────────

def _hash_file(p: Path) -> tuple[str, int]:
    """Devuelve (md5_hex, size_bytes). Lee en chunks de 256 KB."""
    h = hashlib.md5()
    size = 0
    with open(p, "rb") as f:
        while True:
            chunk = f.read(256 * 1024)
            if not chunk:
                break
            h.update(chunk)
            size += len(chunk)
    return h.hexdigest(), size


def hash_files(paths_in: Iterable[Path | None]) -> tuple[str, list[str]]:
    """
    Hashea una lista de archivos. Para cada uno produce "<basename>:<md5>:<size>"
    y los concatena en un MD5 final. Si un Path es None o el archivo no existe,
    añade "MISSING:<repr>".

    Devuelve (digest_hex, descripciones).
    Si la lista está vacía o todos MISSING → digest = "EMPTY".
    """
    descripciones: list[str] = []
    h = hashlib.md5()
    contribuyentes = 0

    for p in paths_in:
        if p is None:
            descripciones.append("MISSING:<unresolved>")
            h.update(b"MISSING:<unresolved>\n")
            continue
        p = Path(p)
        if not p.is_file():
            descripciones.append(f"MISSING:{p.name}")
            h.update(f"MISSING:{p.name}\n".encode("utf-8"))
            continue
        digest, size = _hash_file(p)
        descripciones.append(f"{p.name}:{digest}:{size}")
        h.update(f"{p.name}:{digest}:{size}\n".encode("utf-8"))
        contribuyentes += 1

    if not descripciones:
        return "EMPTY", []
    if contribuyentes == 0:
        return "EMPTY", descripciones
    return h.hexdigest(), descripciones


# ─────────────────────────────────────────────────────────────────────────────
# Run IDs
# ─────────────────────────────────────────────────────────────────────────────

def gen_run_id() -> str:
    """Genera un run_id único: YYYYMMDDTHHMMSS_<6 hex>."""
    return f"{datetime.now().strftime('%Y%m%dT%H%M%S')}_{secrets.token_hex(3)}"


# ─────────────────────────────────────────────────────────────────────────────
# Log writer
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class RunEvent:
    run_id: str
    ts_inicio: str
    ts_fin: str
    duracion_s: float
    mes: int
    anio: int
    etiqueta: str
    etl: str
    status: str            # OK | FAIL | SKIPPED
    error_msg: str | None
    input_hash: str | None
    output_hash: str | None
    files_input: list[str]
    files_output: list[str]

    def to_jsonl(self) -> str:
        return json.dumps(self.__dict__, ensure_ascii=False) + "\n"


def log_run(event: RunEvent) -> None:
    """
    Registra el evento en Supabase (tabla etl_run_log) — así el historial
    persiste entre corridas de GitHub Actions, donde el disco es efímero y
    un .jsonl local se perdía en cada ejecución.

    Si Supabase no está configurado (sin SUPABASE_URL/KEY) o la carga
    falla, cae a un respaldo local en SCRIPTS/logs/run_log.jsonl — no
    bloquea el pipeline por esto.
    """
    if supabase is not None:
        try:
            registro = asdict(event)
            supabase.table(TABLA_RUN_LOG).insert(registro).execute()
            return
        except Exception as e:
            print(f"  ⚠ No se pudo registrar el run en Supabase ({e}); guardo respaldo local")

    _ensure_logs_dir()
    with open(RUN_LOG_PATH, "a", encoding="utf-8") as f:
        f.write(event.to_jsonl())


# ─────────────────────────────────────────────────────────────────────────────
# Lectura / consultas
# ─────────────────────────────────────────────────────────────────────────────

def _iter_eventos() -> Iterable[dict]:
    if not RUN_LOG_PATH.is_file():
        return
    with open(RUN_LOG_PATH, "r", encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if not linea:
                continue
            try:
                yield json.loads(linea)
            except json.JSONDecodeError:
                continue


def last_successful_run(etl: str, mes: int, anio: int) -> dict | None:
    """
    Último evento OK para ese (etl, periodo). None si no existe.
    Consulta Supabase primero (persiste entre corridas de GitHub Actions);
    si no está disponible, cae al .jsonl local (útil para pruebas offline).
    """
    if supabase is not None:
        try:
            resp = (
                supabase.table(TABLA_RUN_LOG)
                .select("*")
                .eq("etl", etl).eq("mes", mes).eq("anio", anio).eq("status", "OK")
                .order("ts_fin", desc=True)
                .limit(1)
                .execute()
            )
            filas = resp.data or []
            return filas[0] if filas else None
        except Exception as e:
            print(f"  ⚠ No se pudo consultar Supabase ({e}); reviso el .jsonl local")

    candidato = None
    for ev in _iter_eventos():
        if ev.get("etl") == etl and ev.get("mes") == mes and ev.get("anio") == anio \
           and ev.get("status") == "OK":
            candidato = ev
    return candidato


# ─────────────────────────────────────────────────────────────────────────────
# Decisión de skip
# ─────────────────────────────────────────────────────────────────────────────

def should_skip(etl: str, spec: pr.PeriodoSpec) -> tuple[bool, str]:
    """
    Determina si se puede skipear un (etl, periodo). Reglas:

      1. Existe un run_log OK previo para este (etl, mes, anio).
      2. El input_hash actual coincide con el del run anterior.
      3. Todos los outputs declarados existen en disco.

    Si las 3 se cumplen → True con motivo. Si no → False con motivo.

    Para ETLs sin inputs resolvibles (ventas, impactos, ...) devuelve
    False con motivo "etl-sin-input-tracking" (nunca skipea).
    """
    inputs = inputs_de(etl, spec)
    if not inputs:
        return False, "etl-sin-input-tracking"

    cur_hash, _ = hash_files(inputs)
    if cur_hash == "EMPTY":
        return False, "inputs-no-resueltos"

    prev = last_successful_run(etl, spec.mes, spec.anio)
    if prev is None:
        return False, "sin-run-previo"

    if prev.get("input_hash") != cur_hash:
        return False, "input-cambio"

    outs = outputs_de(etl)
    faltantes = [str(p) for p in outs if not Path(p).is_file()]
    if faltantes:
        return False, f"outputs-faltantes:{len(faltantes)}"

    return True, f"hash-coincide:{cur_hash[:8]}"


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de alto nivel para el orquestador
# ─────────────────────────────────────────────────────────────────────────────

def calcular_input_hash(etl: str, spec: pr.PeriodoSpec) -> tuple[str | None, list[str]]:
    """Devuelve (hash, files_input) o (None, []) si no aplica."""
    inputs = inputs_de(etl, spec)
    if not inputs:
        return None, []
    digest, files = hash_files(inputs)
    return digest, files


def calcular_output_hash(etl: str) -> tuple[str | None, list[str]]:
    outs = outputs_de(etl)
    if not outs:
        return None, []
    digest, files = hash_files(outs)
    return digest, files


# ─────────────────────────────────────────────────────────────────────────────
# CLI de inspección
# ─────────────────────────────────────────────────────────────────────────────

def _obtener_eventos_filtrados(periodo: str | None, etl: str | None, limit: int) -> list[dict]:
    if supabase is not None:
        try:
            q = supabase.table(TABLA_RUN_LOG).select("*").order("ts_fin", desc=True)
            if periodo:
                spec = pr.parsear_periodo_str(periodo)
                q = q.eq("mes", spec.mes).eq("anio", spec.anio)
            if etl:
                q = q.eq("etl", etl)
            if limit > 0:
                q = q.limit(limit)
            resp = q.execute()
            eventos = resp.data or []
            eventos.reverse()  # para imprimir en orden cronológico ascendente, igual que antes
            return eventos
        except Exception as e:
            print(f"  ⚠ No se pudo consultar Supabase ({e}); reviso el .jsonl local\n")

    eventos = list(_iter_eventos())
    if periodo:
        spec = pr.parsear_periodo_str(periodo)
        eventos = [e for e in eventos if e.get("mes") == spec.mes and e.get("anio") == spec.anio]
    if etl:
        eventos = [e for e in eventos if e.get("etl") == etl]
    return eventos[-limit:] if limit > 0 else eventos


def _cli_list(args: argparse.Namespace) -> None:
    eventos = _obtener_eventos_filtrados(args.periodo, args.etl, args.limit)
    if not eventos:
        print("(sin runs registrados con esos filtros)")
        return

    print(f"{'ts_fin':<19}  {'periodo':<12}  {'etl':<20}  {'status':<8}  {'dur(s)':>7}  {'input_hash':<10}  {'motivo/error':<40}")
    print("─" * 130)
    for ev in eventos:
        ts = (ev.get("ts_fin") or ev.get("ts_inicio") or "")[:19].replace("T", " ")
        et = ev.get("etiqueta", "?")
        etl = ev.get("etl", "?")
        st  = ev.get("status", "?")
        d   = ev.get("duracion_s", 0.0)
        ih  = (ev.get("input_hash") or "—")[:8]
        msg = (ev.get("error_msg") or "")[:40]
        print(f"{ts:<19}  {et:<12}  {etl:<20}  {st:<8}  {d:>7.1f}  {ih:<10}  {msg:<40}")


def _main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="Inspector del run_log de Eficacia")
    parser.add_argument("--list", action="store_true",
                        help="Listar runs registrados (default si no se pasa nada)")
    parser.add_argument("--periodo", type=str, default=None,
                        help="Filtrar por periodo MM/AAAA")
    parser.add_argument("--etl", type=str, default=None,
                        help="Filtrar por nombre de ETL")
    parser.add_argument("--limit", type=int, default=20,
                        help="Cantidad máxima a mostrar (default 20, 0=todos)")

    args = parser.parse_args()
    _cli_list(args)


if __name__ == "__main__":
    _main()