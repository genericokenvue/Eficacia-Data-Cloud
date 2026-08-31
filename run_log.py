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

# ─────────────────────────────────────────────────────────────────────────────
# Localización del run_log
# ─────────────────────────────────────────────────────────────────────────────
# El historial de corridas vive en SharePoint, NO en Supabase: es un registro
# operativo del pipeline (qué ETL corrió, cuándo, con qué hashes), no un dato
# de negocio para BI. Mezclarlo con las tablas analíticas ensuciaba la base.
#
# Se guarda como JSONL — un evento por línea — porque solo se escribe
# añadiendo al final y se lee de corrido.

RUTA_CARPETA_LOGS_CLOUD = f"{paths._SALIDAS_ROOT}/LOGS"
RUTA_RUN_LOG_CLOUD      = f"{RUTA_CARPETA_LOGS_CLOUD}/run_log.jsonl"

LOGS_DIR     = Path(__file__).parent / "logs"
RUN_LOG_PATH = LOGS_DIR / "run_log.jsonl"


# ─────────────────────────────────────────────────────────────────────────────
# Acceso a SharePoint (JSONL como archivo de texto plano)
# ─────────────────────────────────────────────────────────────────────────────

_cache_headers: dict | None = None
_cache_site_id: str | None = None


def _conexion_cloud() -> tuple[dict, str] | None:
    """Devuelve (headers, site_id) o None si no hay credenciales/conexión."""
    global _cache_headers, _cache_site_id
    if _cache_headers is not None and _cache_site_id is not None:
        return _cache_headers, _cache_site_id
    try:
        import msal
        import requests

        if not all([paths.AZURE_TENANT_ID, paths.AZURE_CLIENT_ID, paths.AZURE_CLIENT_SECRET]):
            return None

        app = msal.ConfidentialClientApplication(
            paths.AZURE_CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{paths.AZURE_TENANT_ID}",
            client_credential=paths.AZURE_CLIENT_SECRET,
        )
        res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in res:
            return None

        headers = {"Authorization": f"Bearer {res['access_token']}"}
        url = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{paths.SHAREPOINT_SITE_NAME}"
        r = requests.get(url, headers=headers).json()
        if "id" not in r:
            return None

        _cache_headers, _cache_site_id = headers, r["id"]
        return _cache_headers, _cache_site_id
    except Exception:
        return None


def _leer_run_log_cloud() -> str:
    """Contenido actual del run_log en SharePoint ('' si no existe)."""
    conn = _conexion_cloud()
    if not conn:
        return ""
    headers, site_id = conn
    try:
        import urllib.parse
        import requests
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}"
            f"/drive/root:/{urllib.parse.quote(RUTA_RUN_LOG_CLOUD)}:/content"
        )
        r = requests.get(url, headers=headers)
        return r.text if r.status_code == 200 else ""
    except Exception:
        return ""


def _asegurar_carpeta_logs(headers: dict, site_id: str) -> None:
    """
    Crea SALIDAS/LOGS en SharePoint si no existe.

    Graph no crea las carpetas intermedias al subir un archivo: si la carpeta
    falta, el PUT devuelve 404. En un runner nuevo de GitHub Actions nadie la
    ha creado a mano, así que hay que asegurarla antes del primer registro.
    """
    import urllib.parse
    import requests

    padre, _, nombre = RUTA_CARPETA_LOGS_CLOUD.rpartition("/")
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{urllib.parse.quote(padre)}:/children"
    )
    requests.post(
        url,
        headers={**headers, "Content-Type": "application/json"},
        json={
            "name": nombre,
            "folder": {},
            # Si ya existe, no se toca ni se duplica.
            "@microsoft.graph.conflictBehavior": "fail",
        },
    )


def _escribir_run_log_cloud(contenido: str) -> bool:
    conn = _conexion_cloud()
    if not conn:
        return False
    headers, site_id = conn
    try:
        import urllib.parse
        import requests
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}"
            f"/drive/root:/{urllib.parse.quote(RUTA_RUN_LOG_CLOUD)}:/content"
        )
        cabeceras = {**headers, "Content-Type": "text/plain; charset=utf-8"}
        r = requests.put(url, headers=cabeceras, data=contenido.encode("utf-8"))

        # 404 = la carpeta LOGS todavía no existe. Se crea y se reintenta una vez.
        if r.status_code == 404:
            _asegurar_carpeta_logs(headers, site_id)
            r = requests.put(url, headers=cabeceras, data=contenido.encode("utf-8"))

        return r.status_code in (200, 201)
    except Exception:
        return False


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
    Añade el evento al run_log de SharePoint (SALIDAS/LOGS/run_log.jsonl).

    Se guarda en la nube porque en GitHub Actions el disco es efímero: un
    .jsonl local se perdería en cada corrida y `--skip-if-cached` nunca
    encontraría el run anterior.

    JSONL no admite "append" por API, así que se descarga, se agrega la línea
    y se vuelve a subir. El archivo es pequeño (una línea por ETL y periodo),
    así que el costo es despreciable.

    Si SharePoint no está disponible, cae a un respaldo local — no bloquea
    el pipeline por no poder registrar el evento.
    """
    linea = event.to_jsonl()

    contenido = _leer_run_log_cloud()
    if contenido and not contenido.endswith("\n"):
        contenido += "\n"
    if _escribir_run_log_cloud(contenido + linea):
        return

    print("  ⚠ No se pudo registrar el run en SharePoint; guardo respaldo local")
    _ensure_logs_dir()
    with open(RUN_LOG_PATH, "a", encoding="utf-8") as f:
        f.write(linea)


# ─────────────────────────────────────────────────────────────────────────────
# Lectura / consultas
# ─────────────────────────────────────────────────────────────────────────────

def _iter_eventos() -> Iterable[dict]:
    """Recorre el run_log: primero el de SharePoint, si no el respaldo local."""
    contenido = _leer_run_log_cloud()

    if not contenido and RUN_LOG_PATH.is_file():
        contenido = RUN_LOG_PATH.read_text(encoding="utf-8")

    for linea in contenido.splitlines():
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
    Lee el run_log de SharePoint; si no está disponible usa el respaldo local.
    """
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