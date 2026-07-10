"""
run_revision_local.py
─────────────────────
Orquestador de UN comando para una sesión de pruebas presencial.

Hace:
  1) Detector de anomalías (si falta el Excel del periodo o se pasa --regenerar)
  2) Importa el Excel a SQLite (upsert seguro — conserva decisiones previas)
  3) Levanta el servidor Flask en localhost:5000
  4) Abre el navegador en http://localhost:5000

El supervisor en el kiosko verá la pantalla de login → selecciona su nombre →
revisa sus casos. Otro supervisor refresca, selecciona el suyo, y sigue.

Uso típico
──────────
  python run_revision_local.py                  # arranca todo
  python run_revision_local.py --reset          # borra BD y empieza limpio
  python run_revision_local.py --regenerar      # re-corre el detector
  python run_revision_local.py --puerto 8000    # otro puerto
  python run_revision_local.py --host 0.0.0.0   # multi-PC en red local
  python run_revision_local.py --no-browser     # no abrir el navegador
"""

from __future__ import annotations

import os
import sys
import argparse
import subprocess
import webbrowser
from pathlib import Path

# UTF-8 stdout
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import paths


DIR_REV = paths.ALERTAS_DIR / "REVISION"


def _excel_periodo_existe() -> Path | None:
    candidatos = sorted(DIR_REV.glob("Casos_Revision_*.xlsx"))
    return candidatos[-1] if candidatos else None


def main():
    parser = argparse.ArgumentParser(
        description="Sesión de pruebas presencial — orquestador local",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("--reset",      action="store_true",
                        help="Borra la BD antes de empezar (pierde decisiones previas)")
    parser.add_argument("--regenerar",  action="store_true",
                        help="Re-corre el detector aunque ya exista el Excel")
    parser.add_argument("--puerto",     type=int, default=5000)
    parser.add_argument("--host",       default="127.0.0.1",
                        help="127.0.0.1 (kiosko) o 0.0.0.0 (multi-PC en red local)")
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    print("═" * 70)
    print("  SESIÓN DE REVISIÓN PRESENCIAL — Eficacia")
    print("═" * 70)

    # 1) Generar/encontrar el Excel del periodo
    excel = _excel_periodo_existe()
    if not excel or args.regenerar:
        print("\n[1/3] Corriendo detector_anomalias.py ...")
        cmd = [sys.executable, str(_SCRIPTS / "detector_anomalias.py")]
        r = subprocess.run(cmd)
        if r.returncode != 0:
            print("❌ Detector falló.")
            sys.exit(1)
        excel = _excel_periodo_existe()
        if not excel:
            print("❌ El detector no produjo Casos_Revision_*.xlsx")
            sys.exit(1)
    else:
        print(f"\n[1/3] Excel existente — usando: {excel.name}")
        print("       (usa --regenerar para forzar re-detección)")

    # 2) Importar a SQLite vía local_server.py (--no-server)
    print(f"\n[2/3] Importando a SQLite ...")
    cmd = [sys.executable, str(_SCRIPTS / "local_server.py"),
            "--import", str(excel), "--no-server"]
    if args.reset:
        cmd.append("--reset")
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print("❌ Importación a SQLite falló.")
        sys.exit(1)

    # 3) Levantar servidor + abrir navegador
    print(f"\n[3/3] Levantando servidor en http://{args.host}:{args.puerto} ...")
    if not args.no_browser:
        # Abrir browser cuando el server esté arriba (delay 1.5s)
        url = f"http://localhost:{args.puerto}/"
        import threading, time
        def _abrir():
            time.sleep(1.5)
            webbrowser.open(url)
        threading.Thread(target=_abrir, daemon=True).start()
        print(f"       Abriendo browser en: {url}")

    print(f"\n  ───────────────────────────────────────────────────────")
    print(f"  Instrucciones para los supervisores:")
    print(f"    1. En el browser: selecciona tu nombre del dropdown")
    print(f"    2. Revisa tus casos pendientes")
    print(f"    3. Confirma decisiones (cada PATCH se guarda al instante)")
    print(f"    4. Cuando termines, click 'Cambiar usuario' para el siguiente")
    print(f"")
    print(f"  Ctrl+C aquí para detener el servidor.")
    print(f"  ───────────────────────────────────────────────────────\n")

    # Ejecutar el server (bloquea hasta Ctrl+C)
    cmd = [sys.executable, str(_SCRIPTS / "local_server.py"),
            "--port", str(args.puerto), "--host", args.host]
    try:
        subprocess.run(cmd)
    except KeyboardInterrupt:
        print("\n\n  ✋ Sesión detenida por el usuario.")


if __name__ == "__main__":
    main()
