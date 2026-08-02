"""
config_loader.py
────────────────
Lectura del archivo ALERTAS/config.env (formato KEY=VALUE).

Importar `cargar_config()` desde cualquier script de la Fase 2 al inicio
para que las claves del archivo queden disponibles en os.environ. Las
variables de entorno reales tienen precedencia (permiten override en
CI/tests). Maneja comillas, espacios y los marcadores < > que se suelen
dejar accidentalmente al pegar tokens.
"""

import os
from pathlib import Path

# La carpeta del propio archivo (ALERTAS/) SIEMPRE existe — es donde vive
# config.env junto a los demás scripts de Fase 2. El fallback anterior a
# una ruta fija de Windows era código muerto (nunca se alcanzaba) y además
# no tenía sentido fuera de esa máquina (p. ej. GitHub Actions). Se deja
# EFICACIA_BASE como override explícito por consistencia con paths.py.
_OVERRIDE = os.environ.get("EFICACIA_BASE", "").strip()
DIR_ALERTAS = (Path(_OVERRIDE) / "SCRIPTS" / "ALERTAS") if _OVERRIDE else Path(__file__).resolve().parent
RUTA_ENV    = DIR_ALERTAS / "config.env"


def cargar_config() -> dict:
    """
    Vuelca el contenido de config.env a os.environ y devuelve el dict
    cargado. Idempotente: usa os.environ.setdefault para no pisar
    variables ya exportadas externamente.
    """
    cargado: dict[str, str] = {}
    if not RUTA_ENV.exists():
        return cargado
    for linea in RUTA_ENV.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#") or "=" not in linea:
            continue
        clave, _, valor = linea.partition("=")
        clave = clave.strip()
        valor = valor.strip().strip('"').strip("'")
        if valor.startswith("<") and valor.endswith(">"):
            valor = valor[1:-1].strip()
        cargado[clave] = valor
        os.environ.setdefault(clave, valor)
    return cargado
