"""
bot_listener.py
───────────────
Listener continuo del bot de Telegram para capturar chat_ids fuera de la
ventana de 24 horas que limita `getUpdates`.

Funcionamiento
──────────────
1. Hace long-polling cada N segundos a `getUpdates` con `offset` para
   marcar como leídas las actualizaciones (así no se pierden).
2. Para CADA mensaje nuevo recibido:
   • Lo registra/actualiza en `ALERTAS/telegram_chat_ids.csv`
     (mismo formato que `setup_telegram.py`).
   • Le RESPONDE al usuario con su chat_id, para que lo pueda
     compartir con coordinación si la captura automática falla.
3. Es seguro mantenerlo corriendo días/semanas — usa offset persistente
   guardado en `ALERTAS/.bot_listener_offset` para no re-procesar.

Uso
───
    cd ALERTAS
    python bot_listener.py             # corre indefinidamente
    python bot_listener.py --once      # un solo poll y sale
    python bot_listener.py --silent    # no responde con el chat_id

Para detener: Ctrl+C.

Por qué resuelve el problema de la ventana de 24h
─────────────────────────────────────────────────
Mientras este listener esté corriendo, cualquier supervisor que le
escriba al bot queda capturado al instante. Si el listener se detiene,
los nuevos mensajes se acumulan en Telegram hasta 24h; al reiniciarlo,
los recoge todos.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

from config_loader import cargar_config

# Reusamos los extractores de updates de setup_telegram (DRY)
from setup_telegram import (
    _extraer_registro,
    _ahora_colombia,
    _leer_csv_existente,
    _guardar_csv,
    CSV_SALIDA,
    DIR_ALERTAS,
)


OFFSET_FILE = DIR_ALERTAS / ".bot_listener_offset"
INTERVALO_POLL_S = 5            # polling cada 5s (Telegram permite long polling de 50s)
TIMEOUT_LONGPOLL_S = 30         # long polling: el server espera hasta 30s a que llegue algo


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

def _cargar_token() -> str:
    cargar_config()
    token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    if not token or token.startswith("<") or "REEMPLAZAR" in token.upper():
        raise ValueError(
            "TELEGRAM_TOKEN no configurado.\n"
            "Edita ALERTAS/config.env y pon: TELEGRAM_TOKEN=NN:AA... sin < >"
        )
    return token


def _leer_offset() -> int | None:
    if not OFFSET_FILE.exists():
        return None
    try:
        return int(OFFSET_FILE.read_text(encoding="utf-8").strip())
    except Exception:
        return None


def _guardar_offset(update_id: int) -> None:
    OFFSET_FILE.write_text(str(update_id), encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# CONSULTA DE UPDATES + RESPUESTA AL USUARIO
# ─────────────────────────────────────────────────────────────────────────────

def _enviar_respuesta_chat_id(token: str, chat_id: str, nombre: str) -> None:
    """
    Le responde al usuario con su chat_id como un mensaje. Útil para que
    el supervisor lo vea y, si la captura automática falla por algún motivo,
    pueda enviarlo manualmente al admin.
    """
    saludo = f"¡Hola {nombre.title()}!" if nombre else "¡Hola!"
    texto = (
        f"{saludo} Soy el bot de alertas de Eficacia.\n\n"
        f"📋 *Tu chat_id es:* `{chat_id}`\n\n"
        f"Ya estás registrado para recibir el reporte mensual de "
        f"cumplimiento de tu equipo."
    )
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    try:
        resp = requests.post(
            url,
            json={"chat_id": chat_id, "text": texto, "parse_mode": "Markdown"},
            timeout=10,
        )
        if resp.status_code != 200:
            print(f"      ⚠️  No pude responder a {chat_id}: HTTP {resp.status_code}")
    except Exception as e:
        print(f"      ⚠️  Error enviando respuesta a {chat_id}: {e}")


def _consultar_updates(token: str, offset: int | None) -> list:
    """getUpdates con long-polling. Devuelve la lista de updates."""
    url = f"https://api.telegram.org/bot{token}/getUpdates"
    params = {
        "timeout": TIMEOUT_LONGPOLL_S,
        # allowed_updates vacío → todos los tipos por defecto
    }
    if offset is not None:
        params["offset"] = offset

    try:
        resp = requests.get(url, params=params, timeout=TIMEOUT_LONGPOLL_S + 10)
        data = resp.json()
        if not data.get("ok"):
            print(f"  ⚠️  Telegram rechazó el getUpdates: {data}")
            return []
        return data.get("result", [])
    except requests.exceptions.Timeout:
        return []
    except Exception as e:
        print(f"  ⚠️  Error en getUpdates: {e}")
        return []


# ─────────────────────────────────────────────────────────────────────────────
# CICLO PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def procesar_un_ciclo(token: str, responder: bool = True) -> tuple[int, int]:
    """
    Procesa un poll. Devuelve (n_nuevos, n_total_csv).
    """
    offset = _leer_offset()
    updates = _consultar_updates(token, offset)
    if not updates:
        # Cargar solo para reportar total
        total = len(_leer_csv_existente(CSV_SALIDA))
        return 0, total

    registros = _leer_csv_existente(CSV_SALIDA)
    n_nuevos = 0
    chat_ids_respondidos: set[str] = set()  # evita doble respuesta en un mismo poll

    for upd in updates:
        registro = _extraer_registro(upd)
        if registro is None:
            continue

        chat_id = registro["chat_id"]
        es_nuevo = chat_id not in registros
        registros[chat_id] = registro
        if es_nuevo:
            n_nuevos += 1
            print(
                f"  ✅ Nuevo chat_id capturado: {chat_id}  "
                f"({registro.get('nombre_telegram','-')} @{registro.get('username','-')})"
            )

        # Responder al usuario con su chat_id (solo si es 'message' / 'edited_message'
        # — no en eventos como my_chat_member para evitar spam)
        if responder and chat_id not in chat_ids_respondidos \
                and registro.get("tipo_update") in ("message", "edited_message",
                                                     "business_message"):
            _enviar_respuesta_chat_id(
                token, chat_id, registro.get("nombre_telegram", "")
            )
            chat_ids_respondidos.add(chat_id)

    if updates:
        _guardar_csv(CSV_SALIDA, registros)
        ultimo_id = max(int(u.get("update_id", 0)) for u in updates)
        _guardar_offset(ultimo_id + 1)

    return n_nuevos, len(registros)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Listener continuo del bot Telegram (captura chat_ids)."
    )
    parser.add_argument("--once", action="store_true",
                        help="Ejecuta un solo poll y sale.")
    parser.add_argument("--silent", action="store_true",
                        help="No responde al usuario con su chat_id "
                             "(solo lo guarda en el CSV).")
    parser.add_argument("--intervalo", type=int, default=INTERVALO_POLL_S,
                        help=f"Segundos entre polls (default: {INTERVALO_POLL_S})")
    args = parser.parse_args()

    token = _cargar_token()
    responder = not args.silent

    print("═" * 60)
    print("  Bot Listener Eficacia — captura continua de chat_ids")
    print("═" * 60)
    print(f"  CSV histórico:    {CSV_SALIDA}")
    print(f"  Offset state:     {OFFSET_FILE}")
    print(f"  Modo:             {'one-shot' if args.once else 'continuo'}")
    print(f"  Responder al user: {'sí' if responder else 'no (silent)'}")
    print()
    print("  Pídeles a los supervisores que escriban /start al bot.")
    print("  Detén con Ctrl+C cuando hayas capturado todos.")
    print("═" * 60)

    if args.once:
        n_nuevos, n_total = procesar_un_ciclo(token, responder=responder)
        print(f"\nResultado: {n_nuevos} nuevos | total en CSV: {n_total}")
        return 0

    try:
        while True:
            n_nuevos, n_total = procesar_un_ciclo(token, responder=responder)
            estado = (
                f"[{datetime.now().strftime('%H:%M:%S')}] "
                f"+{n_nuevos} nuevo(s) | total CSV: {n_total}"
            )
            # Imprimir solo si hubo cambio o cada minuto para no saturar
            if n_nuevos > 0:
                print(f"  {estado}")
            time.sleep(args.intervalo)
    except KeyboardInterrupt:
        print("\n\n  ⏹  Detenido por el usuario. CSV final:", CSV_SALIDA)
        return 0


if __name__ == "__main__":
    sys.exit(main())
