"""
capturar_respuestas_correo.py
───────────────────────────────
Reemplaza el paso que iba a hacer Power Automate: busca en la bandeja de
entrada de generico_kenvue@eficacia.com.co los correos "CAPTURA DE ERRORES"
que traen un Excel adjunto, guarda cada adjunto en
SALIDAS/RESPUESTAS_ERRORES (SharePoint), y de ahí en adelante
`aplicar_correcciones.py` hace su parte — este script NO aplica ninguna
corrección, solo captura y guarda el archivo.

POR QUÉ NO POWER AUTOMATE
──────────────────────────
El trigger "Cuando llega un nuevo correo" de Power Automate necesita crear
un enganche persistente con Exchange, y esa operación está fallando con
"Mailbox move in progress" mientras el buzón de generico_kenvue sigue en
migración. Este script hace lo mismo con simples consultas GET a Microsoft
Graph (la misma API que ya usan todos los ETL de este proyecto) — no crea
ningún recurso persistente, así que no debería tropezar con ese bloqueo.
Ver conversación del 2026-09-18.

CÓMO EVITA REPETIR O PERDER UN CORREO
───────────────────────────────────────
No se usa "leído/no leído": cualquiera puede marcar un correo como leído
sin querer (Outlook al abrirlo, el celular, etc.), y eso lo sacaría del
proceso sin haber guardado nada. En cambio:

  1. Se buscan correos SOLO en la Bandeja de entrada (no en Procesados).
  2. Se descarga el adjunto y se sube a RESPUESTAS_ERRORES.
  3. Recién SI el paso 2 salió bien, se MUEVE el correo a la carpeta
     "Procesados" del mismo buzón (se crea si no existe).

Si algo falla en el paso 2, el correo se queda tal cual en la Bandeja de
entrada — no se mueve — y la próxima corrida (cada 15-20 min, por cron) lo
vuelve a intentar sola, sin que nadie tenga que hacer nada. Un correo
fallido no frena a los demás: se sigue con el resto y se avisa cuál quedó
pendiente.

USO
───
    python capturar_respuestas_correo.py
"""
from __future__ import annotations

import base64
import datetime
import sys
import urllib.parse

import requests

import etl_espacios_errores as esp
import paths

BUZON = "generico_kenvue@eficacia.com.co"
ASUNTO_CLAVE = "CAPTURA DE ERRORES"
CARPETA_PROCESADOS = "Procesados"


def _graph_url(resto: str) -> str:
    return f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(BUZON)}{resto}"


def obtener_carpeta_id(headers: dict, nombre: str, padre: str = "inbox") -> str:
    """
    Id de la subcarpeta `nombre` dentro de `padre` — la crea si todavía no
    existe. Idempotente: si ya existe, Graph la devuelve tal cual, no la
    duplica.
    """
    url = _graph_url(f"/mailFolders/{padre}/childFolders")
    r = requests.get(url, headers=headers, params={"$filter": f"displayName eq '{nombre}'"})
    r.raise_for_status()
    existentes = r.json().get("value", [])
    if existentes:
        return existentes[0]["id"]

    r = requests.post(url, headers=headers, json={"displayName": nombre})
    r.raise_for_status()
    print(f"  ℹ️  Carpeta '{nombre}' creada dentro de la bandeja de entrada.")
    return r.json()["id"]


def buscar_correos_pendientes(headers: dict) -> list[dict]:
    """
    Correos en la BANDEJA DE ENTRADA con adjunto. El filtro de Graph solo
    revisa `hasAttachments` (filtrar texto parcial del asunto por OData es
    poco confiable); el 'CAPTURA DE ERRORES' se revisa acá, en Python.
    """
    url = _graph_url("/mailFolders/inbox/messages")
    params = {"$filter": "hasAttachments eq true", "$select": "id,subject,from", "$top": "50"}
    r = requests.get(url, headers=headers, params=params)
    r.raise_for_status()
    correos = r.json().get("value", [])
    return [c for c in correos if ASUNTO_CLAVE in (c.get("subject") or "").upper()]


def descargar_adjuntos_xlsx(headers: dict, mensaje_id: str) -> list[tuple[str, bytes]]:
    """Solo adjuntos de archivo (no imágenes/firmas) que terminen en .xlsx."""
    url = _graph_url(f"/messages/{mensaje_id}/attachments")
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    adjuntos = []
    for a in r.json().get("value", []):
        nombre = a.get("name", "")
        if a.get("@odata.type") == "#microsoft.graph.fileAttachment" and nombre.lower().endswith(".xlsx"):
            adjuntos.append((nombre, base64.b64decode(a["contentBytes"])))
    return adjuntos


def subir_adjunto_sharepoint(headers: dict, site_id: str, nombre: str, contenido: bytes) -> None:
    """
    Sube el adjunto tal cual a RESPUESTAS_ERRORES, con la fecha/hora
    adelante en el nombre para que dos respuestas no se pisen si llegan
    con el mismo nombre de archivo.
    """
    nombre_final = f"{datetime.datetime.now():%Y%m%d_%H%M%S}_{nombre}"
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(paths.RUTA_CARPETA_RESPUESTAS_ERRORES)}/"
           f"{urllib.parse.quote(nombre_final)}:/content")
    cab = {**headers,
           "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
    r = requests.put(url, headers=cab, data=contenido)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"No se pudo subir {nombre_final}: {r.status_code} - {r.text}")
    print(f"    ✅ Guardado en SharePoint: RESPUESTAS_ERRORES/{nombre_final}")


def mover_a_procesados(headers: dict, mensaje_id: str, carpeta_id: str) -> None:
    url = _graph_url(f"/messages/{mensaje_id}/move")
    r = requests.post(url, headers=headers, json={"destinationId": carpeta_id})
    if r.status_code not in (200, 201):
        raise RuntimeError(f"No se pudo mover el correo: {r.status_code} - {r.text}")


def run() -> int:
    print("\n" + "=" * 60)
    print("  CAPTURAR RESPUESTAS DE CORRECCIONES POR CORREO")
    print(f"  Buzón: {BUZON}")
    print("=" * 60)

    headers = {"Authorization": f"Bearer {esp.obtener_token_azure()}"}
    site_id = esp.obtener_site_id(headers)

    print(f"\nBuscando correos '{ASUNTO_CLAVE}' con adjunto en la bandeja de entrada...")
    correos = buscar_correos_pendientes(headers)
    print(f"  {len(correos)} correo(s) encontrado(s).")

    if not correos:
        return 0

    carpeta_procesados_id = obtener_carpeta_id(headers, CARPETA_PROCESADOS)

    guardados, fallidos = 0, 0
    for correo in correos:
        asunto = correo.get("subject", "")
        remitente = (correo.get("from") or {}).get("emailAddress", {}).get("address", "")
        print(f"\n  · {asunto!r} (de {remitente})")
        try:
            adjuntos = descargar_adjuntos_xlsx(headers, correo["id"])
            if not adjuntos:
                print("    ⚠️  No trae ningún adjunto .xlsx — se deja en la bandeja para revisar a mano.")
                continue
            for nombre, contenido in adjuntos:
                subir_adjunto_sharepoint(headers, site_id, nombre, contenido)
            mover_a_procesados(headers, correo["id"], carpeta_procesados_id)
            print("    📁 Correo movido a 'Procesados'.")
            guardados += 1
        except Exception as e:
            fallidos += 1
            print(f"    ❌ Este correo falló y se deja en la bandeja para reintentar: {e}")

    print(f"\n{'=' * 60}")
    print(f"  {guardados} correo(s) procesado(s) con éxito"
          f"{f', {fallidos} pendiente(s) por reintentar' if fallidos else ''}.")
    print("=" * 60)
    return 0


def main() -> int:
    return run()


if __name__ == "__main__":
    sys.exit(main())
