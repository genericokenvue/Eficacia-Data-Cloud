"""
alertas_email.py
────────────────
Fase C del sistema de alertas de Eficacia.

Genera los adjuntos Excel por supervisor y envía los correos directamente
vía Microsoft Graph API (sendMail), sin depender de Outlook Desktop, Excel
COM ni macros VBA — así corre igual en Windows local que en un runner de
GitHub Actions (Linux, sin GUI).

Flujo de trabajo
────────────────
1. Lee MAESTRO_SUPERVISORES.xlsx (SharePoint) para obtener correos.
2. Lee los adjuntos ya generados por calcular_cumplimientos.py
   (locales si el mismo proceso los generó en esta corrida —vía
   generar_adjuntos_por_supervisor—, o desde SharePoint como respaldo si
   se reutilizan de una corrida anterior, p.ej. con --solo email).
3. Por cada supervisor con correo configurado, arma el asunto + cuerpo
   HTML + adjunto y llama a Microsoft Graph `/users/{remitente}/sendMail`.

Antes esto lo hacía una macro VBA (EnviarCorreos.xlsm) disparada con
win32com + Outlook Desktop — sólo podía correr en una máquina Windows con
Outlook abierto y sesión activa. Ese mecanismo quedó reemplazado por
completo; EnviarCorreos.xlsm / EnviarCorreos.bas ya no se usan.

Prerequisitos
─────────────
  · Variables de entorno (mismo App Registration que usa paths.py para
    SharePoint): AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET.
  · El App Registration debe tener el permiso de aplicación (no delegado)
    "Mail.Send" en Microsoft Graph, con consentimiento de administrador.
  · CORREO_REMITENTE: la casilla (UPN, ej. alertas@eficaciacomco.com) desde
    la que se envían los correos. Con permisos de aplicación, Graph exige
    /users/{remitente}/sendMail (no existe "/me" en client credentials).
  · MAESTRO_SUPERVISORES.xlsx con CORREO completo (en SharePoint).
  · Adjuntos generados por calcular_cumplimientos.py.
"""

import os
import sys
import io
import time
import base64
import urllib.parse
import pandas as pd
import requests
from pathlib import Path
from datetime import datetime

# Cargar config.env antes de leer UMBRAL_OK/WARNING desde os.environ
from config_loader import cargar_config
cargar_config()

from alertas_logger import get_logger
_log = get_logger("alertas_email")

import paths   # alertas_logger ya añadió SCRIPTS/ a sys.path
BASE         = paths.BASE
DIR_ALERTAS  = paths.ALERTAS_DIR
DIR_ADJUNTOS = paths.ALERTAS_ADJUNTOS
RUTA_MAESTRA = paths.ALERTAS_MAESTRO

UMBRAL_OK      = float(os.environ.get("UMBRAL_OK",      "0.90"))

# Contacto del analista para reportar inconsistencias (Sprint 13.3)
ANALISTA_NOMBRE = os.environ.get("ANALISTA_NOMBRE", "Giovanny Restrepo")
ANALISTA_EMAIL  = os.environ.get("ANALISTA_EMAIL",  "giovanny_restrepo@eficacia.com.co")
UMBRAL_WARNING = float(os.environ.get("UMBRAL_WARNING",  "0.70"))

# Casilla remitente para Graph sendMail (requiere permiso de aplicación
# Mail.Send sobre este buzón, con consentimiento de administrador).
CORREO_REMITENTE = os.environ.get("CORREO_REMITENTE", "").strip()

# Carpeta en la nube donde calcular_cumplimientos.py sube los adjuntos por
# supervisor (respaldo cuando no hay copia local en disco, p.ej. --solo email).
RUTA_CARPETA_ADJUNTOS_CLOUD = f"{paths._SALIDAS_ROOT}/ALERTAS/ADJUNTOS"

# Maestra de supervisores (nombre, correo, chat_id) — vive en BASES DE
# RESPUESTAS/ALERTAS, no en SALIDAS/ALERTAS (donde vive el resto del
# output del pipeline). A nivel de módulo para que scripts standalone como
# reintentar_email.py puedan importarla directamente.
RUTA_MAESTRA_CLOUD = f"{paths._BASES_ROOT}/ALERTAS/MAESTRO_SUPERVISORES.xlsx"


# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCCIÓN DEL CUERPO HTML DEL CORREO
# ─────────────────────────────────────────────────────────────────────────────

def _umbral_ok_de(col: str | None) -> float:
    """Resuelve el UMBRAL_OK específico de la columna (Sprint 13.4)."""
    if not col:
        return UMBRAL_OK
    try:
        from calcular_cumplimientos import umbral_de
        return umbral_de(col)
    except Exception:
        return UMBRAL_OK


def _color_celda(valor, col: str | None = None) -> str:
    """Color de fondo HTML según umbral específico del KPI (Sprint 13.4)."""
    if pd.isna(valor):
        return "#F2F2F2"
    umb_ok = _umbral_ok_de(col)
    if valor >= umb_ok:
        return "#C6EFCE"
    if valor >= UMBRAL_WARNING:
        return "#FFEB9C"
    return "#FFC7CE"


def _semaforo(valor, col: str | None = None) -> str:
    if pd.isna(valor):
        return "—"
    umb_ok = _umbral_ok_de(col)
    if valor >= umb_ok:
        return "✅"
    if valor >= UMBRAL_WARNING:
        return "⚠️"
    return "❌"


def _pct(valor) -> str:
    if pd.isna(valor):
        return "N/D"
    return f"{valor * 100:.1f}%"


def _label_global_rol(rol: str) -> str:
    """Sprint 17.15 — label del cumplimiento global según rol."""
    if rol == "GDD":
        return "Cumplimiento global del periodo"
    if rol == "LIDER":
        return "Cumplimiento global de tu zona"
    return "Cumplimiento global del equipo"


def _intro_rol(rol: str, n_gestores: int) -> str:
    """Sprint 17.15 — saludo según rol del destinatario."""
    if rol == "GDD":
        return "A continuación encontrará su <strong>cumplimiento personal</strong>"
    if rol == "LIDER":
        if n_gestores < 1:
            return "A continuación encontrará su <strong>cumplimiento personal</strong>"
        return (
            f"A continuación encontrará el resumen de cumplimiento "
            f"de los <strong>{n_gestores} supervisores</strong> a su cargo "
            f"y su cumplimiento personal"
        )
    # SUPERVISOR (default)
    if n_gestores == 0:
        return "A continuación encontrará su <strong>cumplimiento personal</strong>"
    # Restamos 1 por la fila del propio supervisor cuando aplica.
    n_mercs = max(n_gestores - 1, 1)
    return (
        f"A continuación encontrará el resumen de cumplimiento de su equipo de "
        f"<strong>{n_mercs} mercaderistas</strong>"
    )


def construir_cuerpo_html(
    supervisor: str,
    df_equipo: pd.DataFrame,
    mes: int,
    anio: int,
    rango_periodo: dict | None = None,
    rol_destinatario: str = "SUPERVISOR",
) -> str:
    """
    Genera el HTML del cuerpo del correo para un supervisor.
    Incluye tabla de resumen del equipo y nota sobre el adjunto.

    rango_periodo (Sprint 13.1): dict de detectar_rango_periodo() con
    fecha_inicio/fecha_fin/rango_legible/avance_esperado_pct. Si None,
    el correo solo muestra mes/año sin rango específico.
    """
    meses_es = {
        1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
        5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
        9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre",
    }
    mes_nombre = meses_es.get(mes, str(mes))
    rango_periodo = rango_periodo or {}
    rango_legible = rango_periodo.get("rango_legible", "")
    avance_esperado = rango_periodo.get("avance_esperado_pct", 0.0)
    dias_trans = rango_periodo.get("dias_transcurridos", 0)
    dias_total = rango_periodo.get("dias_mes_total", 0)

    # Tabla HTML con los mercaderistas del equipo.
    # Sprint 17.12 — la columna "Exh. Pagadas" muestra CAPTURA (no ejecución).
    # Sprint 17.13 — la columna "Exh. Gratis" pasa de conteo (#) a % cumplimiento
    # promedio = (CUMP_ALTO + CUMP_MEDIO) / 2, con targets por canal del gestor
    # (PROXIMITY/TAT → 3 alto + 5 medio; resto → 2 alto + 8 medio).
    cols_modulos = {
        # Operativos
        "CIF_%"                  : "CIF<br>(Visitas)",
        "NP_%"                   : "No<br>Presencia",
        "PRECIOS_%"              : "Precios",
        "SOS_%"                  : "SOS",
        "EXHIB_PAG_CAPTURA_%"    : "Exh.<br>Pagadas",
        # Comerciales D&P
        "VENTA_%"                : "Venta<br>vs Cuota",
        "IMPACTOS_%"             : "Impactos",
        "MSL_%"                  : "MSL",
        "PROD_NUEVOS_%"          : "Productos<br>Nuevos",
        "EXHIB_GRATIS_PROM_%"    : "Exh.<br>Gratis",
        "CUMPL_GLOBAL_%"         : "Global",
    }
    if rol_destinatario == "LIDER":
        for _c in ("VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"):
            cols_modulos.pop(_c, None)

    # Sprint 17.13 — derivar EXHIB_GRATIS_PROM_% = (CUMP_ALTO + CUMP_MEDIO)/2.
    # Sprint 17.14 — cada componente se capa a 1.0 (100%) antes del promedio
    # para evitar que outliers (gestores con conteos anómalos) inflen el global.
    if {"EXHIB_GRA_ALTO_%", "EXHIB_GRA_MEDIO_%"}.issubset(df_equipo.columns):
        df_equipo = df_equipo.copy()
        alto_cap  = df_equipo["EXHIB_GRA_ALTO_%"].clip(upper=1.0)
        medio_cap = df_equipo["EXHIB_GRA_MEDIO_%"].clip(upper=1.0)
        df_equipo["EXHIB_GRATIS_PROM_%"] = pd.concat(
            [alto_cap, medio_cap], axis=1
        ).mean(axis=1, skipna=True)

    # CUMPL_GLOBAL: SIEMPRE recalcular para que use EXHIB_GRATIS_PROM_% capeado
    # (no los componentes ALTO/MEDIO sin techo, ni el EXHIB_PAG_% de ejecución
    # que ya no se reporta en el HTML).
    cols_kpi = ["CIF_%", "NP_%", "PRECIOS_%", "SOS_%", "EXHIB_PAG_CAPTURA_%",
                "EXHIB_GRATIS_PROM_%"]
    if rol_destinatario != "LIDER":
        cols_kpi += ["VENTA_%", "IMPACTOS_%", "MSL_%", "PROD_NUEVOS_%"]
    cols_disp = [c for c in cols_kpi if c in df_equipo.columns]
    if cols_disp:
        df_equipo = df_equipo.copy()
        df_equipo["CUMPL_GLOBAL_%"] = df_equipo[cols_disp].mean(axis=1, skipna=True)

    filas_html = ""
    for i, (_, row) in enumerate(df_equipo.iterrows()):
        bg = "#FFFFFF" if i % 2 == 0 else "#F7F7F7"
        nombre = row.get("NOMBRE", "—")
        filas_html += f'<tr style="background:{bg}">\n'
        filas_html += f'  <td style="padding:6px 10px;border:1px solid #ddd">{nombre}</td>\n'
        for col in cols_modulos:
            if col in row.index:
                val = row[col]
                # Semáforo y color según umbral del KPI (todas las cols ahora son %).
                color = _color_celda(val, col)
                texto = f"{_pct(val)} {_semaforo(val, col)}"
                filas_html += (
                    f'  <td style="padding:6px 10px;border:1px solid #ddd;'
                    f'background:{color};text-align:center">{texto}</td>\n'
                )
        filas_html += "</tr>\n"

    encabezados_html = '<th style="padding:8px 10px;background:#1F3864;color:#FFF;border:1px solid #ddd">Mercaderista</th>\n'
    for etiqueta in cols_modulos.values():
        encabezados_html += (
            f'<th style="padding:8px 10px;background:#1F3864;color:#FFF;'
            f'border:1px solid #ddd;text-align:center">{etiqueta}</th>\n'
        )

    n_gestores = len(df_equipo)
    global_equipo_val = df_equipo["CUMPL_GLOBAL_%"].mean() if "CUMPL_GLOBAL_%" in df_equipo.columns else float("nan")
    global_equipo_str = _pct(global_equipo_val)
    global_sem        = _semaforo(global_equipo_val, "CUMPL_GLOBAL_%")

    html = f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin:0;padding:0">

  <div style="background:#1F3864;padding:18px 24px">
    <h2 style="margin:0;color:#FFFFFF;font-size:16px">
      📊 Reporte de Cumplimiento — {mes_nombre} {anio}
    </h2>
    {f'<p style="margin:4px 0 0;color:#C5D4E8;font-size:12px">Corte del {rango_legible} — día {dias_trans} de {dias_total} ({int(avance_esperado*100)}% del mes transcurrido)</p>' if rango_legible else '<p style="margin:4px 0 0;color:#C5D4E8;font-size:12px">Eficacia</p>'}
  </div>

  <div style="padding:20px 24px">

    <p>Estimada/o <strong>{supervisor.title()}</strong>,</p>
    <p>{_intro_rol(rol_destinatario, n_gestores)} {f'para el periodo del <strong>{rango_legible} de {anio}</strong>' if rango_legible else f'para el periodo <strong>{mes_nombre} {anio}</strong>'}.</p>

    <p><strong>{_label_global_rol(rol_destinatario)}: {global_equipo_str} {global_sem}</strong></p>

    <table style="border-collapse:collapse;width:100%;margin-top:12px">
      <thead>
        <tr>{encabezados_html}</tr>
      </thead>
      <tbody>
        {filas_html}
      </tbody>
    </table>

    <p style="margin-top:20px;font-size:12px;color:#666">
      El archivo adjunto contiene el detalle completo por punto de venta.<br>
      Ante cualquier inconsistencia, comunícate con <strong>{ANALISTA_NOMBRE}</strong> &mdash;
      <a href="mailto:{ANALISTA_EMAIL}" style="color:#1F3864;text-decoration:underline">{ANALISTA_EMAIL}</a>.
    </p>

  </div>

  <div style="background:#F2F2F2;padding:10px 24px;font-size:11px;color:#888">
    Eficacia S.A. — Reporte generado automáticamente el {datetime.now().strftime("%d/%m/%Y %H:%M")}
  </div>

</body>
</html>
"""
    return html


# ─────────────────────────────────────────────────────────────────────────────
# CONEXIÓN A LA NUBE (SharePoint + envío por Microsoft Graph API)
# ─────────────────────────────────────────────────────────────────────────────
_DRIVE_ID_CACHE: dict = {}


def _obtener_default_drive_id(token: str) -> str:
    """Resuelve el Drive ID del sitio SharePoint configurado en paths.SHAREPOINT_SITE_NAME."""
    if "drive_id" in _DRIVE_ID_CACHE:
        return _DRIVE_ID_CACHE["drive_id"]

    headers = {"Authorization": f"Bearer {token}"}
    nombre_sitio = getattr(paths, "SHAREPOINT_SITE_NAME", "eficacia")

    res_site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites?search={urllib.parse.quote(nombre_sitio)}",
        headers=headers,
    )
    if res_site.status_code == 200:
        sites = res_site.json().get("value", [])
        if sites:
            site = next(
                (s for s in sites
                 if nombre_sitio.lower() in (s.get("name", "").lower(), s.get("displayName", "").lower())),
                sites[0],
            )
            res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site.get('id')}/drive", headers=headers)
            if res_drive.status_code == 200:
                _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
                return _DRIVE_ID_CACHE["drive_id"]

    res_root = requests.get("https://graph.microsoft.com/v1.0/sites/root", headers=headers)
    if res_root.status_code == 200:
        site_id = res_root.json().get("id")
        res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
        if res_drive.status_code == 200:
            _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
            return _DRIVE_ID_CACHE["drive_id"]

    raise Exception(f"No se pudo obtener el Drive ID del sitio SharePoint '{nombre_sitio}': {res_site.text}")


def _limpiar_ruta_graph(ruta_sharepoint: str) -> str:
    """Normaliza separadores. NO remueve 'Equipo Información/' — es una carpeta real del sitio."""
    ruta_sharepoint = ruta_sharepoint.replace("\\", "/")
    for prefijo in ["Documentos compartidos/", "Shared Documents/"]:
        if ruta_sharepoint.startswith(prefijo):
            ruta_sharepoint = ruta_sharepoint.replace(prefijo, "", 1)
    return ruta_sharepoint.lstrip("/")


def _descargar_bytes_cloud(ruta_sharepoint: str, descripcion: str) -> bytes:
    """Descarga el contenido crudo (bytes) de un archivo en SharePoint vía Graph API."""
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    rutas_candidatas = [ruta_limpia, f"Documentos compartidos/{ruta_limpia}"]

    response = None
    for r in rutas_candidatas:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(r)}:/content"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.content

    status = response.status_code if response is not None else "N/A"
    text = response.text if response is not None else "Sin respuesta"
    raise FileNotFoundError(f"No se pudo descargar {descripcion} desde SharePoint ({ruta_sharepoint}): {status} - {text}")


def _leer_excel_cloud(ruta_sharepoint: str, descripcion: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(_descargar_bytes_cloud(ruta_sharepoint, descripcion)))


def _obtener_bytes_adjunto(nombre_sup: str, ruta_o_valor: str, nombre_archivo: str) -> bytes | None:
    """
    Devuelve los bytes del adjunto del supervisor. Si `calcular_cumplimientos`
    corrió en este mismo proceso (caso normal de run_alertas.py), el archivo
    sigue en el temporal local y se lee directo de disco. Si no existe (p.ej.
    `--solo email` reutilizando una corrida anterior), se descarga como
    respaldo desde SharePoint (RUTA_CARPETA_ADJUNTOS_CLOUD).
    """
    if ruta_o_valor:
        p = Path(ruta_o_valor)
        if p.is_file():
            return p.read_bytes()
    try:
        return _descargar_bytes_cloud(f"{RUTA_CARPETA_ADJUNTOS_CLOUD}/{nombre_archivo}", f"adjunto de {nombre_sup}")
    except Exception as e:
        _log.warning(f"No se pudo obtener el adjunto de {nombre_sup} (ni local ni en la nube): {e}")
        return None


def _validar_remitente() -> bool:
    """
    Intenta, ANTES de enviar nada, resolver CORREO_REMITENTE contra Graph
    (GET /users/{id}) solo para dejar constancia en el log de qué valor se
    está usando — el error real de sendMail ("ErrorInvalidUser") enmascara
    el valor con '***', así que esto lo muestra explícito.

    IMPORTANTE: este chequeo usa GET /users/{id}, que requiere el permiso
    de aplicación User.Read.All — un permiso DISTINTO a Mail.Send (el que
    de verdad se necesita para enviar correos). Si la app solo tiene
    Mail.Send (que es lo correcto/mínimo necesario), este GET puede fallar
    con 403 aunque el envío real funcione bien — por eso NUNCA bloquea el
    envío, es solo diagnóstico informativo.
    """
    print(f"  ℹ️  CORREO_REMITENTE configurado: '{CORREO_REMITENTE}'")
    try:
        token = paths.obtener_token_azure()
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(CORREO_REMITENTE)}"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            print(f"  ✓ Graph reconoce el usuario: {data.get('displayName')} ({data.get('userPrincipalName')})")
        elif response.status_code == 403:
            print("  ℹ️  No se pudo verificar el remitente con GET /users (falta el permiso "
                  "User.Read.All, distinto a Mail.Send) — no es necesariamente un problema, "
                  "se continúa e intenta el envío real igual.")
        else:
            print(f"  ⚠️  Graph no pudo resolver CORREO_REMITENTE='{CORREO_REMITENTE}' "
                  f"(status {response.status_code}): {response.text}")
            print("     Si ese correo SÍ existe en Outlook/M365, la causa más común de un 404 "
                  "aquí (ErrorInvalidUser) es una 'Application Access Policy' en Exchange Online "
                  "que no incluye este buzón en el scope permitido para la app — se revisa del "
                  "lado de Exchange (PowerShell), no en este código. Se continúa igual e intenta "
                  "el envío real, que es la prueba definitiva.")
    except Exception as e:
        print(f"  ⚠️  No se pudo pre-validar CORREO_REMITENTE: {e}")
    return True  # Nunca bloquea — es solo diagnóstico. El intento real de sendMail manda.


def _enviar_correo_graph(destinatario: str, asunto: str, cuerpo_html: str,
                          nombre_adjunto: str | None, bytes_adjunto: bytes | None) -> None:
    """Envía un correo vía Microsoft Graph `/users/{CORREO_REMITENTE}/sendMail`."""
    if not CORREO_REMITENTE:
        raise RuntimeError(
            "Falta la variable de entorno CORREO_REMITENTE (la casilla desde la que "
            "se envían los correos vía Graph — requiere permiso de aplicación Mail.Send)."
        )

    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    mensaje = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": cuerpo_html},
        "toRecipients": [{"emailAddress": {"address": destinatario}}],
    }
    if bytes_adjunto and nombre_adjunto:
        mensaje["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": nombre_adjunto,
            "contentBytes": base64.b64encode(bytes_adjunto).decode(),
        }]

    url = f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(CORREO_REMITENTE)}/sendMail"
    response = requests.post(url, headers=headers, json={"message": mensaje, "saveToSentItems": "true"})
    if response.status_code not in (200, 202):
        raise Exception(f"Graph sendMail devolvió {response.status_code}: {response.text}")


# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCCIÓN DE LA COLA DE ENVÍOS (en memoria, ya no en un .xlsm)
# ─────────────────────────────────────────────────────────────────────────────

def construir_cola_envios(
    df_detalle: pd.DataFrame,
    df_maestro: pd.DataFrame,
    rutas_adjuntos: dict,
    mes: int,
    anio: int,
    rango_periodo: dict | None = None,
) -> list[dict]:
    """
    Arma la lista de correos a enviar (uno por supervisor/GDD/líder con
    correo configurado). Cada elemento trae CORREO, ASUNTO, CUERPO_HTML,
    NOMBRE_ADJUNTO y RUTA_ADJUNTO (local o vacío).
    """
    # Normalizar maestra
    df_m = df_maestro.copy()
    df_m.columns = df_m.columns.str.strip().str.upper()
    # IMPORTANTE: fillna("") ANTES de astype(str). En versiones recientes de
    # pandas (dtype "str" nativo), astype(str) sobre una columna con NaN NO
    # convierte esos valores a la palabra "nan" — se quedan como NaN real
    # (float), y el filtro de abajo (comparar contra "nan"/"NAN") nunca los
    # atrapaba. Por eso se estaban armando envíos con CORREO=NaN, que al
    # imprimirse en el f-string se veía como el texto "nan".
    df_m["NOMBRE_SUPERVISOR"] = df_m["NOMBRE_SUPERVISOR"].fillna("").astype(str).str.strip().str.upper()
    df_m["CORREO"]            = df_m["CORREO"].fillna("").astype(str).str.strip()

    meses_es = {
        1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
        5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
        9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre",
    }
    mes_nombre = meses_es.get(mes, str(mes))
    if rango_periodo and rango_periodo.get("rango_legible"):
        asunto_base = (
            f"Cumplimiento {mes_nombre} {anio} "
            f"(corte {rango_periodo['rango_legible']}) — Tu equipo Eficacia"
        )
    else:
        asunto_base = f"Cumplimiento {mes_nombre} {anio} — Tu equipo Eficacia"

    filas = []
    sin_correo = []

    # Sprint 17.15 — destinatarios: SUPERVISOR + GDD + LIDER. Cada uno recibe correo propio.
    if {"ES_SUPERVISOR", "ES_GDD", "ES_LIDER"}.issubset(df_detalle.columns):
        sups_df = (
            df_detalle[
                (df_detalle["ES_SUPERVISOR"] == True) |
                (df_detalle["ES_GDD"]        == True) |
                (df_detalle["ES_LIDER"]      == True)
            ][["ACRONIMO", "NOMBRE"]]
                     .drop_duplicates(subset=["ACRONIMO"])
                     .sort_values("NOMBRE")
        )
    elif "ES_SUPERVISOR" in df_detalle.columns:
        sups_df = (
            df_detalle[df_detalle["ES_SUPERVISOR"] == True][["ACRONIMO", "NOMBRE"]]
                     .drop_duplicates(subset=["ACRONIMO"])
                     .sort_values("NOMBRE")
        )
    else:
        sups_df = pd.DataFrame({
            "ACRONIMO": [""] * df_detalle["SUPERVISOR_LIDER"].nunique(),
            "NOMBRE":   sorted(df_detalle["SUPERVISOR_LIDER"].dropna().unique()),
        })

    try:
        from calcular_cumplimientos import (
            LIDER_POR_CIUDAD as _LIDER_POR_CIUDAD,
            construir_equipo_lider as _construir_equipo_lider,
        )
    except ImportError:
        _LIDER_POR_CIUDAD = {}
        _construir_equipo_lider = None

    for _, fila_sup in sups_df.iterrows():
        acr_sup    = fila_sup["ACRONIMO"]
        nombre_sup = str(fila_sup["NOMBRE"]).strip().upper()
        if not nombre_sup:
            continue

        correo_row = df_m[df_m["NOMBRE_SUPERVISOR"] == nombre_sup]
        if correo_row.empty or correo_row["CORREO"].iloc[0] in ("", "nan", "NAN"):
            sin_correo.append(nombre_sup)
            continue
        correo = correo_row["CORREO"].iloc[0]

        propio = df_detalle[df_detalle["ACRONIMO"] == acr_sup]
        es_gdd   = bool(propio["ES_GDD"].iloc[0])   if not propio.empty and "ES_GDD"   in propio.columns else False
        es_lider = bool(propio["ES_LIDER"].iloc[0]) if not propio.empty and "ES_LIDER" in propio.columns else False
        if es_gdd:
            df_equipo = propio.copy()
        elif es_lider and _construir_equipo_lider is not None:
            df_equipo = _construir_equipo_lider(df_detalle, acr_sup)
        elif "ACRONIMO_SUP" in df_detalle.columns and acr_sup:
            gestores = df_detalle[df_detalle["ACRONIMO_SUP"] == acr_sup]
            df_equipo = pd.concat([gestores, propio], ignore_index=True)
        else:
            df_equipo = df_detalle[df_detalle.get("SUPERVISOR_LIDER", "") == nombre_sup].copy()

        cuerpo = construir_cuerpo_html(
            nombre_sup, df_equipo, mes, anio,
            rango_periodo=rango_periodo,
            rol_destinatario=("GDD" if es_gdd else ("LIDER" if es_lider else "SUPERVISOR")),
        )
        ruta_adjunto   = rutas_adjuntos.get(nombre_sup, "")
        nombre_adjunto = f"Detalle_{nombre_sup.replace(' ', '_')}_{mes:02d}_{anio}.xlsx"

        filas.append({
            "CORREO"        : correo,
            "ASUNTO"        : asunto_base,
            "CUERPO_HTML"   : cuerpo,
            "NOMBRE_SUP"    : nombre_sup,
            "RUTA_ADJUNTO"  : ruta_adjunto,
            "NOMBRE_ADJUNTO": nombre_adjunto,
        })

    if sin_correo:
        print(f"  ⚠️  {len(sin_correo)} supervisor(es) sin correo en la maestra:")
        for s in sin_correo:
            print(f"     · {s}")

    print(f"  ✓ Cola armada: {len(filas)} correos")
    return filas


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL EXPORTADA
# ─────────────────────────────────────────────────────────────────────────────

def enviar_correos(
    df_detalle: pd.DataFrame,
    df_maestro: pd.DataFrame,
    rutas_adjuntos: dict,
    mes: int,
    anio: int,
    modo_prueba: bool = False,
    rango_periodo: dict | None = None,
) -> dict:
    """
    Arma la cola de correos y los envía uno a uno vía Microsoft Graph
    (`/users/{CORREO_REMITENTE}/sendMail`). Corre igual en local que en
    GitHub Actions — no depende de Outlook Desktop ni de Excel.

    Retorna dict con 'filas_cola', 'enviados', 'errores', 'macro_ok'
    (esta última clave se conserva por compatibilidad — True si no hubo
    errores de envío).
    """
    print("\n" + "═" * 55)
    print("  ALERTAS EMAIL — Eficacia (Graph API)")
    print("═" * 55)

    if modo_prueba:
        print("  ⚠️  MODO PRUEBA — se arma la cola pero no se envía nada real\n")
    elif not CORREO_REMITENTE:
        print("  ❌ Falta CORREO_REMITENTE en las variables de entorno.")
        return {"filas_cola": 0, "enviados": 0, "errores": 0, "macro_ok": False}
    elif not _validar_remitente():
        return {"filas_cola": 0, "enviados": 0, "errores": 0, "macro_ok": False}

    print("\nArmando cola de envíos:")
    filas = construir_cola_envios(df_detalle, df_maestro, rutas_adjuntos, mes, anio, rango_periodo=rango_periodo)

    if not filas:
        print("  ⚠️  Cola vacía — no hay supervisores con correo configurado.")
        return {"filas_cola": 0, "enviados": 0, "errores": 0, "macro_ok": False}

    print("\nEnviando correos vía Microsoft Graph:")
    enviados = errores = 0
    for fila in filas:
        if modo_prueba:
            print(f"  🧪 [PRUEBA] Se enviaría a {fila['CORREO']} — adjunto: {fila['NOMBRE_ADJUNTO']}")
            enviados += 1
            continue
        try:
            bytes_adj = _obtener_bytes_adjunto(fila["NOMBRE_SUP"], fila["RUTA_ADJUNTO"], fila["NOMBRE_ADJUNTO"])
            _enviar_correo_graph(
                fila["CORREO"], fila["ASUNTO"], fila["CUERPO_HTML"],
                fila["NOMBRE_ADJUNTO"], bytes_adj,
            )
            print(f"  ✓ Enviado a {fila['CORREO']} ({fila['NOMBRE_SUP']})")
            enviados += 1
        except Exception as e:
            print(f"  ❌ Error enviando a {fila['CORREO']} ({fila['NOMBRE_SUP']}): {e}")
            _log.error(f"Fallo enviando correo a {fila['CORREO']} ({fila['NOMBRE_SUP']}): {e}")
            errores += 1
        time.sleep(0.3)  # margen frente a throttling de Graph

    macro_ok = errores == 0
    _log.info(f"Email — filas_cola={len(filas)} enviados={enviados} errores={errores} prueba={modo_prueba}")

    print("\n" + "═" * 55)
    print(f"  Correos en cola: {len(filas)}")
    print(f"  Enviados: {enviados}  |  Errores: {errores}")
    print("═" * 55 + "\n")

    return {"filas_cola": len(filas), "enviados": enviados, "errores": errores, "macro_ok": macro_ok}


# ─────────────────────────────────────────────────────────────────────────────
# EJECUCIÓN STANDALONE
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Envío de correos vía Microsoft Graph — Eficacia")
    parser.add_argument("--prueba", action="store_true",
                        help="Modo prueba: arma la cola sin enviar nada real")
    args = parser.parse_args()

    # Sprint 17.11 — invocar calcular_cumplimientos.main() para obtener TODO
    # en memoria (incluye rango_periodo con fechas legibles).
    import calcular_cumplimientos as cc
    out = cc.main()
    df_detalle     = out["df_detalle"]
    mes            = out["mes"]
    anio           = out["anio"]
    rutas_adj      = out["rutas_adjuntos"]
    rango_periodo  = out.get("rango_periodo")

    # Maestra de supervisores: en la nube (mismo lugar donde
    # calcular_cumplimientos.py la guarda).
    try:
        df_maestro = _leer_excel_cloud(RUTA_MAESTRA_CLOUD, "Maestra de supervisores")
    except Exception as e:
        print(f"❌ No se pudo leer la maestra de supervisores desde SharePoint: {e}")
        sys.exit(1)
    if df_maestro.empty:
        print("❌ MAESTRO_SUPERVISORES.xlsx está vacío o no se encontró.")
        sys.exit(1)

    enviar_correos(
        df_detalle, df_maestro, rutas_adj, mes, anio,
        modo_prueba=args.prueba,
        rango_periodo=rango_periodo,
    )