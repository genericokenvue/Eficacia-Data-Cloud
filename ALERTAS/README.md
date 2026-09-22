# ALERTAS — Sistema de cumplimiento (Telegram + Correo)

Independiente del ETL principal (`SCRIPTS/`). Calcula el cumplimiento de
cada gestor/supervisor y avisa por Telegram y/o correo. Se lanza completo
con el workflow `1C_alertas.yml` (ver `.github/workflows/README.md`).

## Las 3 fases (en orden)

| Fase | Script | Qué hace |
|---|---|---|
| A | `calcular_cumplimientos.py` | Calcula los KPIs de cumplimiento y los guarda en SharePoint. |
| B | `alertas_telegram.py` | Manda un mensaje por Telegram a cada supervisor con el resumen de su equipo. |
| C | `alertas_email.py` | Genera el Excel por supervisor y manda el correo (vía Microsoft Graph, sin Outlook). |

`run_alertas.py` corre las tres fases seguidas — es lo que ejecuta el workflow `1C_alertas.yml`.

## Utilidades

| Script | Para qué |
|---|---|
| `diagnostico_cruce.py` | Revisión PREVIA de solo lectura — no manda nada, avisa si algo va a fallar antes de correr en serio. Correr esto primero si algo se ve raro. |
| `reintentar_email.py` / `reintentar_telegram.py` | Reenvía solo a supervisores puntuales, sin repetir el envío completo. |
| `bot_listener.py` / `setup_telegram.py` | Capturan y mantienen los `chat_id` de Telegram de cada supervisor. |
| `base_cupos.py` | Resuelve identidades de personas entre las distintas fuentes (PT, D&P). |
| `cumplimiento_dyp.py` | Cálculo de los 4 KPIs comerciales D&P (Venta, Impactos, Clientes, etc.). |
| `config_loader.py` | Lee `config.env` (llaves de Telegram, correo remitente, etc.). |
| `alertas_logger.py` | Comparte el logger del ETL principal para que todo quede en el mismo log. |

## Dónde se ajustan las cosas que cambian seguido

- **Supervisores, correos y chat_id de Telegram:** misma `MAESTRO_SUPERVISORES.xlsx` que usa `envio_errores.py` (ver README de `SCRIPTS/`).
- **Llaves/tokens (Telegram, correo remitente):** archivo `config.env` en esta carpeta — no se sube a GitHub, vive solo en SharePoint/local.
- **Umbrales de cumplimiento:** fijos en `calcular_cumplimientos.py` y `cumplimiento_dyp.py`.

Antes de correr `run_alertas.py` en serio, si algo cambió recientemente, conviene
correr primero `diagnostico_cruce.py` para confirmar que los cruces de datos están bien.
