# Workflows

Todos corren manual desde la pestaña **Actions** de GitHub (botón "Run workflow"). Ninguno tiene horario automático activo por ahora.

## Parte 1 — Datos (independientes entre sí)

| Workflow | Qué hace |
|---|---|
| `1A_run_etl_manual.yml` | Corre un ETL suelto a mano (precios, SOS, CIF, exhibiciones, etc.) o todos juntos — para backfill o probar un módulo sin esperar al diario. |
| `1B_run_etl_diario.yml` | Corre todos los ETLs del periodo actual y regenera GOLD al final — es la versión "sin botones", pensada para el cron diario (hoy desactivado, se lanza a mano). |
| `1C_alertas.yml` | Sistema de alertas de cumplimiento (Telegram y correo) — calcula cumplimientos y avisa a los supervisores. No depende de los demás workflows. |

## Parte 2 — Ciclo de corrección de errores (secuencia estricta, uno depende del anterior)

| Orden | Workflow | Qué hace |
|---|---|---|
| 1 | `2A_calculo_errores.yml` | Detecta errores de precios, exhibiciones y espacios (SOS) para el periodo, y deja un archivo por módulo en SharePoint con las columnas para corregir. |
| 2 | `2B_envio_errores.yml` | Le manda a cada supervisor un correo con SUS errores (de los 3 módulos juntos) y el archivo adjunto para corregir. |
| 3 | `2C_capturar_respuestas_correo.yml` | Revisa la bandeja de entrada, busca las respuestas de los supervisores ("CAPTURA DE ERRORES" + adjunto) y guarda cada adjunto en SharePoint. Reemplaza lo que iba a hacer Power Automate. |
| 4 | `2D_aplicar_correcciones.yml` | Lee lo que dejó el paso anterior y escribe cada corrección en el archivo real (la encuesta cruda de cada módulo), para que sobreviva al próximo ETL. |

**Por qué en ese orden:** el 2B necesita que el 2A ya haya calculado los errores; el 2D necesita que el 2C ya haya guardado las respuestas. Correrlos fuera de orden no rompe nada, pero no va a encontrar nada nuevo que procesar.
