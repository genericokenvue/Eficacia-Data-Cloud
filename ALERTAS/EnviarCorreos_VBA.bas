''
'' EnviarCorreos.xlsm  —  Módulo VBA: EnviarTodos
'' ──────────────────────────────────────────────
'' Instrucciones de instalación:
''   1. Crea un nuevo archivo Excel, guárdalo como EnviarCorreos.xlsm
''      (tipo "Libro de Excel habilitado para macros") en la carpeta ALERTAS/.
''   2. Abre el editor VBA (Alt+F11).
''   3. Inserta un módulo nuevo (Insertar → Módulo).
''   4. Pega todo el código de este archivo en ese módulo.
''   5. Cierra el editor y guarda el archivo.
''
'' La hoja COLA es escrita por alertas_email.py antes de llamar a esta macro.
'' Columnas de la hoja COLA:
''   A: CORREO | B: ASUNTO | C: CUERPO_HTML | D: RUTA_ADJUNTO | E: ESTADO
''

Option Explicit

Sub EnviarTodos()
    ''
    '' Itera la hoja COLA y envía un correo por cada fila con ESTADO = "PENDIENTE".
    '' Actualiza la columna ESTADO a "ENVIADO" o "ERROR: <detalle>" por fila.
    ''

    Dim ws          As Worksheet
    Dim oApp        As Object
    Dim oMail       As Object
    Dim oAttach     As Object
    Dim ultimaFila  As Long
    Dim i           As Long
    Dim correo      As String
    Dim asunto      As String
    Dim cuerpo      As String
    Dim adjunto     As String
    Dim estado      As String
    Dim errMsg      As String

    ' Verificar que la hoja COLA existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("COLA")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No se encontró la hoja 'COLA' en este libro." & vbCrLf & _
               "Ejecuta primero alertas_email.py para generar la cola.", _
               vbCritical, "Error — EnviarCorreos"
        Exit Sub
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If ultimaFila < 2 Then
        MsgBox "La hoja COLA está vacía (sin filas de datos).", _
               vbInformation, "Cola vacía"
        Exit Sub
    End If

    ' Crear instancia de Outlook (usa la sesión activa)
    On Error GoTo ErrorOutlook
    Set oApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    Dim enviados    As Long
    Dim errores     As Long
    enviados = 0
    errores  = 0

    ' Iterar filas desde la 2 (fila 1 = encabezados)
    For i = 2 To ultimaFila

        estado  = Trim(ws.Cells(i, 5).Value)   ' columna E: ESTADO

        ' Solo procesar pendientes (permite re-ejecutar sin duplicar)
        If UCase(estado) <> "PENDIENTE" Then
            GoTo SiguienteFila
        End If

        correo  = Trim(ws.Cells(i, 1).Value)   ' columna A: CORREO
        asunto  = Trim(ws.Cells(i, 2).Value)   ' columna B: ASUNTO
        cuerpo  = ws.Cells(i, 3).Value          ' columna C: CUERPO_HTML
        adjunto = Trim(ws.Cells(i, 4).Value)   ' columna D: RUTA_ADJUNTO

        ' Validar que el correo no esté vacío
        If correo = "" Or correo = "nan" Then
            ws.Cells(i, 5).Value = "ERROR: correo vacío"
            errores = errores + 1
            GoTo SiguienteFila
        End If

        ' Crear y configurar el correo
        On Error GoTo ErrorCorreo
        Set oMail = oApp.CreateItem(0)   ' 0 = olMailItem

        With oMail
            .To       = correo
            .Subject  = asunto
            .HTMLBody = cuerpo

            ' Adjuntar archivo si la ruta existe
            If adjunto <> "" And adjunto <> "nan" Then
                If Dir(adjunto) <> "" Then
                    .Attachments.Add adjunto
                Else
                    ' El adjunto no existe pero el correo se envía igual
                    ' con nota en ESTADO
                    adjunto = "ADJUNTO_NO_ENCONTRADO"
                End If
            End If

            .Send
        End With

        ' Marcar como enviado
        If adjunto = "ADJUNTO_NO_ENCONTRADO" Then
            ws.Cells(i, 5).Value = "ENVIADO (sin adjunto)"
        Else
            ws.Cells(i, 5).Value = "ENVIADO"
        End If
        enviados = enviados + 1

        Set oMail = Nothing

        ' Pequeña pausa entre envíos para no saturar el servidor SMTP
        Application.Wait Now + TimeValue("00:00:01")

        GoTo SiguienteFila

ErrorCorreo:
        errMsg = Err.Description
        ws.Cells(i, 5).Value = "ERROR: " & Left(errMsg, 100)
        errores = errores + 1
        If Not oMail Is Nothing Then Set oMail = Nothing
        On Error GoTo 0

SiguienteFila:
    Next i

    ' Guardar el libro para persistir los estados
    ThisWorkbook.Save

    ' Resumen final
    MsgBox "Proceso completado." & vbCrLf & vbCrLf & _
           "✅ Enviados: " & enviados & vbCrLf & _
           "❌ Errores:  " & errores & vbCrLf & vbCrLf & _
           "Revisa la columna ESTADO en la hoja COLA.", _
           vbInformation, "EnviarCorreos — Resumen"

    Set oApp = Nothing
    Exit Sub

ErrorOutlook:
    MsgBox "No se pudo conectar con Outlook." & vbCrLf & _
           "Asegúrate de que Outlook esté abierto con una sesión activa." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "Error Outlook"
    Set oApp = Nothing

End Sub


'' ─────────────────────────────────────────────────────────────────────────────
'' Sub auxiliar: limpiar la cola (pone todas las filas en PENDIENTE de nuevo)
'' Útil para reenviar en caso de errores masivos.
'' ─────────────────────────────────────────────────────────────────────────────
Sub ReiniciarCola()
    Dim ws         As Worksheet
    Dim ultimaFila As Long
    Dim i          As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("COLA")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No se encontró la hoja COLA.", vbCritical
        Exit Sub
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then Exit Sub

    Dim respuesta As Integer
    respuesta = MsgBox("¿Deseas marcar todas las filas como PENDIENTE?" & vbCrLf & _
                       "Esto permitirá re-enviar todos los correos.", _
                       vbYesNo + vbQuestion, "Reiniciar Cola")

    If respuesta = vbYes Then
        For i = 2 To ultimaFila
            ws.Cells(i, 5).Value = "PENDIENTE"
        Next i
        ThisWorkbook.Save
        MsgBox "Cola reiniciada. " & (ultimaFila - 1) & " filas marcadas como PENDIENTE.", _
               vbInformation
    End If
End Sub
