Attribute VB_Name = "Modulo1"
''
'' EnviarCorreos.xlsm -- Modulo VBA: EnviarTodos (sin MsgBox).
''
'' Cambios vs version anterior:
''   - Sin MsgBox: el resumen y los errores se escriben a
''     <ThisWorkbook.Path>\logs\macro_envio.log (timestamped).
''   - Los errores por fila siguen quedando en la columna ESTADO de la hoja COLA.
''   - alertas_email.py lee este log + la hoja COLA tras la macro y los
''     emite al logger Python (sin dialogos modales que bloqueen el cierre).
''

Option Explicit

Private Const LOG_REL_PATH As String = "logs\macro_envio.log"

' Path absoluto local del log es escrito por Python en Hoja1!Y1 antes
' de invocar la macro (ThisWorkbook.Path devuelve URL https cuando el
' archivo esta en SharePoint/OneDrive y no sirve para Open #).
Private Function RutaLog() As String
    Dim r As String
    On Error Resume Next
    r = ThisWorkbook.Sheets("Hoja1").Range("Y1").Value
    On Error GoTo 0
    If r = "" Then
        r = ThisWorkbook.Path & "\logs\macro_envio.log"
    End If
    RutaLog = r
End Function

Private Sub EscribirLog(nivel As String, mensaje As String)
    Dim ruta As String
    Dim fnum As Integer
    ruta = RutaLog()

    On Error Resume Next
    fnum = FreeFile
    Open ruta For Append As #fnum
    If Err.Number = 0 Then
        Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & nivel & " | " & mensaje
        Close #fnum
    End If
    On Error GoTo 0
End Sub

Sub EnviarTodos()
    Dim ws          As Worksheet
    Dim oApp        As Object
    Dim oMail       As Object
    Dim ultimaFila  As Long
    Dim i           As Long
    Dim correo      As String
    Dim asunto      As String
    Dim cuerpo      As String
    Dim adjunto     As String
    Dim estado      As String
    Dim errMsg      As String
    Dim enviados    As Long
    Dim errores     As Long

    Call EscribirLog("INFO", "EnviarTodos: inicio")

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("COLA")
    On Error GoTo 0

    If ws Is Nothing Then
        Call EscribirLog("ERROR", "No se encontro la hoja COLA en este libro.")
        Exit Sub
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If ultimaFila < 2 Then
        Call EscribirLog("WARN", "La hoja COLA esta vacia (sin filas de datos).")
        Exit Sub
    End If

    On Error GoTo ErrorOutlook
    Set oApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    enviados = 0
    errores = 0

    For i = 2 To ultimaFila
        estado = Trim(ws.Cells(i, 5).Value)
        If UCase(estado) <> "PENDIENTE" Then GoTo SiguienteFila

        correo = Trim(ws.Cells(i, 1).Value)
        asunto = Trim(ws.Cells(i, 2).Value)
        cuerpo = ws.Cells(i, 3).Value
        adjunto = Trim(ws.Cells(i, 4).Value)

        If correo = "" Or correo = "nan" Then
            ws.Cells(i, 5).Value = "ERROR: correo vacio"
            errores = errores + 1
            Call EscribirLog("ERROR", "Fila " & i & " -- correo vacio")
            GoTo SiguienteFila
        End If

        On Error GoTo ErrorCorreo
        Set oMail = oApp.CreateItem(0)

        With oMail
            .To = correo
            .Subject = asunto
            .HTMLBody = cuerpo
            If adjunto <> "" And adjunto <> "nan" Then
                If Dir(adjunto) <> "" Then
                    .Attachments.Add adjunto
                Else
                    adjunto = "ADJUNTO_NO_ENCONTRADO"
                End If
            End If
            .Send
        End With

        If adjunto = "ADJUNTO_NO_ENCONTRADO" Then
            ws.Cells(i, 5).Value = "ENVIADO (sin adjunto)"
            Call EscribirLog("WARN", "Fila " & i & " -- enviado a " & correo & " sin adjunto (no se encontro: " & ws.Cells(i, 4).Value & ")")
        Else
            ws.Cells(i, 5).Value = "ENVIADO"
        End If
        enviados = enviados + 1

        Set oMail = Nothing
        Application.Wait Now + TimeValue("00:00:01")
        GoTo SiguienteFila

ErrorCorreo:
        errMsg = Err.Description
        ws.Cells(i, 5).Value = "ERROR: " & Left(errMsg, 100)
        errores = errores + 1
        Call EscribirLog("ERROR", "Fila " & i & " -- " & correo & " -- " & errMsg)
        If Not oMail Is Nothing Then Set oMail = Nothing
        On Error GoTo 0

SiguienteFila:
    Next i

    ThisWorkbook.Save

    Call EscribirLog("INFO", "Resumen -- enviados=" & enviados & " errores=" & errores & " total=" & (ultimaFila - 1))

    Set oApp = Nothing
    Exit Sub

ErrorOutlook:
    Call EscribirLog("ERROR", "No se pudo conectar con Outlook. " & Err.Description)
    Set oApp = Nothing
End Sub


''' Sub auxiliar: pone todas las filas en PENDIENTE (sin MsgBox).
Sub ReiniciarCola()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("COLA")
    On Error GoTo 0

    If ws Is Nothing Then
        Call EscribirLog("ERROR", "ReiniciarCola -- hoja COLA no encontrada.")
        Exit Sub
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then Exit Sub

    For i = 2 To ultimaFila
        ws.Cells(i, 5).Value = "PENDIENTE"
    Next i
    ThisWorkbook.Save
    Call EscribirLog("INFO", "ReiniciarCola -- " & (ultimaFila - 1) & " filas marcadas como PENDIENTE.")
End Sub
