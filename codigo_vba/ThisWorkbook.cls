Private Sub Workbook_Open()
    On Error GoTo ErrorHandler

    ' GARANTIZAR QUE EL ESTADO SEA CORRECTO
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Configurar el menu desplegable
    Call ConfigurarMenuDesplegable

    ' RESETEAR EL MENU SI QUEDÓ CON VALOR INCORRECTO
    Dim ws As Worksheet
    Set ws = HojaPacientes(ThisWorkbook)
    If Not ws Is Nothing Then
        On Error Resume Next
        If ws.Range("A6").Value <> "Seleccionar acción..." Then
            ws.Range("A6").Value = "Seleccionar acción..."
        End If
        On Error GoTo 0
    End If

    ' Asegurar que la hoja de pacientes esta activa y proteger columnas T y U
    ThisWorkbook.Worksheets("Pacientes").Activate
    ProtegerColumnasTU

    Call ProgramarVigilante
    Exit Sub

ErrorHandler:
    Debug.Print "Error en Workbook_Open: " & Err.Number & " - " & Err.Description
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Call CancelarVigilante
    On Error GoTo 0
End Sub
