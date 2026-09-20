Option Explicit
Private Const MAX_FILAS_BUSQUEDA As Long = 20
Private mRangoSeleccionado As Range



Private Sub Worksheet_Change(ByVal Target As Range)
    If gRotacionEnCurso Then Exit Sub
    Dim defaultText As String
    Dim accion As String
    Debug.Print "WS_Change: Target=" & Target.Address
    If Intersect(Target, Me.Range("A6:B6")) Is Nothing Then
        If Not Intersect(Target, Me.Range("B8")) Is Nothing Then
            Debug.Print "WS_Change: filtrar por apellido"
            AplicarFiltroApellidos
        End If
        Exit Sub
    End If
    defaultText = "Seleccionar acci" & ChrW(243) & "n..."
    If IsError(Me.Range("A6").Value) Then
        Debug.Print "WS_Change: valor convertido a formula/error, se resetea"
        Application.EnableEvents = False
        Me.Range("A6").Value = defaultText
        Application.EnableEvents = True
        Exit Sub
    End If
    accion = CStr(Me.Range("A6").Value)
    Debug.Print "WS_Change: accion leida = '" & accion & "'"
    If accion = defaultText Or accion = "" Then Exit Sub
    If Left(accion, 1) = "*" Then
        Debug.Print "WS_Change: separador, se resetea el menu"
        Application.EnableEvents = False
        Me.Range("A6").Value = defaultText
        Application.EnableEvents = True
        Exit Sub
    End If
    Application.EnableEvents = False
    On Error GoTo ErrorHandler
    Select Case accion
        Case "Nuevo Paciente"
            Debug.Print "WS_Change: ejecutar NuevoPaciente"
            NuevoPaciente
            Debug.Print "WS_Change: NuevoPaciente completado"
        Case "Eliminar " & ChrW(218) & "ltimo Registro"
            Debug.Print "WS_Change: ejecutar EliminarUltimoPaciente"
            EliminarUltimoPaciente
            Debug.Print "WS_Change: EliminarUltimoPaciente completado"
        Case "Eliminar Seleccionados"
            Debug.Print "WS_Change: ejecutar EliminarPacienteSeleccionado"
            If mRangoSeleccionado Is Nothing Then
                Debug.Print "WS_Change: sin seleccion previa, no se elimina"
                MsgBox "Seleccione primero la fila o filas a eliminar y despues elija esta opcion.", vbExclamation
            Else
                Debug.Print "WS_Change: usando seleccion guardada " & mRangoSeleccionado.Address
                On Error Resume Next
                mRangoSeleccionado.Select
                On Error GoTo 0
                EliminarPacienteSeleccionado
            End If
            Debug.Print "WS_Change: EliminarPacienteSeleccionado completado"
        Case "Crear Nuevo Volumen"
            Debug.Print "WS_Change: ejecutar CrearNuevoVolumenManual"
            CrearNuevoVolumenManual
            Debug.Print "WS_Change: CrearNuevoVolumenManual completado"
        Case "Generar Estad" & ChrW(237) & "sticas"
            Debug.Print "WS_Change: ejecutar GenerarEstadisticas"
            GenerarEstadisticas
            Debug.Print "WS_Change: GenerarEstadisticas completado"
        Case Else
            MsgBox "Accion no reconocida: '" & accion & "'", vbExclamation
    End Select
    Set mRangoSeleccionado = Nothing
    Me.Range("A6").Value = defaultText
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    Debug.Print "WS_Change: ERROR en '" & accion & "': " & Err.Number & " - " & Err.Description
    MsgBox "Error al ejecutar '" & accion & "':" & vbCrLf & Err.Description, vbCritical
    Set mRangoSeleccionado = Nothing
    Me.Range("A6").Value = defaultText
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A6:B6")) Is Nothing Then Exit Sub
    If Not Intersect(Target, Me.Range("B8")) Is Nothing Then Exit Sub
    On Error Resume Next
    Set mRangoSeleccionado = Selection
    On Error GoTo 0
End Sub

Private Sub AplicarFiltroApellidos()
    Dim hoja As Worksheet
    Dim filaEncab As Long, filaDatos As Long, ultimaFila As Long
    Dim ultimaCol As Long
    Dim colApellidos As Long
    Dim rngTabla As Range
    Dim textoBusqueda As String
    Dim shp As Shape
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim estabaProtegida As Boolean
    Dim celda As Range
    
    Set hoja = Me
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Desproteger temporalmente si hace falta
    estabaProtegida = False
    On Error Resume Next
    estabaProtegida = hoja.ProtectContents
    On Error GoTo 0
    
    If estabaProtegida Then
        On Error Resume Next
        hoja.Unprotect
        On Error GoTo 0
    End If
    
    textoBusqueda = Trim(hoja.Range("B8").Value)
    
    ' ===== BUSCAR CABECERA DE APELLIDO (For Each + Trim + UCase) =====
    colApellidos = 0
    filaEncab = 0
    
    For Each celda In hoja.Range(hoja.Cells(1, 1), hoja.Cells(MAX_FILAS_BUSQUEDA, hoja.Columns.Count))
        If UCase(Trim(celda.Value)) = "APELLIDO" Or UCase(Trim(celda.Value)) = "APELLIDOS" Then
            colApellidos = celda.Column
            filaEncab = celda.Row
            Exit For
        End If
    Next celda
    
    ' Sin cabecera de apellido: quitar filtro y salir
    If colApellidos = 0 Or filaEncab = 0 Then
        On Error Resume Next
        hoja.ShowAllData
        hoja.AutoFilterMode = False
        On Error GoTo 0
        GoTo Salida
    End If
    
    filaDatos = filaEncab + 2
    ultimaFila = UltimaFilaDatosWs(hoja)
    If ultimaFila < filaDatos Then ultimaFila = filaDatos
    
    ' Determinar ultima columna de la fila de cabecera
    ultimaCol = hoja.Cells(filaEncab, hoja.Columns.Count).End(xlToLeft).Column
    
    ' ===== CASO 1: TEXTO VACIO - QUITAR FILTRO Y MOSTRAR TODO =====
    If textoBusqueda = "" Then
        On Error Resume Next
        hoja.ShowAllData
        hoja.AutoFilterMode = False
        On Error GoTo 0
        
        For Each shp In hoja.Shapes
            If shp.Type = msoFormControl Then
                If shp.FormControlType = xlCheckBox Then
                    If shp.TopLeftCell.Row >= filaDatos Then shp.Visible = True
                End If
            End If
        Next shp
        GoTo Salida
    End If
    
    ' ===== CASO 2: BUSCAR POR APELLIDO =====
    Set rngTabla = hoja.Range(hoja.Cells(filaEncab, 1), hoja.Cells(ultimaFila, ultimaCol))
    If rngTabla.Rows.Count < 2 Then GoTo Salida
    
    On Error Resume Next
    rngTabla.AutoFilter Field:=colApellidos, Criteria1:="=*" & textoBusqueda & "*", Operator:=xlAnd
    On Error GoTo 0
    
    ' ===== OCULTAR/MOSTRAR CHECKBOXES SEGUN FILA VISIBLE =====
    On Error Resume Next
    For Each shp In hoja.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlCheckBox Then
                If shp.TopLeftCell.Row >= filaDatos Then
                    shp.Visible = Not hoja.Rows(shp.TopLeftCell.Row).EntireRow.Hidden
                End If
            End If
        End If
    Next shp
    On Error GoTo 0
    
Salida:
    On Error Resume Next
    If estabaProtegida Then hoja.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en AplicarFiltroApellidos: " & Err.Number & " - " & Err.Description
    Resume Salida
End Sub
