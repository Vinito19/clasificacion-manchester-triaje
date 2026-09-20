Option Explicit

'=============================
' CONFIGURACION
'=============================
Public Const TEXTO_ENCABEZADO As String = "Nombres"
Public Const HOJA_PACIENTES As String = "Pacientes"
Public Const COLUMNA_MANCHESTER As String = "Clasificaci�n Manchester"

'=============================
' CORRECCION #1: UltimaFilaDatosWs optimizada
' Antes: Iteraba ws.CheckBoxes Y ws.Shapes (doble iteracion)
' Ahora: Usa Range.Find (una sola operacion O(1))
'=============================
Public Function UltimaFilaDatosWs(ByVal ws As Worksheet) As Long
    Dim ultima As Long
    Dim lastCell As Range
    
    If ws Is Nothing Then Exit Function
    ultima = PrimeraFilaDatosWs(ws)
    
    ' CORRECCION: Usar Find para encontrar la ultima celda usada
    ' Esto es O(1) en lugar de O(n) iterando todos los shapes
    On Error Resume Next
    Set lastCell = ws.Cells.Find("*", ws.Cells(1, 1), _
        xlValues, , xlByRows, xlPrevious)
    If Not lastCell Is Nothing Then
        If lastCell.Row > ultima Then ultima = lastCell.Row
    End If
    On Error GoTo 0
    
    UltimaFilaDatosWs = ultima
End Function

'=========================================================
' LIMPIAR FILTRO DE FORMA SEGURA (funciona en hoja protegida)
'=========================================================
Public Sub LimpiarFiltroSeguro(ByVal ws As Worksheet)
    Dim estabaProtegida As Boolean
    If ws Is Nothing Then Exit Sub
    estabaProtegida = False
    On Error Resume Next
    estabaProtegida = ws.ProtectContents
    On Error GoTo 0
    On Error Resume Next
    If estabaProtegida Then ws.Unprotect
    If ws.AutoFilterMode Then
        ws.ShowAllData
        ws.AutoFilterMode = False
    End If
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

'=============================
' HOJA DE PACIENTES DE UN LIBRO
'=============================

'=========================================================
' PROTEGER COLUMNAS T Y U (fórmulas) DESDE FILA 11
'=========================================================
Public Sub ProtegerColumnasTU()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim prevScreen As Boolean, prevEvents As Boolean, prevCalc As XlCalculation

    Set ws = ActiveSheet
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo Limpiar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ' Desbloquear todo
    ws.Cells.Locked = False

    ' Bloquear columnas T(20) y U(21) desde fila 11 hasta la última con datos
    ultimaFila = UltimaFilaDatosWs(ws)
    If ultimaFila >= PrimeraFilaDatosWs(ws) Then
        ws.Range(ws.Cells(PrimeraFilaDatosWs(ws), 20), ws.Cells(ultimaFila, 21)).Locked = True
    End If

    ' Proteger hoja (sin contraseña)
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

Limpiar:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub

Public Function HojaPacientes(ByRef wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(HOJA_PACIENTES)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = wb.ActiveSheet
    Set HojaPacientes = ws
End Function

'=============================
' FUNCIONES POR HOJA (no dependen de ActiveSheet)
'=============================
Public Function FilaEncabezadoWs(ByVal ws As Worksheet) As Long
    Dim celda As Range
    If ws Is Nothing Then Exit Function
    Set celda = ws.Cells.Find(What:=TEXTO_ENCABEZADO, _
                              After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
                              LookAt:=xlWhole, LookIn:=xlValues, _
                              SearchOrder:=xlByRows, MatchCase:=False)
    If Not celda Is Nothing Then FilaEncabezadoWs = celda.Row
End Function

Public Function PrimeraFilaDatosWs(ByVal ws As Worksheet) As Long
    Dim filaTitulo As Long
    filaTitulo = FilaEncabezadoWs(ws)
    If filaTitulo > 0 Then
        PrimeraFilaDatosWs = filaTitulo + 2
    Else
        PrimeraFilaDatosWs = 9
    End If
End Function

Public Function UltimaColumnaFormatoWs(ByVal ws As Worksheet) As Long
    Dim filaTitulo As Long
    filaTitulo = FilaEncabezadoWs(ws)
    If filaTitulo = 0 Then
        UltimaColumnaFormatoWs = 15
        Exit Function
    End If
    UltimaColumnaFormatoWs = ws.Cells(filaTitulo, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Function ColumnaPrimerCheckboxWs(ByVal ws As Worksheet) As Long
    Dim c As Range
    Dim filaTitulo As Long
    If ws Is Nothing Then Exit Function
    filaTitulo = FilaEncabezadoWs(ws)
    If filaTitulo = 0 Then Exit Function
    Set c = ws.Rows(filaTitulo).Find(What:=COLUMNA_MANCHESTER, _
                                     After:=ws.Cells(filaTitulo, ws.Columns.Count), _
                                     LookAt:=xlWhole, LookIn:=xlValues, _
                                     SearchOrder:=xlByColumns, MatchCase:=False)
    If Not c Is Nothing Then ColumnaPrimerCheckboxWs = c.Column
End Function

Public Function ColumnaUltimoCheckboxWs(ByVal ws As Worksheet) As Long
    Dim pc As Long
    pc = ColumnaPrimerCheckboxWs(ws)
    If pc > 0 Then ColumnaUltimoCheckboxWs = pc + 4
End Function

Public Function ColumnaInicioColorWs(ByVal ws As Worksheet) As Long
    Dim uc As Long
    uc = ColumnaUltimoCheckboxWs(ws)
    If uc > 0 Then ColumnaInicioColorWs = uc + 1
End Function

'=============================
' COMPATIBILIDAD (libro/hoja activos)
'=============================
Public Function FilaEncabezado() As Long
    FilaEncabezado = FilaEncabezadoWs(ActiveSheet)
End Function

Public Function PRIMERA_FILA_DATOS() As Long
    PRIMERA_FILA_DATOS = PrimeraFilaDatosWs(ActiveSheet)
End Function

Public Function UltimaColumnaFormato() As Long
    UltimaColumnaFormato = UltimaColumnaFormatoWs(ActiveSheet)
End Function

Public Function ColumnaPrimerCheckbox() As Long
    Dim r As Long
    r = ColumnaPrimerCheckboxWs(ActiveSheet)
    If r = 0 Then
    End If
    ColumnaPrimerCheckbox = r
End Function

Public Function ColumnaUltimoCheckbox() As Long
    ColumnaUltimoCheckbox = ColumnaUltimoCheckboxWs(ActiveSheet)
End Function

Public Function ColumnaInicioColor() As Long
    ColumnaInicioColor = ColumnaInicioColorWs(ActiveSheet)
End Function

Public Function UltimaFilaDatos() As Long
    UltimaFilaDatos = UltimaFilaDatosWs(ActiveSheet)
End Function

'=============================
' COLORES POR HOJA
'=============================
Public Sub LimpiarColorFilaWs(ByVal ws As Worksheet, ByVal fila As Long)
    Dim ultimaCol As Long, colInicio As Long
    If ws Is Nothing Then Exit Sub
    ultimaCol = UltimaColumnaFormatoWs(ws)
    colInicio = ColumnaInicioColorWs(ws)
    If colInicio = 0 Or ultimaCol = 0 Then Exit Sub
    ws.Range(ws.Cells(fila, colInicio), ws.Cells(fila, ultimaCol)).Interior.Pattern = xlNone
End Sub

Public Sub AplicarColorWs(ByVal ws As Worksheet, ByVal fila As Long, ByVal columna As Long)
    Dim colorElegido As Long
    Dim ultimaCol As Long, primerCheck As Long, colInicio As Long
    If ws Is Nothing Then Exit Sub
    ultimaCol = UltimaColumnaFormatoWs(ws)
    primerCheck = ColumnaPrimerCheckboxWs(ws)
    colInicio = ColumnaInicioColorWs(ws)
    If primerCheck = 0 Or ultimaCol = 0 Or colInicio = 0 Then Exit Sub
    Select Case columna
        Case primerCheck:     colorElegido = RGB(230, 120, 130)
        Case primerCheck + 1: colorElegido = RGB(240, 170, 100)
        Case primerCheck + 2: colorElegido = RGB(230, 210, 80)
        Case primerCheck + 3: colorElegido = RGB(100, 190, 100)
        Case primerCheck + 4: colorElegido = RGB(100, 150, 220)
        Case Else: Exit Sub
    End Select
    ws.Range(ws.Cells(fila, colInicio), ws.Cells(fila, ultimaCol)).Interior.Color = colorElegido
End Sub

Public Sub LimpiarColorFila(ByVal fila As Long)
    LimpiarColorFilaWs ActiveSheet, fila
End Sub

Public Sub AplicarColor(ByVal fila As Long, ByVal columna As Long)
    AplicarColorWs ActiveSheet, fila, columna
End Sub

'=============================
' CORRECCION #2: ExisteForma con acceso directo
' Antes: Iteraba todos los shapes (O(n))
' Ahora: Acceso directo por nombre (O(1))
'=============================
Public Function ExisteForma(ByVal ws As Worksheet, ByVal nombre As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(nombre)
    On Error GoTo 0
    ExisteForma = Not shp Is Nothing
End Function

Public Function NombreCheckboxLibre(ByVal ws As Worksheet, ByVal fila As Long, ByVal col As Long) As String
    Dim base As String, nombre As String
    Dim n As Long
    base = "chk_" & fila & "_" & col
    nombre = base
    Do While ExisteForma(ws, nombre)
        n = n + 1
        nombre = base & "_" & n
    Loop
    NombreCheckboxLibre = nombre
End Function

Public Sub RelinkCheckboxesWs(ByVal ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlCheckBox Then
                ws.CheckBoxes(shp.Name).LinkedCell = shp.TopLeftCell.Address
                ws.CheckBoxes(shp.Name).OnAction = "CheckboxManchester"
            End If
        End If
    Next shp
    On Error GoTo 0
End Sub

'=============================
' CORRECCION #3: CheckboxManchester con ErrorHandler
' Antes: Saltaba a Salida sin registrar el error
' Ahora: Registra Err.Number y Err.Description en Debug.Print
'=============================
Public Sub CheckboxManchester()
    Dim cb As CheckBox
    Dim fila As Long, columna As Long, col As Long
    Dim primerCheck As Long, ultimoCheck As Long
    Dim ws As Worksheet
    Dim prevScreen As Boolean, prevEvents As Boolean

    If TypeName(Application.Caller) <> "String" Then Exit Sub

    Set ws = ActiveSheet
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set cb = ws.CheckBoxes(CStr(Application.Caller))
    fila = cb.TopLeftCell.Row
    columna = cb.TopLeftCell.Column

    primerCheck = ColumnaPrimerCheckboxWs(ws)
    ultimoCheck = ColumnaUltimoCheckboxWs(ws)

    If primerCheck = 0 Then GoTo Salida
    If fila < PrimeraFilaDatosWs(ws) Then GoTo Salida

    If cb.Value = xlOn Then
        Application.EnableEvents = False
        For col = primerCheck To ultimoCheck
            If col <> columna Then
                If ws.Cells(fila, col).Value = True Then ws.Cells(fila, col).Value = False
            End If
        Next col
        Application.EnableEvents = prevEvents
        LimpiarColorFilaWs ws, fila
        AplicarColorWs ws, fila, columna
    Else
        LimpiarColorFilaWs ws, fila
    End If

Salida:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrorHandler:
    Debug.Print "Error en CheckboxManchester (" & Application.Caller & "): " & _
                Err.Number & " - " & Err.Description
    Resume Salida
End Sub

'=============================
' NUEVO PACIENTE
'=============================
Public Sub NuevoPaciente()
    NuevoPacienteEn ThisWorkbook
End Sub

Public Sub NuevoPacienteEn(ByRef wb As Workbook)
    Dim ws As Worksheet
    Dim ultimaFila As Long, nuevaFila As Long
    Dim ultimaCol As Long
    Dim primerCheck As Long, ultimoCheck As Long
    Dim primeraColColor As Long
    Dim limitePacientes As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    Set ws = HojaPacientes(wb)

    LimpiarFiltroSeguro ws

    primerCheck = ColumnaPrimerCheckboxWs(ws)
    ultimoCheck = ColumnaUltimoCheckboxWs(ws)
    primeraColColor = ColumnaInicioColorWs(ws)
    ultimaCol = UltimaColumnaFormatoWs(ws)

    If primerCheck = 0 Then Exit Sub

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo Limpiar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ultimaFila = UltimaFilaDatosWs(ws)
    If ultimaFila < PrimeraFilaDatosWs(ws) Then ultimaFila = PrimeraFilaDatosWs(ws)

    limitePacientes = ObtenerMaxPacientes()

    If limitePacientes > 0 And (ultimaFila - PrimeraFilaDatosWs(ws) + 1) >= limitePacientes Then
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreen
        Application.Calculation = prevCalc
        MsgBox "Alcanz" & ChrW(243) & " el l" & ChrW(237) & "mite de " & limitePacientes & " pacientes. Se crear" & ChrW(225) & " un nuevo volumen autom" & ChrW(225) & "ticamente.", vbInformation
        CrearNuevoVolumenAutomatico
        Exit Sub
    End If

    nuevaFila = ultimaFila + 1

    ' CORRECCION #4: Usar procedimiento comun para copiar fila
    CopiarFilaPaciente ws, ultimaFila, nuevaFila

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.GoTo ws.Cells(nuevaFila, 1), True
    MsgBox "Nuevo registro creado en la fila " & nuevaFila & ". Complete los datos.", vbInformation
    Exit Sub

Limpiar:
    Application.CutCopyMode = False
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Ocurri" & ChrW(243) & " un error al crear el registro: " & Err.Description, vbCritical
End Sub

'=============================
' CORRECCION #4: Procedimiento comun para copiar fila de paciente
' Antes: NuevoPacienteEn y AgregarPacienteEnLibro tenian ~80% del codigo duplicado
' Ahora: Ambos llaman a este procedimiento comun
'=============================
Private Sub CopiarFilaPaciente(ByVal ws As Worksheet, _
                                ByVal origenFila As Long, _
                                ByVal destinoFila As Long)
    Dim primerCheck As Long, ultimoCheck As Long
    Dim primeraColColor As Long, ultimaCol As Long
    Dim i As Long, celda As Range
    Dim cbNuevo As CheckBox
    Dim ancho As Double, alto As Double
    
    primerCheck = ColumnaPrimerCheckboxWs(ws)
    ultimoCheck = ColumnaUltimoCheckboxWs(ws)
    primeraColColor = ColumnaInicioColorWs(ws)
    ultimaCol = UltimaColumnaFormatoWs(ws)
    
    ' Copiar formatos y validacion
    ws.Rows(origenFila).Copy
    ws.Rows(destinoFila).PasteSpecial xlPasteFormats
    ws.Rows(destinoFila).PasteSpecial xlPasteValidation
    Application.CutCopyMode = False
    
    ' Copiar formulas
    On Error Resume Next
    ws.Range(ws.Cells(origenFila, ultimoCheck + 1), _
             ws.Cells(origenFila, ultimaCol)).Copy
    ws.Range(ws.Cells(destinoFila, ultimoCheck + 1), _
             ws.Cells(destinoFila, ultimaCol)).PasteSpecial xlPasteFormulas
    On Error GoTo 0
    Application.CutCopyMode = False
    
    ' Limpiar constantes
    On Error Resume Next
    ws.Range(ws.Cells(destinoFila, 1), _
             ws.Cells(destinoFila, primerCheck - 1)).SpecialCells(xlCellTypeConstants).ClearContents
    ws.Range(ws.Cells(destinoFila, ultimoCheck + 1), _
             ws.Cells(destinoFila, ultimaCol)).SpecialCells(xlCellTypeConstants).ClearContents
    On Error GoTo 0
    
    ' Limpiar color
    If ultimaCol > 0 And primeraColColor > 0 Then
        ws.Range(ws.Cells(destinoFila, primeraColColor), _
                 ws.Cells(destinoFila, ultimaCol)).Interior.Pattern = xlNone
    End If
    
    ' Crear checkboxes
    For i = primerCheck To ultimoCheck
        Set celda = ws.Cells(destinoFila, i)
        ancho = celda.Width - 4: If ancho < 6 Then ancho = 6
        alto = celda.Height - 4: If alto < 6 Then alto = 6
        Set cbNuevo = ws.CheckBoxes.Add(celda.Left + 2, celda.Top + 2, ancho, alto)
        With cbNuevo
            .Caption = ""
            .Name = NombreCheckboxLibre(ws, destinoFila, i)
            .LinkedCell = celda.Address
            .OnAction = "CheckboxManchester"
            .Placement = xlMoveAndSize
            .Value = xlOff
        End With
    Next i
End Sub

'=============================
' AGREGAR PACIENTE EN UN LIBRO ESPECIFICO
'=============================
Public Sub AgregarPacienteEnLibro(ByRef wb As Workbook)
    Dim ws As Worksheet
    Dim ultimaFila As Long, nuevaFila As Long
    Dim chk As CheckBox
    Dim primerCheck As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    Set ws = HojaPacientes(wb)

    primerCheck = ColumnaPrimerCheckboxWs(ws)
    If primerCheck = 0 Then Exit Sub

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo Limpiar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ultimaFila = UltimaFilaDatosWs(ws)
    If ultimaFila < PrimeraFilaDatosWs(ws) Then ultimaFila = PrimeraFilaDatosWs(ws)

    ' Reutilizar la primera fila (plantilla vacía)
    If ultimaFila = PrimeraFilaDatosWs(ws) And Trim(CStr(ws.Cells(ultimaFila, 1).Value)) = "" Then
        For Each chk In ws.CheckBoxes
            If chk.TopLeftCell.Row = ultimaFila Then chk.Value = xlOff
        Next chk
        LimpiarColorFilaWs ws, ultimaFila
        Application.Calculation = prevCalc
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreen
        Exit Sub
    End If

    nuevaFila = ultimaFila + 1

    ' CORRECCION #4: Usar procedimiento comun para copiar fila
    CopiarFilaPaciente ws, ultimaFila, nuevaFila

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

Limpiar:
    Application.CutCopyMode = False
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub

'=============================
' ELIMINAR ULTIMO PACIENTE
'=============================
Public Sub EliminarUltimoPaciente()
    Dim ultimaFila As Long
    Dim respuesta As VbMsgBoxResult
    Dim ws As Worksheet
    Dim j As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    Set ws = ActiveSheet
    LimpiarFiltroSeguro ws

    ultimaFila = UltimaFilaDatosWs(ws)

    If ultimaFila < PrimeraFilaDatosWs(ws) Then
        MsgBox "No hay registros nuevos para eliminar.", vbInformation
        Exit Sub
    End If
    If ultimaFila = PrimeraFilaDatosWs(ws) And Trim(CStr(ws.Cells(ultimaFila, 1).Value)) = "" Then
        MsgBox "No hay registros nuevos para eliminar.", vbInformation
        Exit Sub
    End If

    respuesta = MsgBox(ChrW(191) & "Desea eliminar el " & ChrW(218) & "ltimo registro creado?", vbYesNo + vbQuestion, "Eliminar registro")
    If respuesta = vbNo Then Exit Sub

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    On Error GoTo Limpiar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For j = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(j).Type = msoFormControl Then
            If ws.Shapes(j).FormControlType = xlCheckBox Then
                If ws.Shapes(j).TopLeftCell.Row = ultimaFila Then ws.Shapes(j).Delete
            End If
        End If
    Next j

    ws.Rows(ultimaFila).Delete
    RelinkCheckboxesWs ws

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox ChrW(218) & "ltimo registro eliminado correctamente.", vbInformation
    Exit Sub

Limpiar:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Ocurri" & ChrW(243) & " un error al eliminar el registro: " & Err.Description, vbCritical
End Sub

'=============================
' CORRECCION #5: EliminarPacienteSeleccionado con Union
' Antes: Bubble sort O(n²) + eliminacion fila por fila (n reorganizaciones)
' Ahora: Union para eliminar todo de una vez (1 reorganizacion, sin ordenar)
'=============================
Public Sub EliminarPacienteSeleccionado()
    Dim ws As Worksheet
    Dim respuesta As VbMsgBoxResult
    Dim i As Long, j As Long, n As Long
    Dim filaActual As Long
    Dim filasAEliminar() As Long
    Dim primera As Long
    Dim rngEliminar As Range
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    Set ws = ActiveSheet
    LimpiarFiltroSeguro ws

    If Not TypeOf Selection Is Range Then
        MsgBox "Seleccione una o varias filas de pacientes.", vbInformation
        Exit Sub
    End If

    primera = PrimeraFilaDatosWs(ws)

    ReDim filasAEliminar(1 To Selection.Rows.Count)
    n = 0
    For i = 1 To Selection.Rows.Count
        filaActual = Selection.Rows(i).Row
        If filaActual < primera Then
            MsgBox "La selecci" & ChrW(243) & "n incluye la fila de encabezados. No se puede eliminar.", vbInformation
            Exit Sub
        End If
        n = n + 1
        filasAEliminar(n) = filaActual
    Next i

    If n = 0 Then Exit Sub
    ReDim Preserve filasAEliminar(1 To n)

    respuesta = MsgBox(ChrW(191) & "Desea eliminar los pacientes seleccionados?", vbYesNo + vbQuestion, "Eliminar pacientes")
    If respuesta = vbNo Then Exit Sub

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    On Error GoTo Limpiar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' CORRECCION: Usar Union para eliminar todo de una vez
    ' No necesitamos ordenar las filas porque Union maneja el orden automaticamente
    For i = 1 To n
        filaActual = filasAEliminar(i)
        
        ' Eliminar checkboxes de esta fila
        For j = ws.Shapes.Count To 1 Step -1
            If ws.Shapes(j).Type = msoFormControl Then
                If ws.Shapes(j).FormControlType = xlCheckBox Then
                    If ws.Shapes(j).TopLeftCell.Row = filaActual Then ws.Shapes(j).Delete
                End If
            End If
        Next j
        
        ' Acumular filas en el rango Union
        If rngEliminar Is Nothing Then
            Set rngEliminar = ws.Rows(filaActual)
        Else
            Set rngEliminar = Union(rngEliminar, ws.Rows(filaActual))
        End If
    Next i
    
    ' Eliminar todas las filas de una vez (una sola reorganizacion)
    If Not rngEliminar Is Nothing Then
        rngEliminar.Delete
    End If

    RelinkCheckboxesWs ws

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Pacientes seleccionados eliminados correctamente.", vbInformation
    Exit Sub

Limpiar:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Ocurri" & ChrW(243) & " un error al eliminar los pacientes: " & Err.Description, vbCritical
End Sub

'=========================================================
' ABRIR PANEL DE CONTROL
'=========================================================
Public Sub AbrirPanelControl()
    On Error GoTo ErrorForm
    frmPanelControl.Show vbModeless
    Exit Sub
ErrorForm:
    MsgBox "No se encontr" & ChrW(243) & " el Panel de Control." & vbCrLf & _
           "Verifique que el UserForm 'frmPanelControl' existe.", vbCritical
End Sub

'=========================================================
' COMBOBOX DE ACCIONES
'=========================================================
Public Sub ConfigurarMenuDesplegable()
    Dim ws As Worksheet
    Dim defaultText As String
    Dim f1 As String
    Set ws = HojaPacientes(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    defaultText = "Seleccionar acci" & ChrW(243) & "n..."

    ' Limpieza de la celda antigua (A5) por compatibilidad
    With ws.Range("A5")
        On Error Resume Next
        .Validation.Delete
        On Error GoTo 0
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
    ws.Rows(5).RowHeight = 15

    ' = CELDA A6:B6: BOTON ELEGANTE OSCURO =
    With ws.Range("A6:B6")
        .Merge
        .Value = defaultText
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Interior.Color = RGB(55, 65, 81)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(100, 116, 139)
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(100, 116, 139)
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(148, 163, 184)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(30, 41, 59)
        End With
    End With
    ws.Rows(6).RowHeight = 28.8
    ws.Columns(1).ColumnWidth = ws.Columns(2).ColumnWidth

    ' Sombra lateral (B6)
    ws.Range("B6").Interior.Color = RGB(30, 41, 59)

    ' Fila 7: separador simple
    ws.Rows(7).RowHeight = 15
    With ws.Range("A7:B7")
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With

    ' Restaurar borde superior de B8
    With ws.Range("B8")
        .Borders(xlEdgeTop).LineStyle = .Borders(xlEdgeLeft).LineStyle
        .Borders(xlEdgeTop).Weight = .Borders(xlEdgeLeft).Weight
        .Borders(xlEdgeTop).Color = .Borders(xlEdgeLeft).Color
    End With

    ' = VALIDACION DE DATOS =
    f1 = defaultText & _
         ",* Pacientes *" & _
         ",Nuevo Paciente" & _
         ",Eliminar " & ChrW(218) & "ltimo Registro" & _
         ",Eliminar Seleccionados" & _
         ",* Gesti" & ChrW(243) & "n *" & _
         ",Crear Nuevo Volumen" & _
         ",Generar Estad" & ChrW(237) & "sticas"
    With ws.Range("A6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=f1
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Men" & ChrW(250) & " de Acciones"
        .InputMessage = "Seleccione una acci" & ChrW(243) & "n del men" & ChrW(250)
        .ErrorTitle = "Opci" & ChrW(243) & "n no v" & ChrW(225) & "lida"
        .ErrorMessage = "Por favor, seleccione una opci" & ChrW(243) & "n v" & C
