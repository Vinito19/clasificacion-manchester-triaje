Option Explicit

'=========================================================
' MODULO: modVolumenes
' Administraci?n de vol?menes, hist?ricos y estadísticas
'=========================================================

Public Const CONFIG_HOJA As String = "CONFIG"

Public Const CFG_VOLUMEN As String = "Volumen"
Public Const CFG_MAXPACIENTES As String = "MaxPacientes"
Public Const CFG_VERSION As String = "Version"
Public Const CFG_HOSPITAL As String = "Hospital"
Public Const CFG_NOMBREBASE As String = "NombreBase"
Public gVigilanteOn As Boolean
Public gVigilanteHora As Date
Public gVigilanteCount As Long
Public gVigilanteCountCalc As Long
Public gRotacionEnCurso As Boolean
Public gRotacionStuck As Long

'=========================================================
' HOJA DE CONFIG
'=========================================================
Private Function HojaConfig(ByRef wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(CONFIG_HOJA)
    On Error GoTo 0
    Set HojaConfig = ws
End Function

'=========================================================
' LEER CONFIG
'=========================================================
Public Function LeerConfig(ByVal Clave As String) As String
    Dim ws As Worksheet
    Dim c As Range
    Set ws = HojaConfig(ThisWorkbook)
    If ws Is Nothing Then Exit Function
    Set c = ws.Columns(1).Find(What:=Clave, _
                               After:=ws.Cells(ws.Rows.Count, 1), _
                               LookAt:=xlWhole, LookIn:=xlValues)
    If Not c Is Nothing Then LeerConfig = CStr(c.Offset(0, 1).Value)
End Function

'=========================================================
' ESCRIBIR CONFIG (crea la clave si no existe)
'=========================================================
Public Sub EscribirConfig(ByVal Clave As String, ByVal valor As Variant)
    Dim ws As Worksheet
    Dim c As Range
    Dim ultimaFila As Long
    Set ws = HojaConfig(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    Set c = ws.Columns(1).Find(What:=Clave, _
                               After:=ws.Cells(ws.Rows.Count, 1), _
                               LookAt:=xlWhole, LookIn:=xlValues)
    If c Is Nothing Then
        ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(ultimaFila, 1).Value = Clave
        ws.Cells(ultimaFila, 2).Value = valor
    Else
        c.Offset(0, 1).Value = valor
    End If
End Sub

'=========================================================
' VOLUMEN
'=========================================================
Public Function ObtenerVolumen() As Long
    Dim s As String
    s = Trim(LeerConfig(CFG_VOLUMEN))
    If s = "" Then
        ObtenerVolumen = 1
    ElseIf IsNumeric(s) Then
        ObtenerVolumen = CLng(s)
    Else
        ObtenerVolumen = 1
    End If
End Function

Public Sub GuardarVolumen(ByVal Numero As Long)
    EscribirConfig CFG_VOLUMEN, Numero
End Sub

Public Sub IncrementarVolumen()
    GuardarVolumen ObtenerVolumen() + 1
End Sub

'=========================================================
' MAXIMO PACIENTES (valor por defecto 9999)
'=========================================================
Public Function ObtenerMaxPacientes() As Long
    Dim valorLeido As String
    valorLeido = Trim(LeerConfig(CFG_MAXPACIENTES))
    If valorLeido = "" Or Not IsNumeric(valorLeido) Then
        ObtenerMaxPacientes = 9999
    Else
        ObtenerMaxPacientes = CLng(valorLeido)
    End If
End Function

'=========================================================
' VERSION
'=========================================================
Public Function ObtenerVersion() As String
    ObtenerVersion = LeerConfig(CFG_VERSION)
End Function

'=========================================================
' HOSPITAL
'=========================================================
Public Function ObtenerHospital() As String
    ObtenerHospital = LeerConfig(CFG_HOSPITAL)
End Function

'=========================================================
' NOMBRE BASE
'=========================================================
Public Function ObtenerNombreBase() As String
    ObtenerNombreBase = LeerConfig(CFG_NOMBREBASE)
End Function

Public Sub GuardarNombreBase(ByVal nombre As String)
    EscribirConfig CFG_NOMBREBASE, nombre
End Sub

'=========================================================
' INICIALIZA EL NOMBRE BASE DEL SISTEMA
'=========================================================
Public Function InicializarNombreBase() As String
    Dim nombre As String
    Dim archivo As String
    Dim pos As Long
    Dim p As Long

    nombre = ObtenerNombreBase()
    If Trim(nombre) = "" Then
        p = InStrRev(ThisWorkbook.Name, ".")
        If p > 0 Then
            archivo = Left$(ThisWorkbook.Name, p - 1)
        Else
            archivo = ThisWorkbook.Name
        End If
        pos = InStrRev(LCase$(archivo), " vol ")
        If pos > 0 Then
            nombre = Left$(archivo, pos - 1)
        Else
            nombre = archivo
        End If
        nombre = Trim(nombre)
        GuardarNombreBase nombre
    End If
    InicializarNombreBase = nombre
End Function

'=========================================================
' RUTA COMPLETA DEL ARCHIVO ACTUAL
'=========================================================
Public Function RutaArchivoActual() As String
    RutaArchivoActual = ThisWorkbook.FullName
End Function

'=========================================================
' CARPETA DONDE EST? EL LIBRO
'=========================================================
Public Function CarpetaTrabajo() As String
    CarpetaTrabajo = ThisWorkbook.Path
End Function

'=========================================================
' EXTENSI?N DEL ARCHIVO (.xlsm)
'=========================================================
Public Function ExtensionArchivo() As String
    Dim p As Long
    p = InStrRev(ThisWorkbook.Name, ".")
    If p > 0 Then
        ExtensionArchivo = Mid$(ThisWorkbook.Name, p)
    Else
        ExtensionArchivo = ".xlsm"
    End If
End Function

'=========================================================
' CREA LA CARPETA HISTORICO SI NO EXISTE
'=========================================================
Public Function CarpetaHistorico() As String
    Dim ruta As String
    ruta = CarpetaTrabajo() & "\Historico"
    If ThisWorkbook.Path <> "" Then
        If Dir(ruta, vbDirectory) = "" Then
            On Error Resume Next
            MkDir ruta
            On Error GoTo 0
        End If
    End If
    CarpetaHistorico = ruta
End Function

'=========================================================
' NOMBRE DEL ARCHIVO HISTORICO (incluye fecha DD-MM-YYYY)
'=========================================================
Public Function NombreHistorico() As String
    NombreHistorico = InicializarNombreBase() & " - Vol " & ObtenerVolumen() & " - Historico " & Format(Date, "DD-MM-YYYY") & ExtensionArchivo()
End Function

'=========================================================
' RUTA COMPLETA DEL HISTORICO (?nica: no sobrescribe)
'=========================================================
Public Function RutaHistorico() As String
    RutaHistorico = CarpetaHistorico() & "\" & NombreHistorico()
End Function

Public Function RutaHistoricoUnica() As String
    Dim ruta As String, base As String, ext As String, n As Long
    ruta = RutaHistorico()
    If Dir(ruta) = "" Then
        RutaHistoricoUnica = ruta
        Exit Function
    End If
    base = NombreHistorico()
    ext = ExtensionArchivo()
    base = Left$(base, Len(base) - Len(ext))
    n = 1
    Do While Dir(CarpetaHistorico() & "\" & base & " (" & n & ")" & ext) <> ""
        n = n + 1
    Loop
    RutaHistoricoUnica = CarpetaHistorico() & "\" & base & " (" & n & ")" & ext
End Function

'=========================================================
' NOMBRE DEL NUEVO VOLUMEN
'=========================================================
Public Function NombreNuevoVolumen(ByVal numVolumen As Long) As String
    NombreNuevoVolumen = InicializarNombreBase() & " - Vol " & numVolumen & ExtensionArchivo()
End Function

'=========================================================
' RUTA DEL NUEVO VOLUMEN
'=========================================================
Public Function RutaNuevoVolumen(ByVal numVolumen As Long) As String
    RutaNuevoVolumen = CarpetaTrabajo() & "\" & NombreNuevoVolumen(numVolumen)
End Function

'=========================================================
' CREA UNA COPIA HISTORICA DEL VOLUMEN ACTUAL
'=========================================================
Public Sub CrearHistorico()
    Dim rutaDestino As String
    If ThisWorkbook.Path = "" Then
        MsgBox "El archivo debe estar guardado antes de crear un hist?rico.", vbCritical
        Exit Sub
    End If
    rutaDestino = RutaHistoricoUnica()
    On Error GoTo ErrCrear
    ThisWorkbook.SaveCopyAs rutaDestino
    MsgBox "Hist?rico creado: " & rutaDestino, vbInformation
    Exit Sub
ErrCrear:
    MsgBox "Error al crear el hist?rico: " & Err.Description, vbCritical
End Sub

'=========================================================
' ABRIR LIBRO DE FORMA SEGURA
'=========================================================
Private Function AbrirLibroSeguro(ByVal ruta As String, ByRef wb As Workbook) As Boolean
    On Error Resume Next
    Set wb = Workbooks.Open(ruta, UpdateLinks:=0)
    AbrirLibroSeguro = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'=========================================================
' LIMPIA EL NUEVO VOLUMEN
'=========================================================
Public Sub LimpiarNuevoVolumen()
    LimpiarNuevoVolumenEn ActiveWorkbook
End Sub

Public Sub LimpiarNuevoVolumenEn(ByRef wb As Workbook)
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long, j As Long
    Dim ultimaCol As Long
    Dim chk As CheckBox
    Dim primera As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation
    
    Set ws = HojaPacientes(wb)
    If ws Is Nothing Then Exit Sub
    
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Desproteger hoja
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    primera = PrimeraFilaDatosWs(ws)
    ultimaFila = UltimaFilaDatosWs(ws)
    
    ' 1) Borrar TODAS las filas de datos (desde ultimaFila hasta primera+1) - OPTIMIZADO
    If ultimaFila > primera Then
        ' PASO 1: Recopilar todos los checkboxes a eliminar (una sola iteracion)
        Dim shapesAEliminar As Collection
        Set shapesAEliminar = New Collection
        On Error Resume Next
        For j = ws.Shapes.Count To 1 Step -1
            If ws.Shapes(j).Type = msoFormControl Then
                If ws.Shapes(j).FormControlType = xlCheckBox Then
                    If ws.Shapes(j).TopLeftCell.Row >= (primera + 1) And _
                       ws.Shapes(j).TopLeftCell.Row <= ultimaFila Then
                        shapesAEliminar.Add ws.Shapes(j)
                    End If
                End If
            End If
        Next j
        On Error GoTo 0

        ' PASO 2: Eliminar todos los checkboxes recopilados
        For j = 1 To shapesAEliminar.Count
            shapesAEliminar(j).Delete
        Next j

        ' PASO 3: Borrar las filas (sin buscar shapes)
        For i = ultimaFila To (primera + 1) Step -1
            ws.Rows(i).Delete Shift:=xlUp
        Next i
    End If
    
    ' 2) Limpiar primera fila (plantilla): solo constantes, mantener fórmulas
    ultimaCol = UltimaColumnaFormatoWs(ws)
    If ultimaCol = 0 Then ultimaCol = 15
    
    On Error Resume Next
    ws.Range(ws.Cells(primera, 1), ws.Cells(primera, ultimaCol)).SpecialCells(xlCellTypeConstants).ClearContents
    On Error GoTo 0
    
    ' 3) Resetear checkboxes de la primera fila
    For Each chk In ws.CheckBoxes
        If chk.TopLeftCell.Row = primera Then chk.Value = xlOff
    Next chk
    
    ' 4) Limpiar color
    LimpiarColorFilaWs ws, primera
    
    ' 5) Reproteger hoja
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub
    
ErrorHandler:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Error al limpiar volumen: " & Err.Description, vbCritical
End Sub

'=========================================================
' ULTIMA FILA CON PACIENTES
'=========================================================
Public Function UltimaFilaDatosVolumen() As Long
    UltimaFilaDatosVolumen = UltimaFilaDatosWs(HojaPacientes(ActiveWorkbook))
End Function

'=========================================================
' ESCRIBIR CONFIG EN UN LIBRO ESPECIFICO
'=========================================================
Public Sub EscribirConfigEn(ByRef wb As Workbook, ByVal Clave As String, ByVal valor As Variant)
    Dim ws As Worksheet
    Dim c As Range
    Dim ultimaFila As Long
    Set ws = HojaConfig(wb)
    If ws Is Nothing Then Exit Sub
    Set c = ws.Columns(1).Find(What:=Clave, _
                               After:=ws.Cells(ws.Rows.Count, 1), _
                               LookAt:=xlWhole, LookIn:=xlValues)
    If c Is Nothing Then
        ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(ultimaFila, 1).Value = Clave
        ws.Cells(ultimaFila, 2).Value = valor
    Else
        c.Offset(0, 1).Value = valor
    End If
End Sub

'=========================================================
' NOMBRE DEL MES EN ESPA?OL
'=========================================================
Private Function MesEspanol(ByVal fecha As Date) As String
    Dim meses As Variant
    meses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", _
                  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    MesEspanol = meses(Month(fecha) - 1)
End Function

'=========================================================
' ELIMINAR ARCHIVO ORIGINAL DE FORMA SEGURA
'=========================================================
Private Sub EliminarArchivoSeguro(ByVal ruta As String)
    Dim i As Long
    Dim maxReintentos As Long
    Dim archivoLiberado As Boolean

    If Dir(ruta) = "" Then Exit Sub

    maxReintentos = 10

    For i = 1 To maxReintentos
        DoEvents
        On Error Resume Next
        If Dir(ruta) <> "" Then
            Kill ruta
        End If
        On Error GoTo 0

        If Dir(ruta) = "" Then
            archivoLiberado = True
            Exit For
        End If

        Application.Wait Now + TimeSerial(0, 0, 1) + (i * 0.1) / 86400
    Next i

    If Not archivoLiberado Then
        Debug.Print "No se pudo eliminar: " & ruta
    End If
End Sub

'=========================================================
' CREAR NUEVO VOLUMEN AUTOMATICO
'=========================================================
Private Function ArchivoEstaBloqueado(ByVal ruta As String) As Boolean
    Dim fileNum As Integer

    If Dir(ruta) = "" Then
        ArchivoEstaBloqueado = False
        Exit Function
    End If

    On Error Resume Next
    fileNum = FreeFile
    Open ruta For Binary Access Read Write As #fileNum
    If Err.Number <> 0 Then
        ArchivoEstaBloqueado = True
    Else
        ArchivoEstaBloqueado = False
        Close #fileNum
    End If
    On Error GoTo 0
End Function

Private Sub MoverAPendientes(ByVal ruta As String)
    Dim carpetaPendientes As String
    Dim nuevaRuta As String

    carpetaPendientes = CarpetaTrabajo() & "\_PendientesEliminar"

    If Dir(carpetaPendientes, vbDirectory) = "" Then
        On Error Resume Next
        MkDir carpetaPendientes
        On Error GoTo 0
    End If

    nuevaRuta = carpetaPendientes & "\" & Mid(ruta, InStrRev(ruta, "\") + 1)

    On Error Resume Next
    Name ruta As nuevaRuta
    On Error GoTo 0
End Sub

Public Sub CrearNuevoVolumenAutomatico()
    Dim rutaHist As String, rutaNuevo As String
    Dim wbNuevo As Workbook
    Dim viejoArchivo As String
    Dim nuevoVolumen As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    If ThisWorkbook.Path = "" Then
        MsgBox "El archivo debe estar guardado antes de crear un nuevo volumen.", vbCritical
        Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    gRotacionEnCurso = True
    On Error GoTo Limpiar

    ThisWorkbook.Save
    rutaHist = RutaHistoricoUnica()
    ThisWorkbook.SaveCopyAs rutaHist

    ' Generar estadísticas del mes actual
    Application.StatusBar = "Generando estadísticas del mes..."
    GenerarEstadisticas
    Application.StatusBar = False

    nuevoVolumen = ObtenerVolumen() + 1
    rutaNuevo = RutaNuevoVolumen(nuevoVolumen)

    If Dir(rutaNuevo) <> "" Then
        If Not AbrirLibroSeguro(rutaNuevo, wbNuevo) Then
            MsgBox "No se pudo abrir el archivo existente: " & rutaNuevo, vbCritical
            GoTo Limpiar
        End If
    Else
        ThisWorkbook.SaveCopyAs rutaNuevo
        If Not AbrirLibroSeguro(rutaNuevo, wbNuevo) Then
            MsgBox "No se pudo abrir el nuevo volumen.", vbCritical
            GoTo Limpiar
        End If
    End If

    EscribirConfigEn wbNuevo, CFG_VOLUMEN, nuevoVolumen
    wbNuevo.Save

    LimpiarNuevoVolumenEn wbNuevo
    wbNuevo.Save

    AgregarPacienteEnLibro wbNuevo
    wbNuevo.Save

    viejoArchivo = ThisWorkbook.FullName

    ' Resetear el menu del libro nuevo ANTES de cerrar
    On Error Resume Next
    wbNuevo.Worksheets(1).Range("A6").Value = "Seleccionar accion..."
    On Error GoTo 0

    ' Guardar el libro nuevo con el menu reseteado
    wbNuevo.Save

    ' RESTAURAR ESTADO ANTES DE CERRAR
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    gRotacionEnCurso = False
    Application.ScreenUpdating = prevScreen

    ' Cerrar el libro viejo
    ThisWorkbook.Close SaveChanges:=False

    ' Activar el nuevo libro
    wbNuevo.Activate

    Call ProgramarVigilante
    ' Dar tiempo a Windows para liberar el archivo
    Dim j As Long
    For j = 1 To 5
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next j

    ' Intentar eliminar el archivo viejo
    If ArchivoEstaBloqueado(viejoArchivo) Then
        EliminarArchivoSeguro viejoArchivo
        If Dir(viejoArchivo) <> "" Then
            MoverAPendientes viejoArchivo
            If Dir(viejoArchivo) <> "" Then
                MsgBox "El nuevo volumen " & nuevoVolumen & " fue creado." & vbCrLf & vbCrLf & _
                       "NOTA: No se pudo eliminar el archivo anterior." & vbCrLf & _
                       viejoArchivo & vbCrLf & "Se movio a _PendientesEliminar.", vbExclamation, "Atencion"
            End If
        End If
    Else
        EliminarArchivoSeguro viejoArchivo
    End If
        If Dir(viejoArchivo) <> "" Then MoverAPendientes viejoArchivo

    MsgBox "Nuevo volumen " & nuevoVolumen & " creado. El paciente ha sido registrado en el nuevo archivo.", vbInformation
    Exit Sub
    gRotacionEnCurso = False

Limpiar:
    On Error Resume Next
    If Not wbNuevo Is Nothing Then wbNuevo.Close SaveChanges:=False
    On Error GoTo 0
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    gRotacionEnCurso = False
    Application.ScreenUpdating = prevScreen
End Sub

'=========================================================
' CREAR NUEVO VOLUMEN MANUAL
'=========================================================
Public Sub CrearNuevoVolumenManual()
    Dim rutaHist As String, rutaNuevo As String
    Dim wbNuevo As Workbook
    Dim respuesta As VbMsgBoxResult
    Dim viejoArchivo As String
    Dim nuevoVolumen As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    If UltimaFilaDatosVolumen() <= PrimeraFilaDatosWs(HojaPacientes(ActiveWorkbook)) Then
        MsgBox "No hay pacientes registrados. No se puede crear un nuevo volumen.", vbExclamation
        Exit Sub
    End If

    respuesta = MsgBox("¿Está seguro de crear un nuevo volumen (sin paciente adicional)?", vbYesNo + vbQuestion)
    If respuesta = vbNo Then Exit Sub

    If ThisWorkbook.Path = "" Then
        MsgBox "El archivo debe estar guardado antes de crear un nuevo volumen.", vbCritical
        Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Limpiar
    gRotacionEnCurso = True

    ThisWorkbook.Save
    rutaHist = RutaHistoricoUnica()
    ThisWorkbook.SaveCopyAs rutaHist

    Application.StatusBar = "Generando estadísticas del mes..."
    GenerarEstadisticas
    Application.StatusBar = False

    nuevoVolumen = ObtenerVolumen() + 1
    rutaNuevo = RutaNuevoVolumen(nuevoVolumen)

    If Dir(rutaNuevo) <> "" Then
        If Not AbrirLibroSeguro(rutaNuevo, wbNuevo) Then
            MsgBox "No se pudo abrir el archivo existente.", vbCritical
            GoTo Limpiar
        End If
    Else
        ThisWorkbook.SaveCopyAs rutaNuevo
        If Not AbrirLibroSeguro(rutaNuevo, wbNuevo) Then
            MsgBox "No se pudo abrir el nuevo volumen.", vbCritical
            GoTo Limpiar
        End If
    End If

    EscribirConfigEn wbNuevo, CFG_VOLUMEN, nuevoVolumen
    wbNuevo.Save

    LimpiarNuevoVolumenEn wbNuevo
    wbNuevo.Save

    viejoArchivo = ThisWorkbook.FullName

    ' Resetear el menu del libro nuevo
    On Error Resume Next
    wbNuevo.Worksheets(1).Range("A6").Value = "Seleccionar accion..."
    On Error GoTo 0

    ' Guardar el libro nuevo
    wbNuevo.Save

    ' RESTAURAR ESTADO ANTES DE CERRAR
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    gRotacionEnCurso = False

    ' Cerrar el libro viejo
    ThisWorkbook.Close SaveChanges:=False

    ' Activar el nuevo libro
    wbNuevo.Activate

    ' Dar tiempo a Windows para liberar el archivo
    Call ProgramarVigilante
    Dim j As Long
    For j = 1 To 5
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next j

    ' Intentar eliminar el archivo viejo
    If ArchivoEstaBloqueado(viejoArchivo) Then
        EliminarArchivoSeguro viejoArchivo
        If Dir(viejoArchivo) <> "" Then
            MoverAPendientes viejoArchivo
            If Dir(viejoArchivo) <> "" Then
                MsgBox "El nuevo volumen " & nuevoVolumen & " fue creado." & vbCrLf & vbCrLf & _
                       "NOTA: No se pudo eliminar el archivo anterior." & vbCrLf & _
                       viejoArchivo & vbCrLf & "Se movio a _PendientesEliminar.", vbExclamation, "Atencion"
            End If
        End If
    Else
        EliminarArchivoSeguro viejoArchivo
    End If

        If Dir(viejoArchivo) <> "" Then MoverAPendientes viejoArchivo
    MsgBox "Nuevo volumen " & nuevoVolumen & " creado (sin pacientes).", vbInformation
    Exit Sub

    gRotacionEnCurso = False
Limpiar:
    On Error Resume Next
    If Not wbNuevo Is Nothing Then wbNuevo.Close SaveChanges:=False
    On Error GoTo 0
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    gRotacionEnCurso = False
End Sub

'=========================================================
' RESCATE DE ESTADO
'=========================================================
Sub RescateExcel()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Application.StatusBar = False
    MsgBox "?Excel ha vuelto a la normalidad!", vbInformation
End Sub


Public Sub ProgramarVigilante()
    On Error Resume Next
    gVigilanteOn = True
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:15"), Procedure:="VigilarEstadoApp"
    On Error GoTo 0
End Sub

Public Sub CancelarVigilante()
    On Error Resume Next
    If gVigilanteOn Then Application.OnTime EarliestTime:=gVigilanteHora, Procedure:="VigilarEstadoApp", Schedule:=False
    On Error GoTo 0
    gVigilanteOn = False
End Sub

Public Sub VigilarEstadoApp()
    Dim proxima As Date
    If Not gVigilanteOn Then Exit Sub
    If Not Application.EnableEvents Then
        gVigilanteCount = gVigilanteCount + 1
        If gVigilanteCount >= 2 Then
            Application.EnableEvents = True
            gVigilanteCount = 0
        End If
    Else
        gVigilanteCount = 0
    End If
    If Application.Calculation <> xlCalculationAutomatic Then
        gVigilanteCountCalc = gVigilanteCountCalc + 1
        If gVigilanteCountCalc >= 2 Then Application.Calculation = xlCalculationAutomatic
    Else
        gVigilanteCountCalc = 0
    End If
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    If gRotacionEnCurso Then
        gRotacionStuck = gRotacionStuck + 1
        If gRotacionStuck >= 8 Then
            gRotacionEnCurso = False
            gRotacionStuck = 0
            On Error Resume Next
            ThisWorkbook.Worksheets("Pacientes").Range("A6").Value = "Seleccionar accion..."
            On Error GoTo 0
        End If
    Else
        gRotacionStuck = 0
    End If
    proxima = Now + TimeValue("00:00:15")
    gVigilanteHora = proxima
    On Error Resume Next
    Application.OnTime proxima, "VigilarEstadoApp"
    On Error GoTo 0
End Sub

