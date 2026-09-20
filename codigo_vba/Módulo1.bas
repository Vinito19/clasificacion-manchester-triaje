Option Explicit

Private s_cacheMesEspanol As Variant

Private Function MesEspanol(ByVal fecha As Date) As String
    If IsEmpty(s_cacheMesEspanol) Then
        s_cacheMesEspanol = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    End If
    MesEspanol = s_cacheMesEspanol(Month(fecha) - 1)
End Function

Private Function ExtraerFechaDesdeNombre(ByVal nombreArchivo As String) As Date
    Dim pos As Long
    Dim fechaStr As String
    Dim partes() As String
    Dim d As Long, m As Long, y As Long
    
    pos = InStr(1, nombreArchivo, " - Historico ", vbTextCompare)
    If pos = 0 Then
        ExtraerFechaDesdeNombre = DateSerial(1900, 1, 1)
        Exit Function
    End If
    
    fechaStr = Mid(nombreArchivo, pos + Len(" - Historico "))
    If Len(fechaStr) < 10 Then
        ExtraerFechaDesdeNombre = DateSerial(1900, 1, 1)
        Exit Function
    End If
    
    fechaStr = Left(fechaStr, 10)
    partes = Split(fechaStr, "-")
    If UBound(partes) <> 2 Then
        ExtraerFechaDesdeNombre = DateSerial(1900, 1, 1)
        Exit Function
    End If
    
    If Not (IsNumeric(partes(0)) And IsNumeric(partes(1)) And IsNumeric(partes(2))) Then
        ExtraerFechaDesdeNombre = DateSerial(1900, 1, 1)
        Exit Function
    End If
    
    d = CLng(partes(0)): m = CLng(partes(1)): y = CLng(partes(2))
    On Error Resume Next
    ExtraerFechaDesdeNombre = DateSerial(y, m, d)
    If Err.Number <> 0 Then
        ExtraerFechaDesdeNombre = DateSerial(1900, 1, 1)
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Sub GenerarEstadisticas()
    ' ===== DECLARACIONES (TODAS AL INICIO, UNA SOLA VEZ) =====
    Dim carpetaHist As String
    Dim archivoHist As String
    Dim wbHist As Workbook
    Dim wsHist As Worksheet
    Dim cb As CheckBox
    Dim fila As Long
    Dim primerCheck As Long
    Dim colCheck As Long
    Dim primeraFila As Long
    Dim ultimaFila As Long
    Dim ultimaCol As Long
    Dim ultimoCheck As Long
    Dim fechaHist As Date
    Dim fechaClave As String
    Dim fechaStr As String
    Dim dictFechas As Object
    Dim dictColores As Object
    Dim key As Variant
    Dim prevScreen As Boolean
    Dim prevAlerts As Boolean
    Dim wbEst As Workbook
    Dim wsEst As Worksheet
    Dim nombreBase As String
    Dim fileName As String
    Dim totalRojo As Long
    Dim totalNaranja As Long
    Dim totalAmarillo As Long
    Dim totalVerde As Long
    Dim totalAzul As Long
    Dim fechasArray() As String
    Dim i As Long
    Dim j As Long
    Dim idxArr As Long
    Dim c As Long
    Dim datosArray As Variant
    Dim idx As Long
    Dim tmp As String
    Dim totalDia As Long
    Dim filaActual As Long
    Dim headers As Variant
    Dim carpetaEstadisticas As String
    
    ' ===== 1. VERIFICAR CARPETA HISTÓRICO =====
    carpetaHist = CarpetaHistorico()
    If Dir(carpetaHist, vbDirectory) = "" Then
        MsgBox "No existe la carpeta Historico.", vbExclamation
        Exit Sub
    End If
    
    ' ===== 2. CONFIGURAR APLICACIÓN =====
    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo Limpiar
    
    Debug.Print "=== GenerarEstadisticas INICIO ==="
    
    ' ===== 3. RECOPILAR ARCHIVOS (FileSystemObject) =====
    Dim fso As Object
    Dim carpeta As Object
    Dim archivo As Object
    Dim archivosHist() As String
    Dim fileCount As Long
    
    fileCount = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(carpetaHist)
    
    For Each archivo In carpeta.Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "xlsm" Then
            fileCount = fileCount + 1
            ReDim Preserve archivosHist(fileCount - 1)
            archivosHist(fileCount - 1) = archivo.Name
        End If
    Next archivo
    
    Debug.Print "Archivos encontrados: " & fileCount
    
    If fileCount = 0 Then
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        MsgBox "No hay archivos historicos.", vbExclamation
        Exit Sub
    End If
    
    Set dictFechas = CreateObject("Scripting.Dictionary")
    
    ' ===== 4. PROCESAR ARCHIVOS UNO A LA VEZ (OPTIMIZADO) =====
    For idx = 0 To fileCount - 1
        archivoHist = archivosHist(idx)
        
        Application.StatusBar = "Procesando " & (idx + 1) & " de " & fileCount & "..."
        Debug.Print "Procesando: " & archivoHist
        
        ' Abrir libro
        Set wbHist = Workbooks.Open(carpetaHist & "\" & archivoHist, UpdateLinks:=0, ReadOnly:=True)
        Set wsHist = wbHist.Worksheets(1)
        
        ' Extraer fecha
        fechaHist = ExtraerFechaDesdeNombre(archivoHist)
        fechaClave = Format(fechaHist, "DD-MM-YYYY")
        
        Debug.Print "  Fecha: " & fechaClave
        
        ' Crear entrada en dictionary si no existe
        If Not dictFechas.Exists(fechaClave) Then
            Set dictColores = CreateObject("Scripting.Dictionary")
            dictColores.Add "Rojo", 0
            dictColores.Add "Naranja", 0
            dictColores.Add "Amarillo", 0
            dictColores.Add "Verde", 0
            dictColores.Add "Azul", 0
            dictFechas.Add fechaClave, dictColores
        End If
        Set dictColores = dictFechas(fechaClave)
        
        ' Obtener información de la hoja
        primerCheck = ColumnaPrimerCheckboxWs(wsHist)
        ultimoCheck = ColumnaUltimoCheckboxWs(wsHist)
        primeraFila = PrimeraFilaDatosWs(wsHist)
        ultimaFila = UltimaFilaDatosWs(wsHist)
        ultimaCol = UltimaColumnaFormatoWs(wsHist)
        
        Debug.Print "  primerCheck=" & primerCheck & ", ultimoCheck=" & ultimoCheck
        Debug.Print "  primeraFila=" & primeraFila & ", ultimaFila=" & ultimaFila
        
        ' ===== VALIDACIÓN CRÍTICA =====
        If primerCheck > 0 And ultimoCheck > 0 And primeraFila > 0 And ultimaFila >= primeraFila Then
            ' Cargar datos en array (una sola lectura)
            datosArray = wsHist.Range(wsHist.Cells(primeraFila, primerCheck), _
                                      wsHist.Cells(ultimaFila, ultimoCheck)).Value
            
            ' Procesar en memoria
            For idxArr = LBound(datosArray, 1) To UBound(datosArray, 1)
                For c = LBound(datosArray, 2) To UBound(datosArray, 2)
                    If datosArray(idxArr, c) = True Then
                        Select Case c
                            Case 1: dictColores("Rojo") = dictColores("Rojo") + 1
                            Case 2: dictColores("Naranja") = dictColores("Naranja") + 1
                            Case 3: dictColores("Amarillo") = dictColores("Amarillo") + 1
                            Case 4: dictColores("Verde") = dictColores("Verde") + 1
                            Case 5: dictColores("Azul") = dictColores("Azul") + 1
                        End Select
                    End If
                Next c
            Next idxArr
            
            Debug.Print "  Datos procesados correctamente"
        Else
            Debug.Print "  ADVERTENCIA: Estructura de hoja inválida, saltando"
        End If
        
        ' Cerrar libro inmediatamente (liberar RAM)
        wbHist.Close SaveChanges:=False
        Set wbHist = Nothing
        Set wsHist = Nothing
    Next idx
    
    Application.StatusBar = False
    Debug.Print "Todos los archivos procesados"
    
    ' ===== 5. ORDENAR FECHAS =====
    ReDim fechasArray(dictFechas.Count - 1)
    i = 0
    For Each key In dictFechas.Keys
        fechasArray(i) = key
        i = i + 1
    Next key
    
    ' Ordenar (burbuja)
    For i = LBound(fechasArray) To UBound(fechasArray) - 1
        For j = i + 1 To UBound(fechasArray)
            If CDate(fechasArray(j)) < CDate(fechasArray(i)) Then
                tmp = fechasArray(i)
                fechasArray(i) = fechasArray(j)
                fechasArray(j) = tmp
            End If
        Next j
    Next i
    
    ' ===== 6. CALCULAR TOTALES =====
    totalRojo = 0: totalNaranja = 0: totalAmarillo = 0
    totalVerde = 0: totalAzul = 0
    
    For Each key In dictFechas.Keys
        Set dictColores = dictFechas(key)
        totalRojo = totalRojo + dictColores("Rojo")
        totalNaranja = totalNaranja + dictColores("Naranja")
        totalAmarillo = totalAmarillo + dictColores("Amarillo")
        totalVerde = totalVerde + dictColores("Verde")
        totalAzul = totalAzul + dictColores("Azul")
    Next key
    
    ' ===== 7. CREAR ARCHIVO DE ESTADÍSTICAS =====
    carpetaEstadisticas = CarpetaTrabajo() & "\Estadisticas Manchester"
    If Dir(carpetaEstadisticas, vbDirectory) = "" Then MkDir carpetaEstadisticas
    
    Set wbEst = Workbooks.Add
    Set wsEst = wbEst.Worksheets(1)
    wsEst.Name = "Estadisticas"
    
    wsEst.Range("A1").Value = "Estadisticas de Clasificacion Manchester"
    wsEst.Range("A2").Value = "Rango: " & fechasArray(LBound(fechasArray)) & " a " & fechasArray(UBound(fechasArray))
    wsEst.Range("A3").Value = "Generado el: " & Format(Date, "DD-MM-YYYY")
    wsEst.Range("A4").Value = "Total de archivos historicos: " & fileCount
    
    headers = Array("Fecha", "Rojo", "Naranja", "Amarillo", "Verde", "Azul", "Total Dia")
    For i = LBound(headers) To UBound(headers)
        wsEst.Cells(6, i + 1).Value = headers(i)
    Next i
    wsEst.Range("A6:G6").Font.Bold = True
    wsEst.Range("A6:G6").Interior.Color = RGB(200, 200, 200)
    wsEst.Range("A6:G6").Borders.LineStyle = xlContinuous
    
    ' ===== 8. ESCRIBIR DATOS POR FECHA =====
    filaActual = 7
    totalRojo = 0: totalNaranja = 0: totalAmarillo = 0
    totalVerde = 0: totalAzul = 0
    
    For idx = LBound(fechasArray) To UBound(fechasArray)
        fechaStr = fechasArray(idx)
        Set dictColores = dictFechas(fechasArray(idx))
        totalDia = dictColores("Rojo") + dictColores("Naranja") + dictColores("Amarillo") + _
                   dictColores("Verde") + dictColores("Azul")
        
        wsEst.Cells(filaActual, 1).Value = fechaStr
        wsEst.Cells(filaActual, 2).Value = dictColores("Rojo")
        wsEst.Cells(filaActual, 3).Value = dictColores("Naranja")
        wsEst.Cells(filaActual, 4).Value = dictColores("Amarillo")
        wsEst.Cells(filaActual, 5).Value = dictColores("Verde")
        wsEst.Cells(filaActual, 6).Value = dictColores("Azul")
        wsEst.Cells(filaActual, 7).Value = totalDia
        
        totalRojo = totalRojo + dictColores("Rojo")
        totalNaranja = totalNaranja + dictColores("Naranja")
        totalAmarillo = totalAmarillo + dictColores("Amarillo")
        totalVerde = totalVerde + dictColores("Verde")
        totalAzul = totalAzul + dictColores("Azul")
        
        wsEst.Range(wsEst.Cells(filaActual, 1), wsEst.Cells(filaActual, 7)).Borders.LineStyle = xlContinuous
        filaActual = filaActual + 1
    Next idx
    
    ' ===== 9. TOTAL GENERAL =====
    filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "TOTAL GENERAL"
    wsEst.Cells(filaActual, 1).Font.Bold = True
    wsEst.Cells(filaActual, 2).Value = totalRojo
    wsEst.Cells(filaActual, 3).Value = totalNaranja
    wsEst.Cells(filaActual, 4).Value = totalAmarillo
    wsEst.Cells(filaActual, 5).Value = totalVerde
    wsEst.Cells(filaActual, 6).Value = totalAzul
    wsEst.Cells(filaActual, 7).Value = totalRojo + totalNaranja + totalAmarillo + totalVerde + totalAzul
    wsEst.Range(wsEst.Cells(filaActual, 1), wsEst.Cells(filaActual, 7)).Font.Bold = True
    wsEst.Range(wsEst.Cells(filaActual, 1), wsEst.Cells(filaActual, 7)).Interior.Color = RGB(220, 220, 220)
    wsEst.Range(wsEst.Cells(filaActual, 1), wsEst.Cells(filaActual, 7)).Borders.LineStyle = xlContinuous
    
    ' ===== 10. RESUMEN POR COLORES =====
    filaActual = filaActual + 2
    wsEst.Cells(filaActual, 1).Value = "RESUMEN POR COLORES"
    wsEst.Cells(filaActual, 1).Font.Bold = True
    filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "Color"
    wsEst.Cells(filaActual, 2).Value = "Total"
    wsEst.Range(wsEst.Cells(filaActual, 1), wsEst.Cells(filaActual, 2)).Font.Bold = True
    filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "Rojo": wsEst.Cells(filaActual, 2).Value = totalRojo: filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "Naranja": wsEst.Cells(filaActual, 2).Value = totalNaranja: filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "Amarillo": wsEst.Cells(filaActual, 2).Value = totalAmarillo: filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "Verde": wsEst.Cells(filaActual, 2).Value = totalVerde: filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "Azul": wsEst.Cells(filaActual, 2).Value = totalAzul: filaActual = filaActual + 1
    wsEst.Cells(filaActual, 1).Value = "TOTAL"
    wsEst.Cells(filaActual, 2).Value = totalRojo + totalNaranja + totalAmarillo + totalVerde + totalAzul
    wsEst.Range(wsEst.Cells(filaActual, 1), wsEst.Cells(filaActual, 2)).Font.Bold = True
    wsEst.Range(wsEst.Cells(filaActual, 1), wsEst.Cells(filaActual, 2)).Interior.Color = RGB(220, 220, 220)
    
    wsEst.Columns("A:G").AutoFit
    
    ' ===== 11. GUARDAR Y CERRAR =====
    nombreBase = InicializarNombreBase()
    fileName = nombreBase & " - Estadisticas - " & Format(Date, "DD-MM-YYYY") & ".xlsx"
    wbEst.SaveAs carpetaEstadisticas & "\" & fileName, FileFormat:=xlOpenXMLWorkbook
    wbEst.Close SaveChanges:=False
    
    Debug.Print "=== GenerarEstadisticas EXITOSO ==="
    
    Application.StatusBar = False
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    MsgBox "Estadisticas generadas correctamente.", vbInformation
    Exit Sub
    
Limpiar:
    Debug.Print "=== ERROR en GenerarEstadisticas: " & Err.Number & " - " & Err.Description & " ==="
    Application.StatusBar = False
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    MsgBox "Error al generar estadisticas: " & Err.Description, vbCritical
End Sub


