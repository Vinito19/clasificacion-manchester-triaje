Option Explicit

'=========================================================
' MACRO DE DIAGNOSTICO - Verificar estado del archivo y VBA
' (version corregida: nunca ejecuta las macros del libro)
' Ejecutar DiagnosticarArchivo para el reporte completo
'=========================================================

Public Sub DiagnosticarArchivo()
    Dim wsDatos As Worksheet
    Dim wsMenu As Worksheet
    Dim resultado As String
    Dim problemas As String
    Dim advertencias As String
    Dim comp As Object
    Dim nComp As Long

    resultado = "=== DIAGNOSTICO DEL ARCHIVO ===" & vbCrLf & vbCrLf

    On Error Resume Next
    nComp = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        MsgBox "No se puede acceder al proyecto VBA." & vbCrLf & vbCrLf & _
               "Activa: Archivo > Opciones > Centro de confianza > " & _
               "'Confiar en acceso al modelo de objetos de proyectos de VBA'.", _
               vbCritical, "Diagnostico"
        Exit Sub
    End If
    On Error GoTo 0

    ' 1. ubicacion del archivo
    resultado = resultado & "1. UBICACION DEL ARCHIVO:" & vbCrLf
    resultado = resultado & "   Ruta: " & ThisWorkbook.FullName & vbCrLf
    If InStr(LCase(ThisWorkbook.FullName), "onedrive") > 0 Or _
       InStr(LCase(ThisWorkbook.FullName), "dropbox") > 0 Or _
       InStr(LCase(ThisWorkbook.FullName), "google drive") > 0 Or _
       InStr(LCase(ThisWorkbook.FullName), "teams") > 0 Then
        problemas = problemas & "   [CRITICO] El archivo esta en una ruta sincronizada." & vbCrLf
        problemas = problemas & "   Pueden producirse conflictos de guardado." & vbCrLf
        problemas = problemas & "   SOLUCION: mueve el archivo a una carpeta local." & vbCrLf & vbCrLf
    Else
        resultado = resultado & "   [OK] El archivo esta en ubicacion local." & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' 2. modo de apertura
    resultado = resultado & "2. MODO DE APERTURA:" & vbCrLf
    If ThisWorkbook.ReadOnly Then
        problemas = problemas & "   [CRITICO] El archivo se abrio en modo SOLO LECTURA." & vbCrLf
        problemas = problemas & "   Los cambios NO se guardaran." & vbCrLf & vbCrLf
    Else
        resultado = resultado & "   [OK] Modo lectura/escritura." & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' 3. cambios sin guardar
    resultado = resultado & "3. ESTADO DE GUARDADO:" & vbCrLf
    If ThisWorkbook.Saved Then
        resultado = resultado & "   [OK] No hay cambios pendientes de guardar." & vbCrLf
    Else
        advertencias = advertencias & "   [ADVERTENCIA] Hay cambios SIN GUARDAR." & vbCrLf
        advertencias = advertencias & "   Presiona Ctrl+S para guardar." & vbCrLf & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' 4. modulos VBA
    resultado = resultado & "4. MODULOS VBA ENCONTRADOS:" & vbCrLf
    Dim countMod As Long, countClass As Long, countForm As Long, countDoc As Long
    countMod = 0: countClass = 0: countForm = 0: countDoc = 0
    On Error Resume Next
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                countMod = countMod + 1
                resultado = resultado & "   [MODULO] " & comp.Name & vbCrLf
            Case 2
                countClass = countClass + 1
                resultado = resultado & "   [CLASE] " & comp.Name & vbCrLf
            Case 3
                countForm = countForm + 1
                resultado = resultado & "   [FORM] " & comp.Name & vbCrLf
            Case 100
                countDoc = countDoc + 1
                resultado = resultado & "   [HOJA/LIBRO] " & comp.Name & vbCrLf
        End Select
    Next comp
    On Error GoTo 0
    resultado = resultado & vbCrLf
    resultado = resultado & "   Total: " & countMod & " modulos, " & countClass & _
                            " clases, " & countForm & " formularios, " & countDoc & _
                            " hojas/libro." & vbCrLf
    resultado = resultado & vbCrLf

    ' 5. funciones criticas (escaneo estatico, NUNCA se ejecutan)
    resultado = resultado & "5. VERIFICACION DE FUNCIONES CRITICAS:" & vbCrLf
    Dim funcionesCriticas As Variant
    Dim func As Variant
    Dim enModulo As String
    funcionesCriticas = Array("NuevoPaciente", "EliminarUltimoPaciente", _
                              "EliminarPacienteSeleccionado", "CrearNuevoVolumenManual", _
                              "GenerarEstadisticas", "ProtegerColumnasTU", _
                              "ConfigurarMenuDesplegable")
    For Each func In funcionesCriticas
        enModulo = ""
        If ExisteProcedimiento(CStr(func), enModulo) Then
            resultado = resultado & "   [OK] " & func & " existe (en " & enModulo & ")" & vbCrLf
        Else
            problemas = problemas & "   [CRITICO] " & func & " NO EXISTE." & vbCrLf
        End If
    Next func
    resultado = resultado & vbCrLf

    ' 6. hoja de datos
    resultado = resultado & "6. HOJA DE PACIENTES:" & vbCrLf
    Set wsDatos = Nothing
    On Error Resume Next
    Set wsDatos = HojaPacientes(ThisWorkbook)
    On Error GoTo 0
    If wsDatos Is Nothing Then
        On Error Resume Next
        Set wsDatos = ThisWorkbook.Worksheets("Pacientes")
        On Error GoTo 0
    End If
    If wsDatos Is Nothing Then
        Set wsDatos = ThisWorkbook.ActiveSheet
        advertencias = advertencias & "   [ADVERTENCIA] No se encontro la hoja de datos 'Pacientes'." & vbCrLf
        advertencias = advertencias & "   Usando hoja activa: " & wsDatos.Name & vbCrLf & vbCrLf
    Else
        resultado = resultado & "   [OK] Hoja de datos: " & wsDatos.Name & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' 7. menu desplegable (hoja con Worksheet_Change)
    resultado = resultado & "7. MENU DESPLEGABLE (A6):" & vbCrLf
    Set wsMenu = HojaMenu()
    If wsMenu Is Nothing Then
        problemas = problemas & "   [CRITICO] No se encontro la hoja con Worksheet_Change." & vbCrLf & vbCrLf
    Else
        resultado = resultado & "   Hoja del menu: " & wsMenu.Name & vbCrLf
        Dim valorMenu As String
        valorMenu = CStr(wsMenu.Range("A6").Value)
        If Len(Trim(valorMenu)) = 0 Then
            problemas = problemas & "   [CRITICO] La celda A6 de " & wsMenu.Name & " esta vacia." & vbCrLf
            problemas = problemas & "   Ejecuta ConfigurarMenuDesplegable." & vbCrLf & vbCrLf
        ElseIf valorMenu = "Seleccionar acci" & ChrW(243) & "n..." Then
            resultado = resultado & "   [OK] Menu configurado correctamente." & vbCrLf
        Else
            advertencias = advertencias & "   [INFO] A6 tiene un valor seleccionado: " & valorMenu & vbCrLf & vbCrLf
        End If

        Dim tieneValidacion As Boolean
        tieneValidacion = False
        On Error Resume Next
        If wsMenu.Range("A6").Validation.Type = 3 Then tieneValidacion = True
        On Error GoTo 0
        If tieneValidacion Then
            resultado = resultado & "   [OK] Validacion de datos (lista) configurada." & vbCrLf
        Else
            problemas = problemas & "   [CRITICO] A6 no tiene validacion de lista." & vbCrLf
            problemas = problemas & "   Ejecuta ConfigurarMenuDesplegable." & vbCrLf & vbCrLf
        End If
    End If
    resultado = resultado & vbCrLf

    ' 8. estado de la aplicacion
    resultado = resultado & "8. ESTADO DE LA APLICACION:" & vbCrLf
    Dim calcTexto As String
    Select Case Application.Calculation
        Case -4105: calcTexto = "Automatico"
        Case -4135: calcTexto = "Manual"
        Case -4136: calcTexto = "SemiAutomatico"
        Case Else: calcTexto = "Otro"
    End Select
    resultado = resultado & "   ScreenUpdating: " & IIf(Application.ScreenUpdating, "True", "False") & vbCrLf
    resultado = resultado & "   EnableEvents: " & IIf(Application.EnableEvents, "True", "False") & vbCrLf
    resultado = resultado & "   Calculation: " & calcTexto & vbCrLf
    If Not Application.EnableEvents Then
        problemas = problemas & "   [CRITICO] EnableEvents esta en FALSE." & vbCrLf
        problemas = problemas & "   Los eventos (Worksheet_Change) NO funcionaran." & vbCrLf
        problemas = problemas & "   SOLUCION: cierra y vuelve a abrir Excel." & vbCrLf & vbCrLf
    Else
        resultado = resultado & "   [OK] EnableEvents activo." & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' 9. configuracion del sistema
    resultado = resultado & "9. CONFIGURACION DEL SISTEMA:" & vbCrLf
    Dim volumen As String
    Dim maxPac As String
    volumen = LeerConfig("Volumen")
    maxPac = LeerConfig("MaxPacientes")
    If Trim(volumen) <> "" Then
        resultado = resultado & "   Volumen actual: " & volumen & vbCrLf
    Else
        advertencias = advertencias & "   [ADVERTENCIA] No hay configuracion de Volumen." & vbCrLf
    End If
    If Trim(maxPac) <> "" Then
        resultado = resultado & "   Max pacientes: " & maxPac & vbCrLf
    Else
        advertencias = advertencias & "   [ADVERTENCIA] No hay configuracion de MaxPacientes." & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' 10. recomendacion final
    resultado = resultado & "=== RECOMENDACION ===" & vbCrLf & vbCrLf
    If problemas = "" And advertencias = "" Then
        resultado = resultado & "[TODO CORRECTO]" & vbCrLf & vbCrLf
        resultado = resultado & "El archivo esta en buen estado." & vbCrLf
    Else
        If problemas <> "" Then
            resultado = resultado & "PROBLEMAS CRITICOS:" & vbCrLf
            resultado = resultado & problemas & vbCrLf
        End If
        If advertencias <> "" Then
            resultado = resultado & "DATOS A REVISAR:" & vbCrLf
            resultado = resultado & advertencias & vbCrLf
        End If
    End If

    MsgBox resultado, vbInformation, "Diagnostico del Archivo"

    Dim archivoDiagnostico As String
    archivoDiagnostico = ThisWorkbook.Path & Application.PathSeparator & _
                         "diagnostico_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    Dim ff As Integer
    ff = FreeFile
    On Error Resume Next
    Open archivoDiagnostico For Output As #ff
    Print #ff, resultado
    Close #ff
    If Err.Number = 0 Then
        MsgBox "Reporte guardado en:" & vbCrLf & archivoDiagnostico, vbInformation, "Diagnostico"
    Else
        MsgBox "No se pudo guardar el reporte: " & Err.Description, vbExclamation, "Diagnostico"
    End If
    On Error GoTo 0
End Sub

'=========================================================
' MACRO: Forzar guardado completo del proyecto VBA
'=========================================================
Public Sub ForzarGuardadoVBA()
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("Esta macro guardara el archivo completo." & vbCrLf & vbCrLf & _
                       "Se recomienda:" & vbCrLf & _
                       "1. Cerrar los demas libros de Excel" & vbCrLf & _
                       "2. Guardar este archivo" & vbCrLf & _
                       "3. Cerrar Excel completamente" & vbCrLf & _
                       "4. Volver a abrir el archivo" & vbCrLf & vbCrLf & _
                       "Continuar?", vbYesNo + vbQuestion, "Forzar Guardado")
    If respuesta = vbNo Then Exit Sub
    On Error GoTo ErrorGuardar
    ThisWorkbook.Save
    DoEvents
    MsgBox "Archivo guardado correctamente." & vbCrLf & vbCrLf & _
           "AHORA:" & vbCrLf & _
           "1. Cierra Excel (Ctrl+Q o Archivo > Salir)" & vbCrLf & _
           "2. Vuelve a abrir el archivo" & vbCrLf & _
           "3. Verifica que los cambios persistan", vbInformation, "Guardado Completo"
    Exit Sub
ErrorGuardar:
    MsgBox "Error al guardar: " & Err.Description & vbCrLf & vbCrLf & _
           "Posibles causas:" & vbCrLf & _
           "- El archivo esta en modo solo lectura" & vbCrLf & _
           "- No tienes permisos de escritura" & vbCrLf & _
           "- El archivo esta siendo usado por otro programa", vbCritical, "Error"
End Sub

'=========================================================
' MACRO: Verificar si un procedimiento existe (solo lectura)
'=========================================================
Public Sub VerificarFuncion(ByVal nombreFuncion As String)
    Dim enModulo As String
    enModulo = ""
    If ExisteProcedimiento(nombreFuncion, enModulo) Then
        Dim comp As Object
        Dim cm As Object
        For Each comp In ThisWorkbook.VBProject.VBComponents
            If comp.Name = enModulo Then
                Set cm = comp.CodeModule
                Dim primeras As String
                primeras = cm.Lines(1, cm.CountOfLines)
                primeras = Left(primeras, 500)
                MsgBox "PROCEDIMIENTO ENCONTRADO:" & vbCrLf & vbCrLf & _
                       "Nombre: " & nombreFuncion & vbCrLf & _
                       "Modulo: " & enModulo & vbCrLf & vbCrLf & _
                       "Codigo (primeras lineas):" & vbCrLf & _
                       primeras & " ...", vbInformation, "Verificacion"
                Exit Sub
            End If
        Next comp
    End If
    MsgBox "El procedimiento '" & nombreFuncion & "' NO se encontro en el proyecto.", _
           vbExclamation, "Verificacion"
End Sub

'=========================================================
' Auxiliar: detecta si un procedimiento existe mediante
' escaneo estatico del codigo. NUNCA ejecuta la macro.
'=========================================================
Private Function ExisteProcedimiento(ByVal nombre As String, ByRef enModulo As String) As Boolean
    Dim comp As Object
    Dim cm As Object
    Dim l As Long
    Dim lin As String
    Dim x As String
    Dim cuerpo As String
    ExisteProcedimiento = False
    On Error Resume Next
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Set cm = comp.CodeModule
        For l = 1 To cm.CountOfLines
            lin = Trim(cm.Lines(l, 1))
            If LCase(Left(lin, 7)) = "public " Then lin = Mid(lin, 8)
            If LCase(Left(lin, 8)) = "private " Then lin = Mid(lin, 9)
            If LCase(Left(lin, 7)) = "friend " Then lin = Mid(lin, 8)
            x = Left(lin, 4)
            If x = "Sub " Or x = "Func" Then
                If x = "Sub " Then
                    cuerpo = Trim(Mid(lin, 5))
                Else
                    cuerpo = Trim(Mid(lin, 10))
                End If
                If LCase(cuerpo) Like LCase(nombre) & "*" Then
                    ExisteProcedimiento = True
                    enModulo = comp.Name
                    Exit Function
                End If
            End If
        Next l
    Next comp
    On Error GoTo 0
End Function

'=========================================================
' Auxiliar: localiza la hoja que contiene Worksheet_Change
' (es la hoja del menu A6)
'=========================================================
Private Function HojaMenu() As Worksheet
    Dim comp As Object
    Dim cm As Object
    Dim wsH As Worksheet
    Set HojaMenu = Nothing
    On Error Resume Next
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 Then
            Set cm = comp.CodeModule
            If cm.CountOfLines > 0 Then
                If InStr(cm.Lines(1, cm.CountOfLines), "Worksheet_Change") > 0 Then
                    For Each wsH In ThisWorkbook.Worksheets
                        If wsH.CodeName = comp.Name Then
                            Set HojaMenu = wsH
                            Exit Function
                        End If
                    Next wsH
                End If
            End If
        End If
    Next comp
    On Error GoTo 0
End Function
