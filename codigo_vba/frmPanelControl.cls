Option Explicit

'=========================================================
' USERFORM: frmPanelControl
' Panel de control centralizado del Sistema Manchester
'=========================================================

'---------------------------------------------------------
' INICIALIZACIÓN DEL FORMULARIO
'---------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Centrar el formulario
    Me.StartUpPosition = 1  ' CenterOwner
    Me.Caption = "Sistema Manchester - Panel de Control"
    Me.Width = 340
    Me.Height = 400
    Me.BackColor = RGB(245, 245, 245)
    
    ' Estilo del título
    With Me.lblTitulo
        .Caption = "Sistema Manchester"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(0, 51, 102)
    End With
    
    With Me.lblSeccionPacientes
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(128, 128, 128)
    End With
    
    With Me.lblSeccionGestion
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(128, 128, 128)
    End With
    
    ' Configurar colores de los botones
    With Me.cmdNuevoPaciente
        .BackColor = RGB(200, 240, 200)
        .ForeColor = RGB(0, 80, 0)
        .Font.Bold = True
    End With
    
    With Me.cmdEliminarUltimo
        .BackColor = RGB(255, 220, 220)
        .ForeColor = RGB(150, 0, 0)
    End With
    
    With Me.cmdEliminarSeleccionados
        .BackColor = RGB(255, 220, 220)
        .ForeColor = RGB(150, 0, 0)
    End With
    
    With Me.cmdNuevoVolumen
        .BackColor = RGB(220, 220, 255)
        .ForeColor = RGB(0, 0, 120)
        .Font.Bold = True
    End With
    
    With Me.cmdEstadisticas
        .BackColor = RGB(255, 245, 200)
        .ForeColor = RGB(100, 80, 0)
        .Font.Bold = True
    End With
    
    With Me.cmdCerrar
        .BackColor = RGB(220, 220, 220)
    End With
End Sub

'---------------------------------------------------------
' BOTÓN: NUEVO PACIENTE
'---------------------------------------------------------
Private Sub cmdNuevoPaciente_Click()
    Unload Me
    On Error GoTo ErrorNuevo
    NuevoPaciente
    Exit Sub
ErrorNuevo:
    MsgBox "Error al crear nuevo paciente: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------
' BOTÓN: ELIMINAR ÚLTIMO
'---------------------------------------------------------
Private Sub cmdEliminarUltimo_Click()
    Unload Me
    On Error GoTo ErrorEliminar
    EliminarUltimoPaciente
    Exit Sub
ErrorEliminar:
    MsgBox "Error al eliminar: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------
' BOTÓN: ELIMINAR SELECCIONADOS
'---------------------------------------------------------
Private Sub cmdEliminarSeleccionados_Click()
    Unload Me
    On Error GoTo ErrorEliminarSel
    EliminarPacienteSeleccionado
    Exit Sub
ErrorEliminarSel:
    MsgBox "Error al eliminar seleccionados: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------
' BOTÓN: NUEVO VOLUMEN
'---------------------------------------------------------
Private Sub cmdNuevoVolumen_Click()
    Unload Me
    On Error GoTo ErrorVolumen
    CrearNuevoVolumenManual
    Exit Sub
ErrorVolumen:
    MsgBox "Error al crear volumen: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------
' BOTÓN: ESTADÍSTICAS
'---------------------------------------------------------
Private Sub cmdEstadisticas_Click()
    Unload Me
    On Error GoTo ErrorEstadisticas
    GenerarEstadisticas
    Exit Sub
ErrorEstadisticas:
    MsgBox "Error al generar estadísticas: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------
' BOTÓN: CERRAR
'---------------------------------------------------------
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

'---------------------------------------------------------
' TECLA ESCAPE PARA CERRAR
'---------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Permitir cierre normal
End Sub
