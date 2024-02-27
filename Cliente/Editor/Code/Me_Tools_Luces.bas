Attribute VB_Name = "Me_Tools_Luces"
Option Explicit

Public Enum eHerramientasLuces
    ninguna = 0
    insertar = 1
    borrar = 2
End Enum

Public infoLuzSeleccionada() As tLuzPropiedades

Public herramientaInternaLuces As eHerramientasLuces
Private Const cantidadSubHerramienta As Byte = 2

Public Sub Clear_Luces_Mapa()
DLL_Luces.Remove_All
'LuzMouse = Engine_Landscape.Light_Create(50, 50, 255, 255, 255, 3)

'Engine_Landscape.Light_Toggle LuzMouse, 0
End Sub



Public Sub mostrarLuzEnFormulario(luz As tLuzPropiedades)
    Dim Color As Long

    'El color
    frmMain.luces_color.BackColor = RGB(luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b)
    
    'El radio
    frmMain.luces_radio.value = luz.LuzRadio
    
    'Tipo de luz
    If (luz.LuzTipo And TipoLuces.Luz_Cuadrada) Then
        frmMain.chkLuzCuadrada.value = 1
    Else
        frmMain.chkLuzCuadrada.value = 0
    End If
    
    If (luz.LuzTipo And TipoLuces.Luz_Fuego) Then
        frmMain.chkAnimacionFuego.value = 1
    Else
        frmMain.chkAnimacionFuego.value = 0
    End If

    'Horario
    If (luz.luzInicio <> 0 Or luz.luzFin <> 0) Then
        frmMain.chkPrendeEn.value = 1
        frmMain.horaInicioLuz.value = luz.luzInicio
        frmMain.horaFinLuz.value = luz.luzFin
    Else
        frmMain.chkPrendeEn.value = 0
        frmMain.horaInicioLuz.value = frmMain.horaInicioLuz.min
        frmMain.horaFinLuz.value = frmMain.horaFinLuz.min
    End If
    
    'Brillo
    If (luz.LuzTipo And TipoLuces.Luz_Normal) Then
        frmMain.chkUtilizarBrillo.value = 1
        frmMain.luz_luminosidad.value = luz.LuzBrillo
    Else
        frmMain.luz_luminosidad.value = frmMain.luz_luminosidad.max / 2
        frmMain.chkUtilizarBrillo.value = 0
    End If

End Sub
Public Sub click_InsertarLuz()
    herramientaInternaLuces = eHerramientasLuces.insertar
    Call ME_Tools.seleccionarTool(frmMain.cmdInsertarLuz, tool_luces)
End Sub

Public Sub click_BorrarLuz()
    herramientaInternaLuces = eHerramientasLuces.borrar
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarLuz, tool_luces)
End Sub

Public Sub seleccionarLuz(luz As tLuzPropiedades)
    ReDim infoLuzSeleccionada(1 To 1, 1 To 1) As tLuzPropiedades
    infoLuzSeleccionada(1, 1) = luz
End Sub


Public Sub rotarHerramientaInterna(paraArriba As Boolean)

    If paraArriba Then
        herramientaInternaLuces = herramientaInternaLuces + 1
        If herramientaInternaLuces > cantidadSubHerramienta Then herramientaInternaLuces = 1
    Else
        herramientaInternaLuces = herramientaInternaLuces - 1
        If herramientaInternaLuces < 1 Then herramientaInternaLuces = cantidadSubHerramienta
    End If
    
   
    Call Me_Tools_Luces.activarUltimaHerramientaLuces

End Sub

Public Sub activarUltimaHerramientaLuces()

    Select Case herramientaInternaLuces
        Case eHerramientasLuces.insertar
            Me_Tools_Luces.click_InsertarLuz
        Case eHerramientasLuces.borrar
            Me_Tools_Luces.click_BorrarLuz
    End Select
    
End Sub

