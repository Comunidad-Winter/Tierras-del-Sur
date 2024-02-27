Attribute VB_Name = "ME_Tools_Acciones"
Option Explicit

Public Enum eHerramientasAccion
    ninguna = 0
    insertar = 1
    borrar = 2
End Enum

Private Const cantidadSubHerramienta As Byte = 2
Public herramientainterna As eHerramientasAccion
Public accionSeleccionada() As iAccionEditor
Public Sub iniciarToolAcciones()
    Call seleccionarAccion(Nothing)
End Sub
Public Sub seleccionarAccion(accion As iAccionEditor)
    ReDim accionSeleccionada(1 To 1, 1 To 1)
    Set accionSeleccionada(1, 1) = accion
End Sub

Public Sub click_InsertarAccion()
    'If Not accionSeleccionada Is Nothing Then
        herramientainterna = eHerramientasAccion.insertar
        Call ME_Tools.seleccionarTool(frmMain.cmdInsertarAccion, tool_acciones)
    'Else
    '    herramientainterna = Ninguna
    'End If
End Sub

Public Sub click_InsertarBorrarAccion()
    herramientainterna = eHerramientasAccion.borrar
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarAccion, tool_acciones)
End Sub


Public Sub rotarHerramientaInterna(paraArriba As Boolean)
    If paraArriba Then
        herramientainterna = herramientainterna + 1
        If herramientainterna > cantidadSubHerramienta Then herramientainterna = 1
    Else
        herramientainterna = herramientainterna - 1
        If herramientainterna < 1 Then herramientainterna = cantidadSubHerramienta
    End If
    
    Call ME_Tools_Acciones.activarUltimaHerramientaAcciones
End Sub

Public Sub activarUltimaHerramientaAcciones()
    Select Case herramientainterna
        Case eHerramientasAccion.insertar
            ME_Tools_Acciones.click_InsertarAccion
        Case eHerramientasAccion.borrar
            ME_Tools_Acciones.click_InsertarBorrarAccion
    End Select
End Sub
