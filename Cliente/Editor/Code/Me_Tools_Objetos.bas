Attribute VB_Name = "Me_Tools_Objetos"
Option Explicit

Public Enum eHerramientasOBJ
    insertar = 1
    borrar = 2
End Enum

Private Const cantidadSubHerramientaOBJ As Byte = 2 ' De 1 a ...
Public herramientaInternaOBJ As eHerramientasOBJ

Public objCantidadSeleccionado() As Integer
Public objIndexSeleccionado() As Integer


Public Sub iniciarToolObjetos()
    Call seleccionarIndexObjeto(0)
    Call seleccionarCantidadObjeto(1)
End Sub

Public Sub seleccionarCantidadObjeto(cantidad As Integer)
    ReDim objCantidadSeleccionado(1 To 1, 1 To 1) As Integer
    objCantidadSeleccionado(1, 1) = cantidad
End Sub

Public Sub seleccionarIndexObjeto(Index As Integer)
    ReDim objIndexSeleccionado(1 To 1, 1 To 1) As Integer
    objIndexSeleccionado(1, 1) = Index
End Sub

'Objetos
Public Sub rotarHerramientaInternaObjeto(paraArriba As Boolean)

    
    If paraArriba Then
        herramientaInternaOBJ = herramientaInternaOBJ + 1
        If herramientaInternaOBJ > cantidadSubHerramientaOBJ Then herramientaInternaOBJ = 1
    Else
        herramientaInternaOBJ = herramientaInternaOBJ - 1
        If herramientaInternaOBJ < 1 Then herramientaInternaOBJ = cantidadSubHerramientaOBJ
    End If

    
    Call Me_Tools_Objetos.activarUltimaHerramientaObjeto

End Sub

Public Sub activarUltimaHerramientaObjeto()
    Select Case herramientaInternaOBJ
        Case eHerramientasOBJ.insertar
            Call Me_Tools_Objetos.click_InsertarOBJ
        Case eHerramientasOBJ.borrar
            Call Me_Tools_Objetos.click_BorrarOBJ
    End Select
End Sub
Public Sub click_BorrarOBJ()
    herramientaInternaOBJ = eHerramientasOBJ.borrar
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarObjeto, Tools.tool_obj)
End Sub

Public Sub click_InsertarOBJ()
    herramientaInternaOBJ = eHerramientasOBJ.insertar
    Call ME_Tools.seleccionarTool(frmMain.cmdInsertarObjeto, Tools.tool_obj)
End Sub


'TODO: Agregar accion a esto?
'No se llama desde ningun lado, pero la dejo porque es util
Public Sub Quitar_Objetos()

If EditWarning Then Exit Sub

    Dim Y As Integer
    Dim X As Integer
    
    For Y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For X = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
            If mapdata(X, Y).OBJInfo.objIndex > 0 Then
                If mapdata(X, Y).Graphic(3).GrhIndex = mapdata(X, Y).ObjGrh.GrhIndex Then mapdata(X, Y).Graphic(3).GrhIndex = 0
                mapdata(X, Y).OBJInfo.objIndex = 0
                mapdata(X, Y).OBJInfo.Amount = 0
            End If
        Next X
    Next Y

End Sub
