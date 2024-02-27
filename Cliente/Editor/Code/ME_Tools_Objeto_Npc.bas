Attribute VB_Name = "Me_Tools_Npc"
Option Explicit


'NPC
Public Enum eHerramientasNPC
    insertar = 1
    borrar = 2
End Enum

Public Type tNPCSeleccionado
    Index As Integer
    zona As Byte
End Type

Private Const cantidadSubHerramientaNPC As Byte = 2 ' De 1 a ...
Public herramientaInternaNPC As eHerramientasNPC
Public NPCSeleccionado() As tNPCSeleccionado

'Objetos
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

Public Sub iniciarToolNPC()
    Call seleccionarIndexNPC(0, 0)
End Sub
Public Sub seleccionarCantidadObjeto(cantidad As Integer)
    ReDim objCantidadSeleccionado(1 To 1, 1 To 1) As Integer
    objCantidadSeleccionado(1, 1) = cantidad
End Sub

Public Sub seleccionarIndexObjeto(Index As Integer)
    ReDim objIndexSeleccionado(1 To 1, 1 To 1) As Integer
    objIndexSeleccionado(1, 1) = Index
End Sub

'Criaturas
Public Sub seleccionarIndexNPC(Index As Integer, zona As Byte)

    ReDim NPCSeleccionado(1 To 1, 1 To 1) As tNPCSeleccionado
    NPCSeleccionado(1, 1).Index = Index
    
End Sub
Public Sub rotarHerramientaInternaNPC(paraArriba As Boolean)

    If paraArriba Then
        herramientaInternaNPC = herramientaInternaNPC + 1
        If herramientaInternaNPC > cantidadSubHerramientaNPC Then herramientaInternaNPC = 1
    Else
        herramientaInternaNPC = herramientaInternaNPC - 1
        If herramientaInternaNPC < 1 Then herramientaInternaNPC = cantidadSubHerramientaNPC
    End If

    Call Me_Tools_Objeto_Npc.activarUltimaHerramientaNPC

End Sub

Public Sub activarUltimaHerramientaNPC()
    
    Select Case herramientaInternaNPC
        Case eHerramientasNPC.insertar
            Call Me_Tools_Objeto_Npc.click_InsertarNPC
        Case eHerramientasNPC.borrar
            Call Me_Tools_Objeto_Npc.click_BorrarNPC
    End Select
    
End Sub

Public Sub click_BorrarNPC()
    herramientaInternaNPC = eHerramientasNPC.borrar
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarNpc, Tools.tool_npc)
End Sub

Public Sub click_InsertarNPC()
    If UBound(NPCSeleccionado, 1) = 1 And UBound(NPCSeleccionado, 2) = 1 Then
        If NPCSeleccionado(1, 1).Index > 0 Then
            herramientaInternaNPC = eHerramientasNPC.insertar
            Call ME_Tools.seleccionarTool(frmMain.cmdInsertarNpc, Tools.tool_npc)
        Else
            MsgBox "Debes seleccionar un npc para insertar.", vbInformation
        End If
    End If
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

    
    Call Me_Tools_Objeto_Npc.activarUltimaHerramientaObjeto

End Sub

Public Sub activarUltimaHerramientaObjeto()
    Select Case herramientaInternaOBJ
        Case eHerramientasOBJ.insertar
            Call Me_Tools_Objeto_Npc.click_InsertarOBJ
        Case eHerramientasOBJ.borrar
            Call Me_Tools_Objeto_Npc.click_BorrarOBJ
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
            If MapData(X, Y).OBJInfo.objIndex > 0 Then
                If MapData(X, Y).Graphic(3).GrhIndex = MapData(X, Y).ObjGrh.GrhIndex Then MapData(X, Y).Graphic(3).GrhIndex = 0
                MapData(X, Y).OBJInfo.objIndex = 0
                MapData(X, Y).OBJInfo.Amount = 0
            End If
        Next X
    Next Y

End Sub
