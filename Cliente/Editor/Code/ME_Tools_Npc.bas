Attribute VB_Name = "Me_Tools_Npc"
Option Explicit


'NPC
Public Enum eHerramientasNPC
    insertar = 1
    borrar = 2
End Enum

Public Type tNPCSeleccionado
    Index As Integer
    Zona As Byte
End Type

Private Const cantidadSubHerramientaNPC As Byte = 2 ' De 1 a ...
Public herramientaInternaNPC As eHerramientasNPC
Public NPCSeleccionado() As tNPCSeleccionado

Public Sub iniciarToolNPC()
    Call seleccionarIndexNPC(0, 0)
End Sub

'Criaturas
Public Sub seleccionarIndexNPC(Index As Integer, Zona As Byte)

    ReDim NPCSeleccionado(1 To 1, 1 To 1) As tNPCSeleccionado
    NPCSeleccionado(1, 1).Index = Index
    NPCSeleccionado(1, 1).Zona = Zona
End Sub
Public Sub rotarHerramientaInternaNPC(paraArriba As Boolean)

    If paraArriba Then
        herramientaInternaNPC = herramientaInternaNPC + 1
        If herramientaInternaNPC > cantidadSubHerramientaNPC Then herramientaInternaNPC = 1
    Else
        herramientaInternaNPC = herramientaInternaNPC - 1
        If herramientaInternaNPC < 1 Then herramientaInternaNPC = cantidadSubHerramientaNPC
    End If

    Call Me_Tools_Npc.activarUltimaHerramientaNPC

End Sub

Public Sub activarUltimaHerramientaNPC()
    
    Select Case herramientaInternaNPC
        Case eHerramientasNPC.insertar
            Call Me_Tools_Npc.click_InsertarNPC
        Case eHerramientasNPC.borrar
            Call Me_Tools_Npc.click_BorrarNPC
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

Public Sub cargarZonasDeNacimiento(zonas() As ZonaNacimientoCriatura)
    Dim loopZona As Byte
        
    ' Limpiamos la lista
    Call frmMain.lstZonaNacimientoCriaturas.Clear
    
    ' Agregamos el generico
    Call frmMain.lstZonaNacimientoCriaturas.AddItem("En cualquier lugar")
    frmMain.lstZonaNacimientoCriaturas.Selected(0) = True
    
    ' Agregamos las zonas creadas
    For loopZona = LBound(zonas) To UBound(zonas)
        If Not zonas(loopZona).nombre = "" Then
            Call frmMain.lstZonaNacimientoCriaturas.AddItem(loopZona + 1 & " - " & zonas(loopZona).nombre)
        End If
    Next loopZona
    
End Sub


Public Function calcularZonaLista(lista As VB.ListBox) As Byte
    Dim Zona As Byte
    Dim loopZona As Byte
    Dim idZona As Byte
    
    Zona = 0
    
    For loopZona = 1 To lista.ListCount - 1 ' La 0 es la generica
        If lista.Selected(loopZona) Then
        
            idZona = CByte(val(mid$(lista.list(loopZona), 1, InStr(1, lista.list(loopZona), " "))))
        
            Zona = Zona Or bitwisetable(idZona - 1)
        End If
    Next loopZona

    calcularZonaLista = Zona
End Function


Public Function obtenerDescripcionAbreviatura(Zona As Byte, zonas() As ZonaNacimientoCriatura) As String

    Dim loopC As Byte
    Dim idZona As Byte
    
    obtenerDescripcionAbreviatura = ""

    For loopC = 0 To UBound(zonas)
        If Not mapinfo.ZonasNacCriaturas(loopC).nombre = "" Then
            If (Zona And bitwisetable(loopC)) Then
                obtenerDescripcionAbreviatura = obtenerDescripcionAbreviatura & (loopC + 1)
            End If
        End If
    Next loopC

End Function


Public Function obtenerDescripcionAbreviaturaTile(ByVal X As Integer, ByVal Y As Integer, zonas() As ZonaNacimientoCriatura) As String

    Dim loopC As Byte
    Dim idZona As Byte
    
    obtenerDescripcionAbreviaturaTile = ""

    For loopC = 0 To UBound(zonas)
        If Not mapinfo.ZonasNacCriaturas(loopC).nombre = "" Then
            If zonas(loopC).Superior.X <= X And zonas(loopC).Inferior.X >= X Then
                If zonas(loopC).Superior.Y <= Y And zonas(loopC).Inferior.Y >= Y Then
                    obtenerDescripcionAbreviaturaTile = obtenerDescripcionAbreviaturaTile & (loopC + 1)
                End If
            End If
        End If
    Next loopC

End Function

