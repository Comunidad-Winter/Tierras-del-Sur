Attribute VB_Name = "InvNpc"
Option Explicit
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'Modulo para controlar los objetos y los inventarios.



Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, ByVal asesino As Integer)
'TIRA TODOS LOS ITEMS DEL NPC
Dim i As Byte
Dim MiObj As obj
    
If npc.Invent.NroItems > 0 Then

        ' Dropeamos los objetos basicos
        For i = 1 To MAX_INVENTORY_SLOTS
            If npc.Invent.Object(i).ObjIndex > 0 Then
                  MiObj.Amount = npc.Invent.Object(i).Amount
                  MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
                  Call TirarItemAlPisoConDuenio(npc.pos, MiObj, asesino)
            End If
        Next i
    
End If
    
If npc.Invent.NroItemsDrop > 0 Then
    
    ' Revisamos los objetos que tienen probailidad.
    For i = 1 To MAX_DROP
        If npc.Invent.ObjectDrop(i).ObjIndex > 0 Then
            If npc.Invent.ObjectDrop(i).Probability >= RandomNumber(1, 100) Then
                MiObj.Amount = npc.Invent.ObjectDrop(i).Amount
                MiObj.ObjIndex = npc.Invent.ObjectDrop(i).ObjIndex
                Call TirarItemAlPisoConDuenio(npc.pos, MiObj, asesino)
            End If
        End If
    Next i
End If
    
End Sub

Function QuedanItems(ByVal npcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)
Dim i As Integer
If NpcList(npcIndex).Invent.NroItems > 0 Then
    For i = 1 To MAX_INVENTORY_SLOTS
        If NpcList(npcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If
QuedanItems = False
End Function

Function EncontrarCant(ByVal npcIndex As Integer, ByVal ObjIndex As Integer) As Integer
'Devuelve la cantidad original del obj de un npc
Dim ln As String, npcfile As String
Dim i As Integer

If NpcList(npcIndex).numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If
For i = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(npcfile, "NPC" & NpcList(npcIndex).numero, "Obj" & i)
    If ObjIndex = val(ReadField(1, ln, 45)) Then
        EncontrarCant = val(ReadField(2, ln, 45))
        Exit Function
    End If
Next
EncontrarCant = 50
End Function

Sub ResetNpcInv(ByVal npcIndex As Integer)
Dim i As Integer

NpcList(npcIndex).Invent.NroItems = 0
For i = 1 To MAX_INVENTORY_SLOTS
   NpcList(npcIndex).Invent.Object(i).ObjIndex = 0
   NpcList(npcIndex).Invent.Object(i).Amount = 0
Next i
NpcList(npcIndex).InvReSpawn = 0
End Sub

Sub QuitarNpcInvItem(ByVal npcIndex As Integer, ByVal slot As Byte, ByVal cantidad As Integer)
Dim ObjIndex As Integer
ObjIndex = NpcList(npcIndex).Invent.Object(slot).ObjIndex

    'Quita un Obj
    If ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).Crucial = 0 Then
        NpcList(npcIndex).Invent.Object(slot).Amount = NpcList(npcIndex).Invent.Object(slot).Amount - cantidad
        If NpcList(npcIndex).Invent.Object(slot).Amount <= 0 Then
            NpcList(npcIndex).Invent.NroItems = NpcList(npcIndex).Invent.NroItems - 1
            NpcList(npcIndex).Invent.Object(slot).ObjIndex = 0
            NpcList(npcIndex).Invent.Object(slot).Amount = 0
            If NpcList(npcIndex).Invent.NroItems = 0 And NpcList(npcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(npcIndex) 'Reponemos el inventario
            End If
        End If
    Else
        NpcList(npcIndex).Invent.Object(slot).Amount = NpcList(npcIndex).Invent.Object(slot).Amount - cantidad
        If NpcList(npcIndex).Invent.Object(slot).Amount <= 0 Then
            NpcList(npcIndex).Invent.NroItems = NpcList(npcIndex).Invent.NroItems - 1
            NpcList(npcIndex).Invent.Object(slot).ObjIndex = 0
            NpcList(npcIndex).Invent.Object(slot).Amount = 0
            If Not QuedanItems(npcIndex, ObjIndex) Then
                   NpcList(npcIndex).Invent.Object(slot).ObjIndex = ObjIndex
                   NpcList(npcIndex).Invent.Object(slot).Amount = EncontrarCant(npcIndex, ObjIndex)
                   NpcList(npcIndex).Invent.NroItems = NpcList(npcIndex).Invent.NroItems + 1
            
            End If
            If NpcList(npcIndex).Invent.NroItems = 0 And NpcList(npcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(npcIndex) 'Reponemos el inventario
            End If
        End If
    End If
End Sub

Sub CargarInvent(ByVal npcIndex As Integer)
'Vuelve a cargar el inventario del npc NpcIndex
Dim loopC As Integer
Dim ln As String
Dim npcfile As String

If NpcList(npcIndex).numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If

NpcList(npcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcList(npcIndex).numero, "NROITEMS"))

For loopC = 1 To NpcList(npcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & NpcList(npcIndex).numero, "Obj" & loopC)
    NpcList(npcIndex).Invent.Object(loopC).ObjIndex = val(ReadField(1, ln, 45))
    NpcList(npcIndex).Invent.Object(loopC).Amount = val(ReadField(2, ln, 45))
Next loopC


End Sub

Public Function TirarOroNPc(pos As WorldPos, obj As obj) As WorldPos

    Dim NuevaPos As WorldPos
    NuevaPos.x = 0
    NuevaPos.y = 0
    
    Call TileLibreParaObjeto(pos, NuevaPos, obj)
    
    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
        Call MakeObj(ToMap, 0, pos.map, obj, pos.map, NuevaPos.x, NuevaPos.y)
        TirarOroNPc = NuevaPos
    End If
    
End Function

