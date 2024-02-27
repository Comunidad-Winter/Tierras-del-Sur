Attribute VB_Name = "modEntrenador"
Option Explicit

Private Const MAXMASCOTASENTRENADOR = 7

Public Sub solicitarCriatura(personaje As User, nombreCriatura As String)
    Dim npcIndex As Integer
    Dim loopCriatura As Byte
    Dim nuevoNpcIndex As Integer
    
    npcIndex = personaje.flags.TargetNPC
    
    ' ¿Entrenador marcado?
    If npcIndex = 0 Then Exit Sub
        
    ' ¿Es un entrenador?
    If Not NpcList(npcIndex).NPCtype = NPCTYPE_ENTRENADOR Then Exit Sub
        
    ' ¿Supero el limite de mascotas que puede invocar?Para que no floodeen todo el mapa
    If NpcList(npcIndex).Mascotas > MAXMASCOTASENTRENADOR Then
        EnviarPaquete Paquetes.DescNpc, Chr$(2) & ITS(NpcList(npcIndex).Char.charIndex), personaje.UserIndex, ToPCArea, personaje.pos.map
        Exit Sub
    End If
    
    nuevoNpcIndex = 0
    ' Buscamos si la criatura esta en la lista de criaturas que puede invocar
    For loopCriatura = 1 To UBound(NpcList(npcIndex).Criaturas)
        If NpcList(npcIndex).Criaturas(loopCriatura).NpcName = nombreCriatura Then
            nuevoNpcIndex = SpawnNpc(NpcList(npcIndex).Criaturas(loopCriatura).npcIndex, NpcList(npcIndex).pos, True, False)
            Exit For
        End If
    Next
             
    ' ¿Se puedo invocar?
    If nuevoNpcIndex > 0 And nuevoNpcIndex <= MAXNPCS Then
        ' El dueño de la criatura es el entrenador
        NpcList(nuevoNpcIndex).MaestroNpc = npcIndex
        NpcList(npcIndex).Mascotas = NpcList(npcIndex).Mascotas + 1
    End If
            

End Sub
