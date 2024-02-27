Attribute VB_Name = "AntiRoboNpc"
Option Explicit
'Este modulo tiene las funciones necesarias para evitar el robo de npcs entre usuarios de la misma legion
Private Const tiempoPosecion = 30000 ' Tiempo en milisegundos que dura la posecion de una persona sobre un npc
' Esta funcion recibe un npcindex y devuelve
' 0 si no esta luchando con nadie
' el Userindex si esta luchando con alguien
Public Function estaLuchando(criatura As npc) As Integer
    If criatura.UserIndexLucha > 0 Then
        If GetTickCount - criatura.UltimoGolpe > tiempoPosecion Then
            Call resetearLuchador(criatura)
            estaLuchando = 0
        Else
            estaLuchando = criatura.UserIndexLucha
        End If
    Else
        estaLuchando = 0 'No esta luchando con nadie
    End If
End Function

' Este procedimieto hace que el personaje no tenga mas posecion
' sobre el npc
Public Sub resetearLuchador(criatura As npc)
    UserList(criatura.UserIndexLucha).LuchandoNPC = 0
    criatura.UserIndexLucha = 0
    criatura.UltimoGolpe = 0
    Exit Sub
End Sub


Public Function puedePegarleAlNpc(UserIndex As Integer, otroUser As Integer) As Boolean
    'Esta intentando atacar un npc que es de otra persona
    'En caso de ciudadanos, de otros ciudadanos o soldados del ejército
    'En caso de miembros del ejército, de otros miembros o ciudadanos
    'En caso de legionarios, de otros legionarios
    ' La primera chequea que no se roben entre ciudadanos / armadas
    ' La segunda entre dos caos
    
    If UserList(UserIndex).PartyIndex = UserList(otroUser).PartyIndex And UserList(otroUser).PartyIndex > 0 Then
        puedePegarleAlNpc = True
        Exit Function
    End If
    
    If UserList(UserIndex).faccion.alineacion = eAlineaciones.Neutro Then
        puedePegarleAlNpc = True
        Exit Function
    End If
    
    If Not UserList(UserIndex).faccion.alineacion = UserList(otroUser).faccion.alineacion Then
        puedePegarleAlNpc = True
        Exit Function
    End If
    
    puedePegarleAlNpc = False
    
End Function

Public Function puedeLucharContraELNPC(criatura As npc, Usuario As User) As Boolean
    Dim otroUsuario As Integer
    
    If criatura.MaestroUser = 0 And MapInfo(Usuario.pos.map).PermiteRoboNPC = 0 Then
        otroUsuario = estaLuchando(criatura)
        
        If Not otroUsuario = Usuario.UserIndex And otroUsuario > 0 Then
            If Not AntiRoboNpc.puedePegarleAlNpc(Usuario.UserIndex, otroUsuario) Then
                EnviarPaquete Paquetes.mensajeinfo, "No puedes atacar a esta criatura por que esta luchando con " & UserList(otroUsuario).Name, Usuario.UserIndex, ToIndex
                puedeLucharContraELNPC = False
                Exit Function
            End If
        Else
            If Usuario.LuchandoNPC <> criatura.npcIndex And Usuario.LuchandoNPC > 0 Then
                ' Si antes le estaba pegando a otro npc, libero a ese npc
                Call AntiRoboNpc.resetearLuchador(NpcList(Usuario.LuchandoNPC))
            End If
            
            criatura.UltimoGolpe = GetTickCount()
            criatura.UserIndexLucha = Usuario.UserIndex
            Usuario.LuchandoNPC = criatura.npcIndex
            
            puedeLucharContraELNPC = True
            Exit Function
        End If
    End If
    
    puedeLucharContraELNPC = True
End Function

