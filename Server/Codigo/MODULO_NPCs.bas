Attribute VB_Name = "NPCs"
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private npcIndexLibres As EstructurasLib.ColaConBloques



Private Inte_Basica As Inteligencia_Basica

Public tiempoLlamadaNpcs As Long

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal npcIndex As Integer)
Dim i As Integer
UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasIndex(i) = npcIndex Then
     UserList(UserIndex).MascotasIndex(i) = 0
     UserList(UserIndex).MascotasType(i) = 0
     Exit For
  End If
Next i
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)

NpcList(Maestro).Mascotas = NpcList(Maestro).Mascotas - 1

End Sub

Public Function DeboEnviarAngulo(ByVal mapa As Integer) As Boolean

    If MapInfo(mapa).usuarios.getCantidadElementos < 10 Then
        DeboEnviarAngulo = True
    Else
        DeboEnviarAngulo = False
    End If
    
End Function

Public Sub UsuarioMataNPC(ByRef asesino As User, ByRef criatura As npc)

    asesino.flags.TargetNPC = 0
    asesino.flags.TargetNpcTipo = 0

     'El user que lo mato tiene mascotas?
     If asesino.NroMacotas > 0 Then
        Dim t As Integer
        For t = 1 To MAXMASCOTAS
            If asesino.MascotasIndex(t) > 0 Then
                If NpcList(asesino.MascotasIndex(t)).TargetNPCID = criatura.npcIndex Then
                    Call FollowAmo(asesino.MascotasIndex(t))
                End If
            End If
        Next t
     End If
    
    ' Experiencia final
     If criatura.flags.ExpCount > 0 Then
        Call CalcularDarExpUltimoGolpe(asesino, criatura)
     Else
        EnviarPaquete Paquetes.MensajeFight, "No has ganado experiencia al matar la criatura.", asesino.UserIndex
     End If
     '
     EnviarPaquete Paquetes.MensajeSimple, Chr(25), asesino.UserIndex
     
     'cambiar
     asesino.Stats.NPCsMuertos = asesino.Stats.NPCsMuertos + 1
    
    If criatura.Hostil = False Then
         ' Al no poder sacarte los puntos de asesino, las criaturas no hostiles ahora entregaran de bandidos.
         Call AddtoVar(asesino.Reputacion.BandidoRep, vlASESINO, MAXREP)
     Else
        Call AddtoVar(asesino.Reputacion.NobleRep, vlASESINO / 2, MAXREP)
     End If
          
     ' Si no son mascotas tira oro.
     If criatura.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCs.npcEntregarOro(criatura, asesino)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(criatura, asesino.UserIndex)
    Else
        Call LogHack("Se asesina a una mascota")
    End If

End Sub

Function MuereNpc(ByRef criatura As npc, Optional ByVal NoResPawn As Byte) As Integer


 ' Sonido de muerte
If criatura.flags.Snd3 > 0 Then
    EnviarPaquete Paquetes.WavSnd, Chr$(criatura.flags.Snd3), criatura.npcIndex, ToNPCArea
End If

 'ReSpawn o no. Si es mascota de un usuario, no respawneo
If NoResPawn = 0 And criatura.MaestroUser = 0 Then
     MuereNpc = ReSpawnNpc(criatura)
 End If
 
'Quitamos el npc
Call QuitarNPC(criatura.npcIndex)
 
End Function

Sub ResetNpcFlags(ByVal npcIndex As Integer)
With NpcList(npcIndex)
'Clear the npc's flags
    .flags.AfectaParalisis = 0
    .flags.BackUp = 0
    .flags.Domable = 0
    .flags.OldMovement = 0
    .flags.Paralizado = 0
    .flags.Inmovilizado = 0
    .flags.Respawn = 0
    .flags.Snd1 = 0
    .flags.Snd2 = 0
    .flags.Snd3 = 0
End With
End Sub

Sub ResetNpcCounters(ByVal npcIndex As Integer)
    NpcList(npcIndex).Contadores.Paralisis = 0
    NpcList(npcIndex).Contadores.TiempoExistencia = 0
    NpcList(npcIndex).Contadores.TiempoUltimoAtaque = 0
End Sub

Sub ResetNpcCharInfo(ByVal npcIndex As Integer)
With NpcList(npcIndex)
    .Char.Body = 0
    .Char.CascoAnim = 0
    .Char.charIndex = 0
    .Char.FX = 0
    .Char.Head = 0
    .Char.heading = 0
    .Char.loops = 0
    .Char.ShieldAnim = 0
    .Char.WeaponAnim = 0
End With
End Sub

Sub ResetNpcCriatures(ByVal npcIndex As Integer)
Dim j As Integer
For j = 1 To NpcList(npcIndex).NroCriaturas
    NpcList(npcIndex).Criaturas(j).npcIndex = 0
    NpcList(npcIndex).Criaturas(j).NpcName = ""
Next j
NpcList(npcIndex).NroCriaturas = 0
End Sub

Sub ResetNpcMainInfo(ByVal npcIndex As Integer)

Dim j As Integer

With NpcList(npcIndex)
    .Attackable = 0
    .Comercia = 0
    .GiveEXP = 0
    .GiveGLD = 0
    .Inflacion = 0
    .InvReSpawn = 0
    
    If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, npcIndex)
    If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc, npcIndex)
    If NpcList(npcIndex).UserIndexLucha > 0 Then Call AntiRoboNpc.resetearLuchador(NpcList(npcIndex))
    
    .MaestroUser = 0
    .MaestroNpc = 0
    .Mascotas = 0
    .Movement = 0
    Set .Inteligencia = Nothing
    .Name = "NPC SIN INICIAR"
    .NPCtype = 0
    .numero = 0
    .Orig.map = 0
    .Orig.x = 0
    .Orig.y = 0
    .PoderAtaque = 0
    .PoderEvasion = 0
    .pos.map = 0
    .pos.x = 0
    .pos.y = 0
    .TargetNPCID = 0
    .TargetUserID = 0
    .TipoItems = 0
    .Veneno = 0
    .desc = ""
    .InmuneAHechizos = 0

    For j = 1 To .NroSpells
        .Spells(j) = 0
    Next j
    
    .Nivel = 0
End With
Call ResetNpcCharInfo(npcIndex)
Call ResetNpcCriatures(npcIndex)
End Sub

Sub QuitarNPC(ByVal npcIndex As Integer)

'Esta index esta libre.
Call npcIndexLibres.agregar(npcIndex)

If SV_PosicionesValidas.existePosicionMundo(NpcList(npcIndex).pos.map, NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y) Then
    Call EraseNPCChar(NpcList(npcIndex).pos.map, npcIndex)
End If
'Nos aseguramos de que el inventario sea removido...
'asi los lobos no volveran a tirar armaduras ;))
Call ResetNpcInv(npcIndex)
Call ResetNpcFlags(npcIndex)
Call ResetNpcCounters(npcIndex)
Call ResetNpcMainInfo(npcIndex)

NumNPCs = NumNPCs - 1

End Sub

Public Function TestSpawnTrigger(pos As WorldPos) As Boolean
    TestSpawnTrigger = (MapData(pos.map, pos.x, pos.y).Trigger And eTriggers.AntiRespawnNpc) = False And (MapData(pos.map, pos.x, pos.y).Trigger And eTriggers.PosicionInvalidaNpc) = False
End Function

'Crea un NPC en el mapa
'En cualquier posicion del mapa que sea valida.
'Tiene privilegio aquella pocion donde no haya un usuario mirando
Function CrearNPC(NroNPC As Integer, mapa As WorldPos, OrigPos As WorldPos, enOrigen As Boolean) As Integer


Dim pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim NPCCercano As Boolean
Dim minX As Integer
Dim minY As Integer
Dim maxX As Integer
Dim maxY As Integer

altpos.map = 0

nIndex = OpenNPC(NroNPC) 'Conseguimos un indice

If nIndex > MAXNPCS Then Exit Function
'Necesita ser respawned en un lugar especifico
If enOrigen And SV_PosicionesValidas.existePosicionMundo(OrigPos.map, OrigPos.x, OrigPos.y) Then
    NpcList(nIndex).Orig = OrigPos
    NpcList(nIndex).pos = OrigPos
Else
    pos.map = mapa.map 'mapa
    altpos.map = mapa.map
    
    If DeboEnviarAngulo(mapa.map) Then
        Do While Not PosicionValida And Iteraciones <= MAXSPAWNATTEMPS
        
            'Obtengo un X de rango 20 sin salirme del limite
            minX = maxi(mapa.x - 20, SV_Constantes.X_MINIMO_USABLE)
                
            'Obtengo un X de rango 20 sin salirme del limite mayor
            maxX = mini(mapa.x + 20, SV_Constantes.X_MAXIMO_USABLE)
                
            'Obtengo un Y de rango 20 sin salirme del limite
            minY = maxi(mapa.y - 20, SV_Constantes.Y_MINIMO_USABLE)
            
            'Obtengo un Y de rango 20 sin salirme del limite
            maxY = mini(mapa.y + 20, SV_Constantes.Y_MAXIMO_USABLE)
            
            pos.x = RandomNumber(minX, maxX)
            pos.y = RandomNumber(minY, maxY)
            
            Call ClosestLegalPosNPC(pos, newpos, NpcList(nIndex))
            
            'Encontro una posicion valida?
            If newpos.map > 0 Then
                'Hay algun usuario viendo?
                If Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
                    'Asignamos las nuevas coordenas solo si son validas
                    NpcList(nIndex).pos.map = newpos.map
                    NpcList(nIndex).pos.x = newpos.x
                    NpcList(nIndex).pos.y = newpos.y
                    PosicionValida = True
                Else
                    'Guardo la posicion que es valido, por las dudas que no haya un lugar donde
                    'no hay nadie mirando
                    altpos = newpos
                End If
            End If
            'for debug
            Iteraciones = Iteraciones + 1
        Loop
    Else
        Do While Not PosicionValida And Iteraciones <= MAXSPAWNATTEMPS
            'Obtengo pociones al azar entre el mapa
            pos.x = SV_Constantes.X_MINIMO_USABLE + CInt(Rnd * (SV_Constantes.X_MAXIMO_USABLE - SV_Constantes.X_MINIMO_USABLE) + 1)
            pos.y = SV_Constantes.Y_MINIMO_USABLE + CInt(Rnd * (SV_Constantes.Y_MAXIMO_USABLE - SV_Constantes.Y_MINIMO_USABLE) + 1)
            
            Call ClosestLegalPosNPC(pos, newpos, NpcList(nIndex))
            
            'Encontro una posicion valida?
            If newpos.map > 0 Then
                'Hay algun usuario viendo?
                If Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
                    'Asignamos las nuevas coordenas solo si son validas
                    NpcList(nIndex).pos.map = newpos.map
                    NpcList(nIndex).pos.x = newpos.x
                    NpcList(nIndex).pos.y = newpos.y
                    PosicionValida = True
                Else
                    'Guardo la posicion que es valido, por las dudas que no haya un lugar donde
                    'no hay nadie mirando
                    altpos = newpos
                End If
            End If
            'for debug
            Iteraciones = Iteraciones + 1
        Loop
    End If
    
    If Iteraciones > MAXSPAWNATTEMPS Then
    'No encontre ninguna posicion ideal.
    'Veo si al menos encontre una legal
        If altpos.map > 0 Then
            NpcList(nIndex).pos.map = altpos.map
            NpcList(nIndex).pos.x = altpos.x
            NpcList(nIndex).pos.y = altpos.y
        Else
            'Una ultima chance! Al centro del mapa
            pos.x = 50
            pos.y = 50
            Call ClosestLegalPosNPC(pos, newpos, NpcList(nIndex))
            'Encontre!!
            If newpos.map > 0 Then
                NpcList(nIndex).pos.map = newpos.map
                NpcList(nIndex).pos.x = newpos.x
                NpcList(nIndex).pos.y = newpos.y
            Else
                'No encontre, chau npc
                Call QuitarNPC(nIndex)
                Call Logs.LogProblemaSpawn(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa.map & " NroNpc:" & NroNPC)
                Exit Function
            End If
        End If
    End If
End If
'Si llegue hasta ca es porque encontre una posicion valida
'Crea el NPC

Call MakeNPCChar(ToMap, 0, NpcList(nIndex).pos.map, nIndex, NpcList(nIndex).pos.map, NpcList(nIndex).pos.x, NpcList(nIndex).pos.y)

CrearNPC = nIndex

End Function

Public Sub EnviarNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, npcIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    EnviarPaquete pCrearNPC, ITS(NpcList(npcIndex).Char.charIndex) & ITS(NpcList(npcIndex).Char.Body) & ITS(NpcList(npcIndex).Char.Head) & ByteToString(NpcList(npcIndex).Char.heading) & Chr(x) & Chr(y), sndIndex, sndRoute, sndMap
End Sub

Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, npcIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
Dim charIndex As Integer

If NpcList(npcIndex).Char.charIndex = 0 Then
    charIndex = NextOpenCharIndex
    NpcList(npcIndex).Char.charIndex = charIndex
    CharList(charIndex) = npcIndex
End If

'Agrego el npc al mapa
MapData(map, x, y).npcIndex = npcIndex
'Me aseguro la relacion npc mapa
NpcList(npcIndex).pos.map = map
NpcList(npcIndex).pos.x = x
NpcList(npcIndex).pos.y = y

'If Npclist(MapData(Map, X, Y).NpcIndex).Movement <> ESTATICO Then
Call MapInfo(map).NPCs.agregar(npcIndex)
'End If

If sndRoute <> ToNone Then 'byGorlok
    Call EnviarNPCChar(sndRoute, sndIndex, sndMap, npcIndex, x, y)
End If

End Sub

Sub ChangeHeadingNpc(npcIndex As Integer, ByVal heading As Byte)
With NpcList(npcIndex)
    If (.Char.heading <> heading) Then
        .Char.heading = heading
        EnviarPaquete Paquetes.CambiarHeadingNpc, ITS(.Char.charIndex) & heading, npcIndex, ToAreaNPC, .pos.map
    End If
End With
End Sub

Sub EraseNPCChar(sndMap As Integer, ByVal npcIndex As Integer)

If NpcList(npcIndex).Char.charIndex <> 0 Then
    CharList(NpcList(npcIndex).Char.charIndex) = 0
End If

'Quitamos del mapa
MapData(NpcList(npcIndex).pos.map, NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y).npcIndex = 0

Call MapInfo(NpcList(npcIndex).pos.map).NPCs.eliminar(npcIndex)
'Actualizamos los cliente

EnviarPaquete Paquetes.BorrarUser, ITS(NpcList(npcIndex).Char.charIndex), 0, ToMap, NpcList(npcIndex).pos.map
'Update la lista npc

NpcList(npcIndex).Char.charIndex = 0

'Pos del npc
NpcList(npcIndex).pos.map = 0
NpcList(npcIndex).pos.x = 0
NpcList(npcIndex).pos.y = 0
'update NumChars
NumChars = NumChars - 1

End Sub

Sub MoveNPCChar(ByVal npcIndex As Integer, ByVal nHeading As Byte)

    Dim nPos As WorldPos
    If nHeading = 0 Then Exit Sub
    
    nPos = NpcList(npcIndex).pos
    
    Call HeadtoPos(nHeading, nPos)
   
   ' es una posicion legal
    If SV_PosicionesValidas.esPosicionJugable(CByte(nPos.x), CByte(nPos.y)) And SV_PosicionesValidas.esPosicionUsableNPC(MapData(NpcList(npcIndex).pos.map, nPos.x, nPos.y), NpcList(npcIndex)) Then
        'Actualizao la posicion del npc
        'Saco en la vieja
        MapData(NpcList(npcIndex).pos.map, NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y).npcIndex = 0
        NpcList(npcIndex).pos = nPos
        NpcList(npcIndex).Char.heading = nHeading
        'Pongo en la nueva
        MapData(NpcList(npcIndex).pos.map, NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y).npcIndex = npcIndex
        'Aviso a los usuarios
        EnviarPaquete Paquetes.MoveNpc, ITS(NpcList(npcIndex).Char.charIndex) & Chr$(NpcList(npcIndex).pos.x) & Chr$(NpcList(npcIndex).pos.y), npcIndex, ToAreaNPC, NpcList(npcIndex).pos.map
    Else
        'No es posicion valida, solo muevo la cabeza.
        Call ChangeHeadingNpc(npcIndex, nHeading)
        
        If NpcList(npcIndex).Movement = NPC_PATHFINDING Then
            'Someone has blocked the npc's way, we must to seek a new path!
            NpcList(npcIndex).PFINFO.PathLenght = 0
        End If
    End If


End Sub

Function NextOpenNPC() As Integer

If npcIndexLibres.getCantidadElementos > 0 Then
    NextOpenNPC = npcIndexLibres.sacar()
Else
    NextOpenNPC = MAXNPCS + 1 'Un valor invalido
End If

End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
Dim N As Integer
N = RandomNumber(1, 100)
If N < 30 Then
    UserList(UserIndex).flags.Envenenado = 1
    EnviarPaquete Paquetes.EstaEnvenenado, "", UserIndex, ToIndex
    EnviarPaquete Paquetes.MensajeSimple, Chr$(46), UserIndex
End If
End Sub

Function SpawnNpc(ByVal NpcNum As Integer, pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer

Dim newpos As WorldPos
Dim nIndex As Integer

nIndex = OpenNPC(NpcNum, Respawn)   'Conseguimos un indice

If nIndex > MAXNPCS Then
    SpawnNpc = nIndex
    Exit Function
End If

'Nos devuelve la posicion valida mas cercana
Call ClosestLegalPosNPC(pos, newpos, NpcList(nIndex))
'Si X e Y son iguales a 0 significa que no se encontro posicion valida
If newpos.map > 0 Then
    'Asignamos las nuevas coordenas solo si son validas
    NpcList(nIndex).pos.map = newpos.map
    NpcList(nIndex).pos.x = newpos.x
    NpcList(nIndex).pos.y = newpos.y
    
    'Crea el NPC
    Call MakeNPCChar(ToMap, 0, newpos.map, nIndex, newpos.map, newpos.x, newpos.y)
    
    If FX Then
        EnviarPaquete Paquetes.HechizoFX, ITS(NpcList(nIndex).Char.charIndex) & ByteToString(FXWARP) & ITS(0) & Chr$(SND_WARP), nIndex, ToNPCArea, newpos.map
    End If
    
    SpawnNpc = nIndex

Else
'No encontre una posicion para spawnear al npc
    Call Logs.LogProblemaSpawn("No se encontro una posicion valida en el mapa:" & pos.map & " para el npc:" & NpcList(nIndex).Name)
    
    Call QuitarNPC(nIndex)
    SpawnNpc = MAXNPCS + 1 'Valor invalido
End If

End Function

Function ReSpawnNpc(ByRef MiNPC As npc) As Integer
    If (MiNPC.flags.Respawn = 0) Then
        ReSpawnNpc = CrearNPC(MiNPC.numero, MiNPC.pos, MiNPC.Orig, True)
    End If
End Function

Public Sub ReSpawnNpcByData(ByVal numero As Integer, ByRef posicionActual As WorldPos, ByRef posicionOriginal As WorldPos)
    Call CrearNPC(numero, posicionActual, posicionOriginal, True)
End Sub

Sub npcEntregarOro(ByRef npc As npc, personaje As User)

    Dim cantidad As Long
    
    cantidad = npc.GiveGLD

    If cantidad = 0 Then
        Exit Sub
    End If
    
    If personaje.PartyIndex > 0 Then
        Call mdParty.entregarOro(personaje.PartyIndex, cantidad)
    Else
        Call modUsuarios.agregarOro(personaje, cantidad)
        EnviarPaquete Paquetes.mensajeinfo, "Has ganado " & cantidad & " monedas de oro.", personaje.UserIndex, ToIndex
    End If

End Sub

Sub NPCTirarOro(ByRef MiNPC As npc, UserIndex As Integer)
'SI EL NPC TIENE ORO LO TIRAMOS
Dim MiObj As obj
Dim cantidad As Long

cantidad = MiNPC.GiveGLD

Do While cantidad > 0
    If cantidad - 10000 > 0 Then
        MiObj.Amount = 10000
        cantidad = cantidad - 10000
    Else
        MiObj.Amount = cantidad
        cantidad = 0
    End If
    
    MiObj.ObjIndex = iORO
    
    Call TirarOroNPc(MiNPC.pos, MiObj)
Loop
End Sub

Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer los NPCS se deberá usar la
'nueva clase clsLeerInis.
'
'Alejo
'
'###################################################
Dim npcIndex As Integer
Dim Leer As clsLeerInis
Dim faccion As String
Dim Terreno As String
Dim loopC As Integer
Dim ln As String
Dim tempbyte As Byte

If NpcNumber > 499 Then
    Set Leer = LeerNPCsHostiles
Else
    Set Leer = LeerNPCs
End If


npcIndex = NextOpenNPC

If npcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = npcIndex
    Exit Function
End If

With NpcList(npcIndex)

    .npcIndex = npcIndex  ' ok
    .numero = NpcNumber  ' ok
    .Name = Leer.DarValor("NPC" & NpcNumber, "Name")  ' ok
    .desc = Leer.DarValor("NPC" & NpcNumber, "Desc")  ' ok
    .Movement = val(Leer.DarValor("NPC" & NpcNumber, "Movement")) '
    .flags.OldMovement = .Movement
    .NPCtype = val(Leer.DarValor("NPC" & NpcNumber, "NpcType")) '
    .Char.Body = val(Leer.DarValor("NPC" & NpcNumber, "Body")) '
    .Char.Head = val(Leer.DarValor("NPC" & NpcNumber, "Head")) '
    .Char.heading = val(Leer.DarValor("NPC" & NpcNumber, "Heading")) '
    .Attackable = val(Leer.DarValor("NPC" & NpcNumber, "Attackable")) '
    .Comercia = val(Leer.DarValor("NPC" & NpcNumber, "Comercia")) '
    
    .Veneno = val(Leer.DarValor("NPC" & NpcNumber, "Veneno")) '
    .flags.Domable = val(Leer.DarValor("NPC" & NpcNumber, "Domable")) '
    
    
    #If TDSFacil Then
        .GiveGLD = val(Leer.DarValor("NPC" & NpcNumber, "GiveGLD")) * 60
        .GiveEXP = val(Leer.DarValor("NPC" & NpcNumber, "GiveEXP")) * 999
    #Else
        .GiveGLD = val(Leer.DarValor("NPC" & NpcNumber, "GiveGLD")) '
        .GiveEXP = val(Leer.DarValor("NPC" & NpcNumber, "GiveEXP")) '
    #End If
    
    .flags.ExpCount = .GiveEXP
    
    .PoderAtaque = val(Leer.DarValor("NPC" & NpcNumber, "PoderAtaque")) '
    .PoderEvasion = val(Leer.DarValor("NPC" & NpcNumber, "PoderEvasion")) '
    .InvReSpawn = val(Leer.DarValor("NPC" & NpcNumber, "InvReSpawn")) '
    .Stats.MaxHP = val(Leer.DarValor("NPC" & NpcNumber, "MaxHP")) '
    .Stats.minHP = val(Leer.DarValor("NPC" & NpcNumber, "MinHP")) '
    .Stats.MaxHIT = val(Leer.DarValor("NPC" & NpcNumber, "MaxHIT")) '
    .Stats.MinHIT = val(Leer.DarValor("NPC" & NpcNumber, "MinHIT")) '
    .Stats.Def = val(Leer.DarValor("NPC" & NpcNumber, "DEF")) '
    
    tempbyte = CByte(val(Leer.DarValor("NPC" & NpcNumber, "Alineacion")))
    
    If tempbyte = 2 Then
        .Hostil = True
    Else
        .Hostil = False
    End If
       
    .InmuneAHechizos = val(Leer.DarValor("NPC" & NpcNumber, "InmuneAHechizos")) '
    
    faccion = Leer.DarValor("NPC" & NpcNumber, "Faccion") '
    
    If faccion = "NEUTRO" Or faccion = "" Then '
        .faccion = eAlineaciones.Neutro '
    ElseIf faccion = "REAL" Then '
        .faccion = eAlineaciones.Real '
    ElseIf faccion = "CAOS" Then '
        .faccion = eAlineaciones.caos '
    End If
  
    Terreno = Leer.DarValor("NPC" & NpcNumber, "Terreno") '
  
    If Terreno = "TIERRA" Then '
        .flags.Terreno = eTerrenoNPC.Tierra '
    ElseIf Terreno = "AGUA" Then '
        .flags.Terreno = eTerrenoNPC.Agua '
    Else '
        .flags.Terreno = eTerrenoNPC.AguayTierra '
    End If
    
    
    If .Movement = NPC_PATHFINDING Then
        Set .Inteligencia = New Inteligencia_Morgolock
        .Movement = NPC_MALO_ATACA_USUARIOS_BUENOS
    Else
        Set .Inteligencia = Inte_Basica
        
    End If
    
    'Cargamos los items que comercia
    .Invent.NroItems = val(Leer.DarValor("NPC" & NpcNumber, "NROITEMS")) '
    For loopC = 1 To .Invent.NroItems '
        ln = Leer.DarValor("NPC" & NpcNumber, "Obj" & loopC) '
        .Invent.Object(loopC).ObjIndex = val(ReadField(1, ln, 45)) '
        .Invent.Object(loopC).Amount = val(ReadField(2, ln, 45)) '
    Next loopC
    
    ' Cargo los objetos que dropean bajo circunstancia especiales.
    .Invent.NroItemsDrop = val(Leer.DarValor("NPC" & NpcNumber, "NROITEMSDROP"))
    For loopC = 1 To .Invent.NroItemsDrop
        ln = Leer.DarValor("NPC" & NpcNumber, "ObjDrop" & loopC)
        .Invent.ObjectDrop(loopC).ObjIndex = val(ReadField(1, ln, 45)) '
        .Invent.ObjectDrop(loopC).Amount = val(ReadField(2, ln, 45)) '
        .Invent.ObjectDrop(loopC).Probability = val(ReadField(3, ln, 45)) '
    Next loopC
    
    'Los hechizos que lanza
    .NroSpells = val(Leer.DarValor("NPC" & NpcNumber, "LanzaSpells")) '
    If .NroSpells > 0 Then ReDim .Spells(1 To .NroSpells) '
    For loopC = 1 To .NroSpells '
        .Spells(loopC) = val(Leer.DarValor("NPC" & NpcNumber, "Sp" & loopC)) '
    Next loopC '
    
    'Las criaturas que invoca
    If .NPCtype = NPCTYPE_ENTRENADOR Then
        .NroCriaturas = val(Leer.DarValor("NPC" & NpcNumber, "NroCriaturas"))
        ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
        For loopC = 1 To .NroCriaturas
            .Criaturas(loopC).npcIndex = Leer.DarValor("NPC" & NpcNumber, "CI" & loopC)
            .Criaturas(loopC).NpcName = Leer.DarValor("NPC" & NpcNumber, "CN" & loopC)
        Next loopC
    End If
    
    .Inflacion = val(Leer.DarValor("NPC" & NpcNumber, "Inflacion")) '
    
    If Respawn Then
        .flags.Respawn = val(Leer.DarValor("NPC" & NpcNumber, "ReSpawn")) '
    Else
        .flags.Respawn = 1 '
    End If
    
    .flags.BackUp = val(Leer.DarValor("NPC" & NpcNumber, "BackUp")) '
    .flags.AfectaParalisis = val(Leer.DarValor("NPC" & NpcNumber, "AfectaParalisis")) '
    .flags.Snd1 = val(Leer.DarValor("NPC" & NpcNumber, "Snd1")) '
    .flags.Snd2 = val(Leer.DarValor("NPC" & NpcNumber, "Snd2")) '
    .flags.Snd3 = val(Leer.DarValor("NPC" & NpcNumber, "Snd3")) '
    
    'Tipo de items con los que comercia
    .TipoItems = val(Leer.DarValor("NPC" & NpcNumber, "TipoItems")) '
    
    .Nivel = val(Leer.DarValor("NPC" & NpcNumber, "NIVEL"))
End With

NumNPCs = NumNPCs + 1
'Devuelve el nuevo Indice
OpenNPC = npcIndex

End Function

Public Sub guardarCriaturas()

Dim NpcNumber As Integer
Dim npcIndex As Integer
Dim archivo As cIniManager
Dim npcfile As String
Dim loopC As Integer

npcfile = App.Path & "/" & "criaturas.dat"

Set archivo = New cIniManager
Call archivo.Initialize(npcfile)

For NpcNumber = 1 To 618

    npcIndex = OpenNPC(NpcNumber)
    
    With NpcList(npcIndex)
          
        'Generales
        Call archivo.ChangeValue(.numero, "Name", .Name)
        Call archivo.ChangeValue(.numero, "DescInterna", "")
        Call archivo.ChangeValue(.numero, "Desc", .desc)
        Call archivo.ChangeValue(.numero, "Tipo", .NPCtype)
        
        Call archivo.ChangeValue(.numero, "MaxHP", .Stats.MaxHP)
        Call archivo.ChangeValue(.numero, "MinHP", .Stats.minHP)
        Call archivo.ChangeValue(.numero, "Renace", IIf(.flags.Respawn = 0, 1, 0))
        Call archivo.ChangeValue(.numero, "Backup", .flags.BackUp)
        Call archivo.ChangeValue(.numero, "Movement", .Movement)
        Call archivo.ChangeValue(.numero, "PoderAtaque", .PoderAtaque)
        Call archivo.ChangeValue(.numero, "PoderEvasion", .PoderEvasion)

        '   Donde puede nacer
        Select Case .flags.Terreno
            
            Case eTerrenoNPC.Tierra
                Call archivo.ChangeValue(.numero, "Terreno", "0")
            Case eTerrenoNPC.Agua
                Call archivo.ChangeValue(.numero, "Terreno", "1")
            Case eTerrenoNPC.AguayTierra
                Call archivo.ChangeValue(.numero, "Terreno", "0-1")
        End Select
    
        'Faccion
        Select Case .faccion
            Case eAlineaciones.Neutro
                Call archivo.ChangeValue(.numero, "Faccion", 0)
            Case eAlineaciones.Real
                Call archivo.ChangeValue(.numero, "Faccion", 2)
            Case eAlineaciones.caos
                Call archivo.ChangeValue(.numero, "Faccion", 4)
        End Select
        
        'Combate
        Call archivo.ChangeValue(.numero, "MaxHIT", .Stats.MaxHIT)
        Call archivo.ChangeValue(.numero, "MinHIT", .Stats.MinHIT)
        Call archivo.ChangeValue(.numero, "DEF", .Stats.Def)
        
        '       Hechizo
            For loopC = 1 To 5
                If .NroSpells >= loopC Then
                    Call archivo.ChangeValue(.numero, "HECHIZO" & loopC, .Spells(loopC))
                Else
                    Call archivo.ChangeValue(.numero, "HECHIZO" & loopC, 0)
                End If
            Next loopC
            
        'Audio
        Call archivo.ChangeValue(.numero, "Snd1", .flags.Snd1)
        Call archivo.ChangeValue(.numero, "Snd2", .flags.Snd2)
        Call archivo.ChangeValue(.numero, "Snd3", .flags.Snd3)
    
        'Visible
        Call archivo.ChangeValue(.numero, "Body", .Char.Body)
        Call archivo.ChangeValue(.numero, "Head", .Char.Head)
        Call archivo.ChangeValue(.numero, "Casco", .Char.CascoAnim)
        Call archivo.ChangeValue(.numero, "Escudo", .Char.ShieldAnim)
        Call archivo.ChangeValue(.numero, "Arma", .Char.WeaponAnim)
        Call archivo.ChangeValue(.numero, "Heading", .Char.heading)
        
        'Comportamiento
        Call archivo.ChangeValue(.numero, "Attackable", .Attackable)
        Call archivo.ChangeValue(.numero, "InmuneAHechizos", .InmuneAHechizos)
        Call archivo.ChangeValue(.numero, "AfectaParalisis", IIf(.flags.AfectaParalisis = 1, 0, 1))
            
        Call archivo.ChangeValue(.numero, "Veneno", .Veneno)
        Call archivo.ChangeValue(.numero, "Domable", .flags.Domable)
        
        'Comercio
        Call archivo.ChangeValue(.numero, "Comercia", .Comercia)
        Call archivo.ChangeValue(.numero, "TipoItems", .TipoItems)
        Call archivo.ChangeValue(.numero, "Inflacion", .Inflacion)
        Call archivo.ChangeValue(.numero, "InvReSpawn", .InvReSpawn)
        
    
        '   Items que comercia
        For loopC = 1 To 20

            If .Invent.NroItems >= loopC Then
                Call archivo.ChangeValue(.numero, "OBJ" & loopC, .Invent.Object(loopC).ObjIndex & "-" & .Invent.Object(loopC).Amount)
            Else
                Call archivo.ChangeValue(.numero, "OBJ" & loopC, "0-0")
            End If
        Next loopC
    
        'Retribuciones
        Call archivo.ChangeValue(.numero, "GiveGLD", .GiveGLD)
        Call archivo.ChangeValue(.numero, "GiveEXP", .GiveEXP)
        
    End With
Next

Call archivo.DumpFile(npcfile)

End
End Sub

Sub EnviarListaCriaturas(ByVal UserIndex As Integer, ByVal npcIndex)
  Dim SD As String
  Dim k As Integer
  SD = SD & NpcList(npcIndex).NroCriaturas & ","
  
  'Supongo que esto esta para evitar MACROS
  If Int(RandomNumber(0, 3)) = 1 Then
    For k = 1 To NpcList(npcIndex).NroCriaturas
        SD = SD & NpcList(npcIndex).Criaturas(k).NpcName & ","
    Next k
  Else
    For k = NpcList(npcIndex).NroCriaturas To 1 Step -1
        SD = SD & NpcList(npcIndex).Criaturas(k).NpcName & ","
    Next k
  End If
  
  'Les envio el nombre de la criatura
  EnviarPaquete Paquetes.EnviarNpclst, SD, UserIndex, ToIndex
End Sub

Sub DoFollow(ByRef criatura As npc, UserIndex As Integer)

'Si lo esta siguiendo lo dejo de seguir
If criatura.MaestroUser > 0 Then
  criatura.MaestroUser = 0
  criatura.Movement = criatura.flags.OldMovement
Else 'Si no me esta siguiendo, hago que me siga
  criatura.MaestroUser = UserIndex
  criatura.Movement = SIGUE_AMO 'follow
End If
End Sub

Sub FollowAmo(ByVal npcIndex As Integer)
  NpcList(npcIndex).Movement = SIGUE_AMO 'follow
  NpcList(npcIndex).TargetUserID = 0
  NpcList(npcIndex).TargetNPCID = 0
End Sub

Public Sub iniciarEstructurasNpcs()
    Dim i As Integer

    Set npcIndexLibres = New EstructurasLib.ColaConBloques
    
    Call npcIndexLibres.setCantidadElementosNodo(500)
    
    'Agrego de atras para adelante para que primero tomes los indexs más chicos
    'Los ultimos serán los primeros.
    For i = MAXNPCS To 1 Step -1
        Call npcIndexLibres.agregar(i)
    Next

    'Inteligencia artificial
    Set Inte_Basica = New Inteligencia_Basica
End Sub


Public Sub procesarNpcs()

Dim npcIndex As Integer

Dim mapa As Integer

Static tiempo As Long

tiempoLlamadaNpcs = GetTickCount

If tiempo = 0 Then tiempo = tiempoLlamadaNpcs
tiempo = tiempoLlamadaNpcs - tiempo

If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    
    For mapa = 1 To NumMaps
        'Solo proceso los npcs de los mapas donde haya usuarios
        If MapInfo(mapa).Existe = False Then GoTo continue
            
        If MapInfo(mapa).usuarios.getCantidadElementos > 0 Then

            MapInfo(mapa).NPCs.itIniciarB
            
            Do While MapInfo(mapa).NPCs.ithasNextB
            
                npcIndex = MapInfo(mapa).NPCs.itnextB
                
                If NpcList(npcIndex).flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(npcIndex, tiempo)
                Else
                        'Usamos AI si hay algun user en el mapa
                        If NpcList(npcIndex).flags.Inmovilizado = 1 Then
                            Call EfectoParalisisNpc(npcIndex, tiempo)
                        End If
                        
                        If NpcList(npcIndex).Movement <> ESTATICO Then
                            Call NPCAI(npcIndex)
                        End If
                End If
            Loop
        End If
continue:
    Next mapa

End If

tiempo = GetTickCount

End Sub

Public Sub eliminarTodasLasMascotas()

Dim mapa As Integer
Dim npcIndex As Integer

For mapa = 1 To NumMaps

    If MapInfo(mapa).Existe Then
        If MapInfo(mapa).NPCs.getCantidadElementos > 0 Then
            
            With MapInfo(mapa).NPCs
                .itIniciar
                
                Do While .ithasNext
                
                    npcIndex = .itnext
                    
                    If NpcList(npcIndex).Contadores.TiempoExistencia > 0 Then
                        Call MuereNpc(NpcList(npcIndex))
                    End If
                Loop
            End With
        End If
    End If
Next mapa

End Sub

'Pone a los guardias en su posición original
Sub ReSpawnOrigPosNpcs()

Dim mapa As Integer
Dim npcIndex As Integer

For mapa = 1 To NumMaps

    If MapInfo(mapa).Existe Then
        If MapInfo(mapa).NPCs.getCantidadElementos > 0 Then
            
            With MapInfo(mapa).NPCs
                .itIniciar
                
                Do While .ithasNext
                
                    npcIndex = .itnext
                        
                    If SV_PosicionesValidas.existePosicionMundo(NpcList(npcIndex).Orig.map, NpcList(npcIndex).Orig.x, NpcList(npcIndex).Orig.y) And NpcList(npcIndex).numero = Guardias Then
                        Call ReSpawnNpc(NpcList(npcIndex))
                        Call QuitarNPC(npcIndex)
                    End If
    
                Loop
            End With
            
        End If
    End If
Next mapa

End Sub

Public Sub getEstadisticas(ByRef formulario As Form)

Dim mapa As Integer
Dim activos As Integer

activos = 0

For mapa = 1 To NumMaps

    If MapInfo(mapa).Existe Then
        activos = activos + MapInfo(mapa).NPCs.getCantidadElementos
        
        MapInfo(mapa).NPCs.itIniciar
    End If
        
Next mapa

formulario.Label1.Caption = "Npcs Activos:" & activos & "( " & NumNPCs & " )"
formulario.Label2.Caption = "Npcs Libres:" & npcIndexLibres.getCantidadElementos
formulario.Label4.Caption = "MAXNPCS:" & MAXNPCS

End Sub

Public Sub ponerEstatico(npcIndex As Integer)

    NpcList(npcIndex).Movement = ESTATICO
    
    'Call MapInfo(Npclist(NpcIndex).Pos.Map).NPCs.eliminar(NpcIndex)
    
End Sub

Public Sub quitarEstatico(npcIndex As Integer)
   
    'Call MapInfo(Npclist(NpcIndex).Pos.Map).NPCs.agregar(NpcIndex)
    
End Sub

Public Sub establecerAmo(UserIndex As Integer, npcIndex As Integer)
    Dim index As Integer

    'Obtengo una posicion libre de las mascotas
    index = FreeMascotaIndex(UserIndex)
    'Guardo los datos de la mascota
    UserList(UserIndex).MascotasIndex(index) = npcIndex
    UserList(UserIndex).MascotasType(index) = NpcList(npcIndex).numero
    'Relaciono a npc con su dueño
    NpcList(npcIndex).MaestroUser = UserIndex
    'Aumento la cantidad de mascotas que tiene
    UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
End Sub
