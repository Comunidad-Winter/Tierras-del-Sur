Attribute VB_Name = "IA_BASICA"
Option Explicit

Enum eTipoObjetivo
    NADA = 0
    Usuario = 1
    criatura = 2
End Enum

Enum eTipoAtaque
    NADA = 0
    Golpe = 1
    Magia = 2
    Flecha = 3
End Enum

Private Const DISTANCIA_MAXIMA = 8

Private Function estaEnRango(pos1 As WorldPos, pos2 As WorldPos) As Boolean

estaEnRango = Abs(pos1.x - pos2.x) <= DISTANCIA_MAXIMA And Abs(pos1.y - pos2.y) <= DISTANCIA_MAXIMA

End Function

Private Function estaPegado(pos1 As WorldPos, pos2 As WorldPos) As Boolean

estaPegado = (Abs(pos1.x - pos2.x) + Abs(pos1.y - pos2.y)) = 1

End Function

Private Function distancia(pos1 As WorldPos, pos2 As WorldPos) As Byte

distancia = Abs(pos1.x - pos2.x) + Abs(pos1.y - pos2.y)

End Function

Private Function puedeAtacar(criatura As npc) As Boolean

Dim ahora As Long

ahora = NPCs.tiempoLlamadaNpcs

If criatura.Contadores.TiempoUltimoAtaque + INTERVALO_ATAQUE < ahora Then
    puedeAtacar = True
    criatura.Contadores.TiempoUltimoAtaque = ahora
Else
    puedeAtacar = False
End If

End Function

Private Function estaEnfrentado(pos1 As WorldPos, pos2 As WorldPos, heading As Byte) As Boolean

If heading = eHeading.NORTH Then
    estaEnfrentado = (pos1.x = pos2.x) And (pos1.y = pos2.y + 1)
ElseIf heading = eHeading.SOUTH Then
    estaEnfrentado = (pos1.x = pos2.x) And (pos1.y + 1 = pos2.y)
ElseIf heading = eHeading.EAST Then
    estaEnfrentado = (pos1.x + 1 = pos2.x) And (pos1.y = pos2.y)
ElseIf heading = eHeading.WEST Then
    estaEnfrentado = (pos1.x = pos2.x + 1) And (pos1.y = pos2.y)
Else
    estaEnfrentado = False
End If


End Function

'---------------------------------------------------------------------------------------
' Procedure : inteligenciaBasica_Seguir_Amo
' Author    : Marce
' Date      : 16/01/2011
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub inteligenciaBasica_Seguir_Amo(ByRef criatura As npc, npcIndex As Integer)


Dim tHeading As Byte
Dim UserIndex As Integer

UserIndex = criatura.MaestroUser
    
If UserIndex > 0 Then
    'Lo sigo solo si esta vivo y no invisible
    If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Invisible = 0 And UserList(UserIndex).flags.Oculto = 0 Then
        'Lo sigo si esta a más de 3 tiles de distancia. Sino el npc se mantiene quieto
        If distancia(criatura.pos, UserList(UserIndex).pos) > 3 Then
            'Determino el movimiento y me muevo
            If Int(RandomNumber(1, 10)) > 2 Then
                Call MoveNPCChar(npcIndex, determinarMovimiento_A1(criatura, Usuario, UserIndex))
            Else
                Call MoveNPCChar(npcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
        End If
    End If
End If

End Sub
Public Sub inteligenciaBasica_Seguir_Agresor(ByRef criatura As npc, npcIndex As Integer)

Dim tipoAtaque As eTipoAtaque
Dim tipoobjetivo As eTipoObjetivo
Dim subTipoAtaque As Byte

Dim VictimaIndex As Integer

Dim atacar As Boolean

Dim tHeading As eHeading

atacar = False

tipoobjetivo = eTipoObjetivo.NADA

If criatura.TargetUserID > 0 Then
    'El personaje esta online?
    
    If Not (criatura.numero = ELEMENTALAGUA) Then
    
        VictimaIndex = IDIndex(NpcList(npcIndex).TargetUserID)
        
        ' ¿Esta Online?
        If VictimaIndex > 0 Then
            ' ¿Esta en el mismo mapa?
            If criatura.pos.map = UserList(VictimaIndex).pos.map Then
                'Cumple condiciones
                If UserList(VictimaIndex).flags.Muerto = 0 And UserList(VictimaIndex).flags.Oculto = 0 And UserList(VictimaIndex).flags.Invisible = 0 And UserList(VictimaIndex).flags.Mimetizado = 0 Then
                
                    If estaEnRango(criatura.pos, UserList(VictimaIndex).pos) Then
                        
                        'Si tiene un dueño y este dueño es armada o ciudadano no le puede pegar a una persona armada / ciudadana
                        If criatura.MaestroUser > 0 Then
                            If Not puedeAtacarFaccion(UserList(criatura.MaestroUser), UserList(VictimaIndex)) Then
                                EnviarPaquete Paquetes.MensajeSimple2, Chr$(126), NpcList(npcIndex).MaestroUser
                                Call RestoreOldMovement(npcIndex)
                            End If
                        End If
                        
                        tipoobjetivo = eTipoObjetivo.Usuario

                        atacar = True
                    End If
                    
                ElseIf UserList(VictimaIndex).flags.Muerto = 1 Then
                    'Si esta muerto, fue.
                    Call RestoreOldMovement(npcIndex)
                End If
            Else
                Call RestoreOldMovement(npcIndex)
            End If
        End If
    Else
        Call RestoreOldMovement(npcIndex)
    End If
ElseIf criatura.TargetNPCID > 0 Then

    VictimaIndex = criatura.TargetNPCID
    
    'Cumple condiciones 'Estos npcs no atacan a otros NPCS.
     If Not (criatura.numero = ELEMENTALFUEGO Or criatura.numero = ELEMENTALTIERRA Or criatura.numero = ESPIRITU_INDOMABLE Or criatura.numero = FUEGO_FACTUO) Then
        If NpcList(VictimaIndex).pos.map = criatura.pos.map Then
            atacar = True
            tipoobjetivo = eTipoObjetivo.criatura
        Else
            'La victima se fue. Lo dejo de buscar
            Call RestoreOldMovement(npcIndex)
        End If
    End If

End If

'Ataco?
If tipoobjetivo <> eTipoObjetivo.NADA Then

    If atacar Then
        If puedeAtacar(criatura) Then
            Call criatura.Inteligencia.determinarAtaque(npcIndex, VictimaIndex, tipoobjetivo, tipoAtaque, subTipoAtaque)
            'Realizo el ataque
            Call realizarAtaqueNPC(npcIndex, VictimaIndex, tipoAtaque, tipoobjetivo, subTipoAtaque)
        End If
    End If
    
    'Listo. Me muevo
    tHeading = criatura.Inteligencia.determinarMovimiento(npcIndex, tipoobjetivo, VictimaIndex)

    If tHeading <> eHeading.Ninguno Then
        If criatura.flags.Inmovilizado = 0 Then
        Call MoveNPCChar(npcIndex, tHeading)
    End If
End If
End If
                    


End Sub

Public Sub realizarAtaqueNPC(npcIndex As Integer, index As Integer, tipoAtaque As eTipoAtaque, tipoobjetivo As eTipoObjetivo, subTipoAtaque As Byte)
    'Decidio que hay que atacarlo?
    If Not tipoAtaque = eTipoAtaque.NADA Then
        If tipoobjetivo = eTipoObjetivo.Usuario Then
            If tipoAtaque = eTipoAtaque.Golpe Then
                Call NpcAtacaUser(npcIndex, index)
            ElseIf eTipoAtaque.Magia Then
                Call NpcLanzaSpellSobreUser(NpcList(npcIndex), UserList(index), hechizos(subTipoAtaque))
            End If
        Else
            If eTipoAtaque.Golpe Then
                Call NpcAtacaNpc(npcIndex, index)
            ElseIf eTipoAtaque.Magia Then
                Call NpcLanzaSpellSobreNpc(npcIndex, index, subTipoAtaque)
            End If
        End If
    End If
End Sub

Public Function determinarMovimiento_A1(criatura As npc, tipoobjetivo As eTipoObjetivo, IndexObjetivo As Integer) As eHeading
'Determino el movimiento
If tipoobjetivo = eTipoObjetivo.Usuario Then
    determinarMovimiento_A1 = FindDirection(criatura.pos, UserList(IndexObjetivo).pos)
ElseIf tipoobjetivo = eTipoObjetivo.criatura Then
    determinarMovimiento_A1 = FindDirection(criatura.pos, NpcList(IndexObjetivo).pos)
End If
End Function
Public Sub inteligenciaBasica(ByRef criatura As npc, npcIndex As Integer)

Dim index As Integer
Dim tipoobjetivo As eTipoObjetivo
Dim tipoAtaque As eTipoAtaque
Dim subTipoAtaque As Byte
Dim tHeading As eHeading

'***************** Determino el objetivo ***********************
'Call determinarObjetivo_A1(criatura, index, tipoobjetivo)
Call criatura.Inteligencia.determinarObjetivo(npcIndex, index, tipoobjetivo)
'***************************************************************
'Puedo atacar
If puedeAtacar(criatura) Then
'Hay algún objetivo para atacar?
    If Not tipoobjetivo = eTipoObjetivo.NADA Then
        '***************** Determino el ataque   ***********************
        'Determino como lo voy a atacar
        Call criatura.Inteligencia.determinarAtaque(npcIndex, index, tipoobjetivo, tipoAtaque, subTipoAtaque)
        'call determinarAtaque_A1(criatura, index, tipoobjetivo, tipoAtaque, subTipoAtaque)
        'Realizo el ataque
        Call realizarAtaqueNPC(npcIndex, index, tipoAtaque, tipoobjetivo, subTipoAtaque)
        '***************************************************************
    End If
End If

'tHeading = determinarMovimiento_A1(criatura, tipoobjetivo, index)
tHeading = criatura.Inteligencia.determinarMovimiento(npcIndex, tipoobjetivo, index)


If tHeading <> eHeading.Ninguno Then
    If criatura.flags.Inmovilizado = 0 Then
        Call MoveNPCChar(npcIndex, tHeading)
   ' Else
    '    Call ChangeHeadingNpc(NpcIndex, tHeading)
    End If
End If
End Sub



Public Sub determinarAtaque_A1(ByRef criatura As npc, indexAtaque As Integer, ByRef tipoobjetivo As eTipoObjetivo, ByRef tipoAtaque As eTipoAtaque, ByRef subTipo As Byte)
'Si esta al costado le pego

tipoAtaque = eTipoAtaque.NADA
subTipo = 0

'Primero le intento pegar


If criatura.flags.Inmovilizado = 0 Then
        'Solo le puedo pegar si lo tengo pegado a mi
        If tipoobjetivo = eTipoObjetivo.Usuario Then
            'El dragon solo pega a gente visible
            If criatura.NroSpells = 0 Or (criatura.NroSpells > 0 And (UserList(indexAtaque).flags.Invisible = 1 Or UserList(indexAtaque).flags.Oculto = 1)) Then   'Esto es feo
                If estaPegado(criatura.pos, UserList(indexAtaque).pos) Then
                    tipoAtaque = eTipoAtaque.Golpe
                    Exit Sub
                End If
            End If
        Else
            'If estaPegado(criatura.Pos, Npclist(indexAtaque).Pos) Then
            If distancia(criatura.pos, NpcList(indexAtaque).pos) <= 3 Then
                tipoAtaque = eTipoAtaque.Golpe
                Exit Sub
            End If
        End If
Else
        'Solo le puedo pegar si esta enfrente mio
        If tipoobjetivo = eTipoObjetivo.Usuario Then
          If criatura.NroSpells = 0 Or (criatura.NroSpells > 0 And (UserList(indexAtaque).flags.Invisible = 1 Or UserList(indexAtaque).flags.Oculto = 1)) Then   'Esto es feo
                If estaEnfrentado(criatura.pos, UserList(indexAtaque).pos, criatura.Char.heading) Then
                    tipoAtaque = eTipoAtaque.Golpe
                    Exit Sub
                End If
            End If
        Else
            If distancia(criatura.pos, NpcList(indexAtaque).pos) <= 3 Then
            'If estaEnfrentado(criatura.Pos, Npclist(indexAtaque).Pos, criatura.Char.heading) Then
                tipoAtaque = eTipoAtaque.Golpe
                Exit Sub
            End If
        End If
End If

'Veo si le puedo tirar un spell
If criatura.NroSpells > 0 Then
     tipoAtaque = eTipoAtaque.Magia
     subTipo = criatura.Spells(RandomNumber(1, criatura.NroSpells))
     Exit Sub
End If

'Veo si le puedo pegar un flechazo

End Sub
Public Sub determinarObjetivo_A1(ByRef criatura As npc, ByRef index As Integer, ByRef tipoobjetivo As eTipoObjetivo)

Dim headingloop As Byte 'Lado donde esta mirando
Dim UserIndex As Integer
Dim npcIndex As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer
Dim nPos As WorldPos
Dim x As Integer
Dim y As Integer

Dim usuarios As EstructurasLib.ColaConBloques
Dim distanciaMinima As Byte
Dim distanciaActual As Byte

index = 0
tipoobjetivo = eTipoObjetivo.NADA


' PRIMERA OPCION
'Primero se fija si hay un personaje en los 4 lados donde esta
'Si no esta paralizado se fija en los 4 costados, y en todo el rango

If criatura.flags.Inmovilizado = 1 Then

    'Si esta paralizado se fija en la linea donde esta mirando, empezando por el más cercano
    Select Case criatura.Char.heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    distanciaActual = 1
    
    For y = criatura.pos.y + SignoNS To criatura.pos.y + SignoNS * 10 Step IIf(SignoNS = 0, 1, SignoNS)
        For x = criatura.pos.x + SignoEO To criatura.pos.x + SignoEO * 10 Step IIf(SignoEO = 0, 1, SignoEO)
            
            'La pos es valida?
            If SV_PosicionesValidas.existePosicionMundo(criatura.pos.map, x, y) Then
                UserIndex = MapData(criatura.pos.map, x, y).UserIndex 'Hay un usuario?
                'Es valido para determinar como objetivo?
                'No es GM
                'No esta muerto
                'No esta paralizado
                If UserIndex > 0 Then
                    If ((UserList(UserIndex).flags.Invisible = 0 And UserList(UserIndex).flags.Oculto = 0) Or distanciaActual = 1) And UserList(UserIndex).flags.Privilegios = 0 And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Mimetizado = 0 Then
                        'Cumple con las condiciones de alineacion
                        If (Not criatura.faccion = UserList(UserIndex).faccion.alineacion) Or (criatura.faccion = Neutro) Then
                            index = UserIndex
                            tipoobjetivo = eTipoObjetivo.Usuario
                            Exit Sub ' No hay nada mas que hacer
                        End If
                    Else 'No no es válido.
                        'Veo si hay un npc
                        npcIndex = MapData(criatura.pos.map, x, y).npcIndex
                        
                        'Hay un npc?
                        If npcIndex > 0 Then
                            'Me fijo si cumple con las condiciones
                            'Sea mascota y no este paralizado
                            If NpcList(npcIndex).MaestroUser > 0 And NpcList(npcIndex).flags.Paralizado = 0 Then
                                    'Ok, es buen candidato
                                    index = npcIndex
                                    tipoobjetivo = eTipoObjetivo.criatura
                                    Exit Sub 'no hay nada mas que hacer
                            End If
                        End If
                    End If
                End If
            End If
            distanciaActual = distanciaActual + 1
        Next x
    Next y
Else
        'El npc no esta paralizado
        'Me fijo en los costados
        For headingloop = eHeading.NORTH To eHeading.WEST
           nPos = criatura.pos
           
           Call HeadtoPos(headingloop, nPos)
                
            If SV_PosicionesValidas.existePosicionMundo(nPos.map, nPos.x, nPos.y) Then
                    UserIndex = MapData(nPos.map, nPos.x, nPos.y).UserIndex 'Hay un usuario?
                    
                    If UserIndex > 0 Then
                         If UserList(UserIndex).flags.Privilegios = 0 And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Mimetizado = 0 Then
                            If (Not criatura.faccion = UserList(UserIndex).faccion.alineacion) Or (criatura.faccion = Neutro) Then
                                index = UserIndex
                                tipoobjetivo = eTipoObjetivo.Usuario
                                Exit Sub
                            End If
                        End If
                    End If
            End If
        Next headingloop
        
        'Me fijo en el rango. El que este más cerca
       
        distanciaMinima = 255
        
        Set usuarios = MapInfo(criatura.pos.map).usuarios

        usuarios.itIniciar
        
        Do While usuarios.ithasNext
            UserIndex = usuarios.itnext
            
            'Cumple con los requisitos
            If UserList(UserIndex).flags.Invisible = 0 And UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.Privilegios = 0 And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Mimetizado = 0 Then
           'If UserList(UserIndex).flags.Invisible = 0 And UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Mimetizado = 0 Then
            'Esta en el rango de vision del npc?
            If estaEnRango(criatura.pos, UserList(UserIndex).pos) Then
                If (Not criatura.faccion = UserList(UserIndex).faccion.alineacion) Or (criatura.faccion = Neutro) Then
                        distanciaActual = distancia(criatura.pos, UserList(UserIndex).pos)
                        'Esta mas cerca que el anterior?
                        If distanciaActual < distanciaMinima Then
                            index = UserIndex
                            tipoobjetivo = eTipoObjetivo.Usuario
                            distanciaMinima = distanciaActual
                        End If
                        Exit Sub
                    End If
                End If
            End If
        Loop
End If

End Sub


Public Sub moverAlAzar(ByRef criatura As npc, npcIndex As Integer)
       
        If Int(RandomNumber(1, 12)) = 3 Then
            Call MoveNPCChar(npcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
        End If
End Sub
