Attribute VB_Name = "Acciones"
Option Explicit

Private Enum eTarget
    eNinguno = 0
    eCriatura = 1
    ePersonaje = 2
    eObjeto = 3
End Enum

Private Const CANTIDAD_MAXIMA_FOGATAS = 20

Private Function obtenerTarget(ByVal mapa As Integer, ByVal x As Integer, ByVal y As Integer, _
                                ByRef outputTargetX As Integer, ByRef outputTargetY As Integer, ByRef outputTargetIndex As Integer) _
                                As eTarget

    Dim TargetX As Integer
    Dim TargetY As Integer

    ' X X X
    ' X 0 1
    ' X 3 2
    
    ' Chequea si la posicion es valida
    '   Si hay un objeto ejecuta la accion ene se tile
    ' Chequea si la posición derecha es un objeto
    ' Chequea si abajo hacia la derecha
    ' Chequea abajo
    ' Chequea Criatura
    '   Ejecuta la accion

    ' Iniciamos las variables
    TargetX = x
    TargetY = y

    ' ¿La posición es valida?
    If Not SV_PosicionesValidas.existePosicionMundo(mapa, TargetX, TargetY) Then
        obtenerTarget = eTarget.eNinguno
        outputTargetX = 0
        outputTargetY = 0
        outputTargetIndex = 0
        Exit Function
    End If

    ' ¿Hay una criatura?
    If MapData(mapa, TargetX, TargetY).npcIndex > 0 Then
        outputTargetIndex = MapData(mapa, TargetX, TargetY).npcIndex
        outputTargetX = TargetX
        outputTargetY = TargetY
        obtenerTarget = eTarget.eCriatura
        Exit Function
    End If
    
    ' ¿hay un objeto?
    If MapData(mapa, TargetX, TargetY).OBJInfo.ObjIndex > 0 Then
        outputTargetIndex = MapData(mapa, TargetX, TargetY).OBJInfo.ObjIndex
        outputTargetX = TargetX
        outputTargetY = TargetY
        obtenerTarget = eTarget.eObjeto
        Exit Function
    End If

    ' Zona Insegura, salimos
    If MapInfo(mapa).Pk = True Then
        obtenerTarget = eTarget.eNinguno
        outputTargetX = 0
        outputTargetY = 0
        outputTargetIndex = 0
        Exit Function
    End If

    ' Si es segura, probamos una ayuda para hacer clic en los personajes
    TargetY = y - 1

    ' Probamos en la pos superior
    If SV_PosicionesValidas.existePosicionMundo(mapa, TargetX, TargetY) Then
        If MapData(mapa, TargetX, TargetY).npcIndex > 0 Then
            outputTargetIndex = MapData(mapa, TargetX, TargetY).npcIndex
            outputTargetX = TargetX
            outputTargetY = TargetY
            obtenerTarget = eTarget.eCriatura
            Exit Function
        End If
    End If
    
    ' Probamos en la pos inferior
    TargetY = y + 1

    If SV_PosicionesValidas.existePosicionMundo(mapa, TargetX, TargetY) Then
        If MapData(mapa, TargetX, TargetY).npcIndex > 0 Then
            outputTargetIndex = MapData(mapa, TargetX, TargetY).npcIndex
            outputTargetX = TargetX
            outputTargetY = TargetY
            obtenerTarget = eTarget.eCriatura
            Exit Function
        End If
    End If

End Function

Public Sub accion(ByRef personaje As User, ByVal x As Integer, ByVal y As Integer)

Dim targetMapa As Integer
Dim TargetX As Integer
Dim TargetY As Integer

Dim targetTipo As eTarget
Dim targetIndex As Integer

' Inicializamos
TargetX = 0
TargetY = 0
targetIndex = 0
targetMapa = personaje.pos.map

' Obtenemos el target donde hizo clic
targetTipo = obtenerTarget(personaje.pos.map, x, y, TargetX, TargetY, targetIndex)

Select Case targetTipo

' *****************************************************************************
    Case eTarget.eNinguno
    
        ' Reseteamos los Targets
        personaje.flags.TargetNPC = 0
        personaje.flags.TargetNpcTipo = 0
        personaje.flags.TargetUser = 0
        personaje.flags.TargetObj = 0

' *****************************************************************************
    Case eTarget.eObjeto
    
        ' Seteamos el Objeto
        personaje.flags.TargetObj = targetIndex

        ' Ejecuto la acción dependiendo el objeto
        Select Case ObjData(targetIndex).ObjType
            Case OBJTYPE_PUERTAS 'Es una puerta
                Call AccionParaPuerta(targetMapa, TargetX, TargetY, personaje.UserIndex)
            Case OBJTYPE_CARTELES 'Es un cartel
                Call AccionParaCartel(targetMapa, TargetX, TargetY, personaje.UserIndex)
            Case OBJTYPE_LEÑA 'Leña
                If targetIndex = FOGATA_APAG Then
                    Call AccionParaRamita(targetMapa, TargetX, TargetY, personaje)
                End If
        End Select
' *****************************************************************************
    Case eTarget.eCriatura

        If NpcList(targetIndex).Comercia = 1 Then ' ¿La criatura es un comerciante?
    
            ' ¿Distancia aceptada?
            If distancia(NpcList(targetIndex).pos, personaje.pos) > 3 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(7), personaje.UserIndex
                Exit Sub
            End If
        
            personaje.flags.TargetNPC = targetIndex
        
            'Iniciamos la rutina pa' comerciar.
            Call IniciarCOmercioNPC(personaje.UserIndex)
        
        ElseIf NpcList(targetIndex).NPCtype = NPCTYPE_BANQUERO Then 'NPC Banquero

            '¿Esta el user muerto? Si es asi no puede comerciar
            If personaje.flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(3), personaje.UserIndex
                Exit Sub
            End If
            
            '¿El target es un NPC valido?
            If distancia(NpcList(targetIndex).pos, personaje.pos) > 3 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(7), personaje.UserIndex
                Exit Sub
            End If
            
            personaje.flags.TargetNPC = targetIndex
            personaje.flags.TargetNpcTipo = NPCTYPE_BANQUERO
            
            ' Iniciamos el deposito
            Call IniciarDeposito(personaje.UserIndex)
    
        ElseIf NpcList(targetIndex).NPCtype = NPCTYPE_REVIVIR Then 'NPC Sacerdote
    
            If personaje.flags.Muerto = 1 Then
            
                ' El usuario está muerto entonces lo resucito
                If distancia(personaje.pos, NpcList(targetIndex).pos) > 10 Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(8), personaje.UserIndex
                    Exit Sub
                End If
                
                ' Revivo al usuario
                Call RevivirUsuario(personaje, 1, 50, 50)
                
                ' Avisamos
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(41), personaje.UserIndex
                
                Exit Sub
            Else ' Lo curo
                
                ' ¿Distancia valida?
                If distancia(personaje.pos, NpcList(targetIndex).pos) > 10 Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(8), personaje.UserIndex
                    Exit Sub
                End If
                
                ' No tiene vida para curarse?
                If (personaje.Stats.MaxHP <= personaje.Stats.minHP) Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(30), personaje.UserIndex
                    Exit Sub
                End If
                
                ' Lo curo
                personaje.Stats.minHP = personaje.Stats.MaxHP
                
                ' Envio
                Call SendUserStatsBox(personaje.UserIndex)
                
                ' Aviso
                EnviarPaquete Paquetes.MensajeSimple, Chr$(17), personaje.UserIndex
                Exit Sub
            End If
        End If

End Select

End Sub


Private Sub agregarFogata(ByVal mapa As Integer, ByVal x As Integer, ByVal y As Integer)

Dim obj As obj
Dim fogataData As ItemMapaData

obj.ObjIndex = FOGATA
obj.Amount = 1

Call MakeObj(ToMap, 0, mapa, obj, mapa, x, y)

' Guardamos la creación de esta fogata en el mapa
Set fogataData = New ItemMapaData

fogataData.x = x
fogataData.y = y
fogataData.index = FOGATA
fogataData.fecha = GetTickCount

' Agregamos a la info del mapa
Call MapInfo(mapa).fogatas.Add(fogataData)

If CANTIDAD_MAXIMA_FOGATAS > 0 And MapInfo(mapa).fogatas.Count > CANTIDAD_MAXIMA_FOGATAS Then
    ' Voy a buscar la mas vieja y la elimino
    Set fogataData = MapInfo(mapa).fogatas.Item(1)
    
    Call MapInfo(mapa).fogatas.Remove(1)
        
    Call EraseObj(ToMap, 0, mapa, 1, mapa, fogataData.x, fogataData.y)
End If



End Sub
'---------------------------------------------------------------------------------------
' Procedure : AccionParaRamita
' DateTime  : 18/02/2007 18:55
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub AccionParaRamita(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByRef personaje As User)

Dim Suerte As Byte
Dim exito As Byte
Dim supervivenciaSkills As Byte
Dim pos As WorldPos

pos.map = map
pos.x = x
pos.y = y

'TODO ESTO ES FEO, SACAR, CHEQUEA SI ES ESPERANZA
If map = 112 Then Exit Sub

If distancia(pos, personaje.pos) > 2 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(5), personaje.UserIndex
    Exit Sub
End If

If MapInfo(personaje.pos.map).Pk = False Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(86), personaje.UserIndex
    Exit Sub
End If

supervivenciaSkills = personaje.Stats.UserSkills(Supervivencia)

If supervivenciaSkills > 1 And supervivenciaSkills < 6 Then
    Suerte = 3
ElseIf supervivenciaSkills >= 6 And supervivenciaSkills <= 10 Then
    Suerte = 2
ElseIf supervivenciaSkills >= 10 And supervivenciaSkills Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

'Sino tiene hambre o sed quizas suba el skill supervivencia
If personaje.flags.Hambre = 0 And personaje.flags.Sed = 0 Then
    Call SubirSkill(personaje.UserIndex, Supervivencia)
End If

If Not exito = 1 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(29), personaje.UserIndex
    Exit Sub
End If

' Creamos la fogata
Call agregarFogata(map, x, y)

EnviarPaquete Paquetes.PrenderFogata, "", personaje.UserIndex, 0, ToPCArea, map

EnviarPaquete Paquetes.MensajeSimple, Chr$(28), personaje.UserIndex

End Sub

Sub AccionParaPuerta(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)


If Not (Distance(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, x, y) > 2) Then
    If ObjData(MapData(map, x, y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(map, x, y).OBJInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(map, x, y).OBJInfo.ObjIndex).Llave = 0 Then
                     MapData(map, x, y).OBJInfo.ObjIndex = ObjData(MapData(map, x, y).OBJInfo.ObjIndex).IndexAbierta
                     Call MakeObj(ToMap, 0, map, MapData(map, x, y).OBJInfo, map, x, y)
                     'Bloquea todos los mapas
                     Call Bloquear(ToMap, 0, map, map, x, y, 0)
                     Call Bloquear(ToMap, 0, map, map, x - 1, y, 0)
                     'Sonido
                     EnviarPaquete Paquetes.WavSnd, Chr$(SND_PUERTA), UserIndex, ToPCArea, UserList(UserIndex).pos.map
                Else
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(20), UserIndex

                End If
        Else
                'Cierra puerta
                MapData(map, x, y).OBJInfo.ObjIndex = ObjData(MapData(map, x, y).OBJInfo.ObjIndex).IndexCerrada
                Call MakeObj(ToMap, 0, map, MapData(map, x, y).OBJInfo, map, x, y)
                Call Bloquear(ToMap, 0, map, map, x - 1, y, 1)
                Call Bloquear(ToMap, 0, map, map, x, y, 1)
                EnviarPaquete Paquetes.WavSnd, Chr$(SND_PUERTA), UserIndex, ToPCArea, UserList(UserIndex).pos.map
        End If
        UserList(UserIndex).flags.TargetObj = MapData(map, x, y).OBJInfo.ObjIndex
    Else
        EnviarPaquete Paquetes.MensajeSimple, Chr$(20), UserIndex
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple, Chr$(5), UserIndex
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : AccionParaCartel
' DateTime  : 18/02/2007 18:54
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub AccionParaCartel(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)

If ObjData(MapData(map, x, y).OBJInfo.ObjIndex).ObjType = 8 Then
  If Len(ObjData(MapData(map, x, y).OBJInfo.ObjIndex).texto) > 0 Then
  EnviarPaquete Paquetes.MostrarCartel, ObjData(MapData(map, x, y).OBJInfo.ObjIndex).texto & "Ç" & ObjData(MapData(map, x, y).OBJInfo.ObjIndex).GrhSecundario, UserIndex
  End If
End If

End Sub

Public Sub AccionLanzarProyectil(ByRef personaje As User, ByVal x As Integer, ByVal y As Integer)
    Dim dummyint As Integer
    Dim tieneEnergia As Boolean
    Dim TU As Integer   ' Target User
    Dim tN As Integer   ' Target Criatura
    Dim wp2 As WorldPos
    
    If personaje.flags.Meditando = True Then Exit Sub
    If personaje.Counters.combateRegresiva > 0 Then Exit Sub
    
    If Not personaje.resucitacionPendiente Is Nothing Then
        Call modResucitar.cancelarResucitacion(personaje.resucitacionPendiente)
    End If
    
    'Nos aseguramos que este usando un arma de proyectiles
    dummyint = 0
    If personaje.Invent.WeaponEqpObjIndex = 0 Then
        dummyint = 1
    ElseIf personaje.Invent.WeaponEqpSlot < 1 Or personaje.Invent.WeaponEqpSlot > personaje.Stats.MaxItems Then
        dummyint = 1
    ElseIf personaje.Invent.MunicionEqpSlot < 1 Or personaje.Invent.MunicionEqpSlot > personaje.Stats.MaxItems Then
        dummyint = 1
    ElseIf personaje.Invent.MunicionEqpObjIndex = 0 Then
        dummyint = 1
    ElseIf ObjData(personaje.Invent.WeaponEqpObjIndex).proyectil <> 1 Then
        dummyint = 2
    ElseIf ObjData(personaje.Invent.MunicionEqpObjIndex).ObjType <> OBJTYPE_FLECHAS Then
        dummyint = 1
    ElseIf personaje.Invent.Object(personaje.Invent.MunicionEqpSlot).Amount < 1 Then
        dummyint = 1
    End If
                    
    If dummyint <> 0 Then
        If dummyint = 1 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr(230), personaje.UserIndex
        End If
            
        Call Desequipar(personaje.UserIndex, personaje.Invent.MunicionEqpSlot)
        Call Desequipar(personaje.UserIndex, personaje.Invent.WeaponEqpSlot)
            
        Exit Sub
    End If
                    
    dummyint = 0
            
    'Quitamos stamina
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, ObjData(personaje.Invent.WeaponEqpObjIndex).QuitaEnergia)
      
    If Not tieneEnergia Then
        EnviarPaquete Paquetes.MensajeSimple, Chr(11), personaje.UserIndex
        Exit Sub
    End If
            
    ' Obtenemos nuestro Objetivo
    Call LookatTileII(personaje.UserIndex, personaje.pos.map, x, y)
            
    TU = personaje.flags.TargetUser
    tN = personaje.flags.TargetNPC
                        
    If tN > 0 Or TU > 0 Then
                
        wp2.map = personaje.pos.map
        wp2.x = x
        wp2.y = y
        
        'Distancia de ataque. Poner las constantes correctas
        If Abs(personaje.pos.x - wp2.x) <= modHechizos.MAX_DISTANCIA_LANZA_HECHIZOS_ANCHO And Abs(personaje.pos.y - wp2.y) <= modHechizos.MAX_DISTANCIA_LANZA_HECHIZOS_ALTO Then
                    
            ' ¿Criatura seleccionada?
            If tN > 0 Then
                
                ' ¿Es tacable?
                If NpcList(tN).Attackable <> 0 Then
                    If (personaje.flags.Oculto = 0 And personaje.flags.Invisible = 0) Then
                        EnviarPaquete Paquetes.FXh, ITS(personaje.Char.charIndex) & ITS(NpcList(tN).Char.charIndex) & Chr(2), personaje.UserIndex, ToPCArea, personaje.pos.map
                    End If
                        
                    Call UsuarioAtacaNpc(personaje.UserIndex, tN)
                End If
                        
            ElseIf TU > 0 Then ' ¿Usuario seleccionado?
                
                If TU = personaje.UserIndex Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr(231), personaje.UserIndex
                    dummyint = 1
                    Exit Sub
                End If
                                            
                If (personaje.flags.Oculto = 0 And personaje.flags.Invisible = 0) And UserList(TU).flags.Privilegios = 0 And puedeAtacar(personaje, UserList(TU)) Then
                    EnviarPaquete Paquetes.FXh, ITS(personaje.Char.charIndex) & ITS(UserList(TU).Char.charIndex) & Chr(2), personaje.UserIndex, ToPCArea, personaje.pos.map
                End If
                        
                Call UsuarioAtacaUsuario(personaje.UserIndex, TU)
            End If
                    
        Else 'Como hizo para pegarle si no lo ve?. Le actualizamos la posicion
            Call enviarPosicion(personaje)
        End If
    End If
           
    ' No importa si le pego a alguien o no
    If dummyint = 0 Then
    
        'Saca 1 flecha
        ' TODO. Chequear esto
        dummyint = personaje.Invent.MunicionEqpSlot
        Call QuitarUserInvItem(personaje.UserIndex, personaje.Invent.MunicionEqpSlot, 1)
        
        If dummyint < 1 Or dummyint > personaje.Stats.MaxItems Then Exit Sub
        
        ' ¿Le quedaron flechas?. TODO Esto seria redundante con el QuitarUserInvItem
        If personaje.Invent.Object(dummyint).Amount > 0 Then
            personaje.Invent.MunicionEqpSlot = dummyint
            EnviarPaquete Paquetes.ActualizaCantidadItem, Chr$(dummyint) & Codify(personaje.Invent.Object(personaje.Invent.MunicionEqpSlot).Amount), personaje.UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.ActualizaCantidadItem, Chr$(dummyint) & Codify(0), personaje.UserIndex, ToIndex
            personaje.Invent.MunicionEqpSlot = 0
            personaje.Invent.MunicionEqpObjIndex = 0
        End If
    End If
    
End Sub
Public Sub AccionConSkill(ByRef personaje As User, ByRef anexo As String)

    Dim tipoSkill As Byte
    Dim x As Integer
    Dim y As Integer
    Dim timeStamp As Single
    Dim hechizo As Byte
    
    If personaje.flags.Muerto = 1 Then Exit Sub

    x = STI(anexo, 1)
    y = STI(anexo, 3)
    tipoSkill = Asc(mid$(anexo, 5, 1))

    If tipoSkill = eSkills.Magia Then
        hechizo = Asc(mid$(anexo, 6, 1))
        timeStamp = StringToSingle(anexo, 7)
    Else
        hechizo = 0
        timeStamp = StringToSingle(anexo, 6)
    End If
    
    personaje.controlCheat.VecesAtack = personaje.controlCheat.VecesAtack + 1
        
    Select Case tipoSkill
        Case eSkills.proyectiles
            
            ' Anticheat
            Call anticheat.chequeoIntervaloCliente(personaje, personaje.Counters.ultimoTickProyectiles, personaje.intervalos.Flecha, timeStamp, "lanzar flechas")
            
            ' Lanzamos
            Call AccionLanzarProyectil(personaje, x, y)
            
        Case eSkills.Magia

            ' Anticheat
            Call anticheat.chequeoIntervaloCliente(personaje, personaje.Counters.ultimoTickMagia, personaje.intervalos.Magia, timeStamp, "lanzar hechizos")

            ' Lanzamos
            Call AccionLanzarHechizo(personaje, x, y, hechizo)
                
        Case eSkills.Robar
                
            Call modRobar.Robar(personaje, x, y)

        Case eSkills.Domar

            Call AccionDoma(personaje, x, y)
                
        Case eSkills.Herreria
                
            Call AccionHerreria(personaje, x, y)
    End Select
    
End Sub

Private Sub AccionLanzarHechizo(ByRef personaje As User, x As Integer, y As Integer, ByVal nMagia As Integer)
    Dim wp2 As WorldPos
    
    'Si estameditando el usuariro no puede lanzar
    If personaje.flags.Meditando = True Then Exit Sub
    
    'Si esta en el contador no puede lanzar magias. TODO cambiarlo solo a magias malas
    If personaje.Counters.combateRegresiva > 0 Then Exit Sub
    
    If Not personaje.resucitacionPendiente Is Nothing Then
        Call modResucitar.cancelarResucitacion(personaje.resucitacionPendiente)
    End If
    
    ' Chequeo de intervalo
    If Not modNuevoTimer.IntervaloPermiteLanzarSpell(personaje, True) Then
        Exit Sub
    End If
    
    ' Busco a que apunta
    personaje.flags.hechizo = nMagia
        
    Call LookatTileII(personaje.UserIndex, personaje.pos.map, x, y)

    wp2.map = personaje.pos.map
    wp2.x = x
    wp2.y = y

    'Distancia de ataque
    If Abs(personaje.pos.x - wp2.x) <= modHechizos.MAX_DISTANCIA_LANZA_HECHIZOS_ANCHO And Abs(personaje.pos.y - wp2.y) <= modHechizos.MAX_DISTANCIA_LANZA_HECHIZOS_ALTO Then
        
        If personaje.flags.hechizo > 0 Then
            Call LanzarHechizo(personaje.flags.hechizo, personaje)
            personaje.flags.hechizo = 0
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr(233), personaje.UserIndex
        End If
            
    Else 'Como hizo para tirar mas lejos de lo que ve?.
        'Le actualizo la posicion por las dudas
        Call enviarPosicion(personaje)
    End If
                
End Sub
Private Sub AccionDoma(ByRef personaje As User, x As Integer, y As Integer)

    Dim CI As Integer
    Dim wpaux As WorldPos

    ' Buscamos Objetivo
    Call LookatTile(personaje.UserIndex, personaje.pos.map, x, y)
    
    CI = personaje.flags.TargetNPC

    '¿Selecciono algo?
    If CI > 0 Then
        ' ¿Es domable?
        If NpcList(CI).flags.Domable > 0 Then
            
            wpaux.map = personaje.pos.map
            wpaux.x = x
            wpaux.y = y
            
            ' ¿Distancia OK?
            If distancia(wpaux, NpcList(personaje.flags.TargetNPC).pos) > 2 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr(5), personaje.UserIndex, ToIndex
                Exit Sub
            End If
            
            ' ¿No esta peleando con nadie?
            If NpcList(CI).TargetUserID <> 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr(243), personaje.UserIndex, ToIndex
                Exit Sub
            End If
            
            ' Trato de domarlo
            Call DoDomar(personaje.UserIndex, CI)
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr$(244), personaje.UserIndex, ToIndex
            Exit Sub
        End If
    Else
        EnviarPaquete Paquetes.MensajeSimple, Chr$(245), personaje.UserIndex, ToIndex
        End If
    Exit Sub
              
End Sub

Private Sub AccionHerreria(ByRef personaje As User, ClickX As Integer, ClickY As Integer)
    ' Obtenemos si hay algun objeto
    Call LookatTile(personaje.UserIndex, personaje.pos.map, ClickX, ClickY)
    
    ' ¿Algo?
    If UserList(personaje.UserIndex).flags.TargetObj > 0 Then
        ' ¿Ok?
        If ObjData(personaje.flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
            Call EnivarArmasConstruibles(personaje.UserIndex)
            Call EnivarArmadurasConstruibles(personaje.UserIndex)
            EnviarPaquete Paquetes.ShowHerreriaForm, "", personaje.UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr$(248), personaje.UserIndex, ToIndex
        End If
    Else
        EnviarPaquete Paquetes.MensajeSimple, Chr$(248), personaje.UserIndex, ToIndex
    End If
End Sub

' El usuario desea tirar oro al suelo.
Public Sub TirarOroAlSuelo(personaje As User, cantidad As Long)

    Dim posicionesNecesarias As Integer
    
    posicionesNecesarias = Round((cantidad / MAX_INVENTORY_OBJS) + 0.5)
    
    ' ¿El mapa tiene control de objetos?
    Dim x As Integer, y As Integer, mapa As Integer, xMin As Integer, xMax As Integer, yMin As Integer, yMax As Integer
    Dim libres As Integer
    
    mapa = personaje.pos.map
    x = personaje.pos.x
    y = personaje.pos.y
    
    If posicionesNecesarias > 1 Or MapData(mapa, x, y).OBJInfo.ObjIndex <> iORO Or MapData(mapa, x, y).OBJInfo.Amount + cantidad > MAX_INVENTORY_OBJS Then
        
        xMin = maxl(SV_Constantes.X_MINIMO_JUGABLE, personaje.pos.x - 3)
        xMax = minl(SV_Constantes.X_MAXIMO_JUGABLE, personaje.pos.x + 3)
        yMin = maxl(SV_Constantes.Y_MINIMO_JUGABLE, personaje.pos.y - 3)
        yMax = minl(SV_Constantes.Y_MAXIMO_JUGABLE, personaje.pos.y + 3)
        
        libres = 0
        
        For x = xMin To xMax
            For y = yMin To yMax
                If (MapData(mapa, x, y).Trigger And eTriggers.TodosBordesBloqueados) = False And MapData(mapa, x, y).accion Is Nothing And MapData(mapa, x, y).OBJInfo.ObjIndex = 0 Then
                    libres = libres + 1
                End If
            Next y
        Next x
        
        If libres + posicionesNecesarias < 25 Then
            EnviarPaquete Paquetes.mensajeinfo, "Muevete un poco más. En esta área ya no hay más lugar para arrojar oro.", personaje.UserIndex, ToIndex
            Exit Sub
        End If
    End If
    
    Dim cantidadTirada As Long
    cantidadTirada = TirarOro(cantidad, personaje)
   
    ' Lo registramos
    If cantidad > 10000 Then
        Call modLogsPersonajes.LogOroArrojado(personaje, cantidadTirada, personaje.pos.map, personaje.pos.x, personaje.pos.y)
    End If
End Sub

