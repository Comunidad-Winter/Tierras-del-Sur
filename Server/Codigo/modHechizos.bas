Attribute VB_Name = "modHechizos"
Option Explicit

Public Const MAX_DISTANCIA_LANZA_HECHIZOS_ALTO = 11
Public Const MAX_DISTANCIA_LANZA_HECHIZOS_ANCHO = 11

'Hechizos
Public Enum eHechizos
    Resucitar = 11
    Provocar_Hambre = 12
    Terrible_Hambre = 13
    
    Invisibilidad = 14
        
    Llamado_naturaleza = 16
    Invocar_Zombies = 17
    Torpeza = 19
    
    Debilidad = 21
    
    Invocar_elemetanl_fuego = 26
    Invocoar_elemental_agua = 27
    Invocoar_elemental_tierra = 28
    Implorar_ayuda = 29
    
    Ayuda_espiritu_indomable = 33
    Mimetismo = 35
    
    Invocar_Mascotas = 39
End Enum

Public Enum eAccionHechizo
    Ninguno = 0
End Enum


Private Function getResistenciaMagica(ByRef personaje As User) As Single
    
Dim skillsResistenciaMagica As Byte

If Not (personaje.clase = eClases.Paladin And personaje.clase = eClases.Guerrero And personaje.clase = eClases.Cazador) Then
    ' Tiene que tener un anillo
    If personaje.Invent.AnilloEqpObjIndex = 0 Then
        getResistenciaMagica = 0
        Exit Function
    End If
End If


skillsResistenciaMagica = personaje.Stats.UserSkills(eSkills.ResistenciaMagica)

If skillsResistenciaMagica = 0 Then
    getResistenciaMagica = 0
ElseIf skillsResistenciaMagica < 31 Then
    getResistenciaMagica = 0.01
ElseIf skillsResistenciaMagica < 61 Then
    getResistenciaMagica = 0.02
ElseIf skillsResistenciaMagica < 91 Then
    getResistenciaMagica = 0.03
ElseIf skillsResistenciaMagica < 100 Then
    getResistenciaMagica = 0.04
Else
    getResistenciaMagica = 0.05
End If

End Function



Sub NpcLanzaSpellSobreUser(ByRef criatura As npc, ByRef personaje As User, ByRef hechizo As tHechizo)

'Este sub fue modificado para que al meditar los hechizos de daño te desconcentren y los otros no se vean.
Dim daño As Integer

If hechizo.SubeHP = 1 Then
    
    daño = RandomNumber(hechizo.minHP, hechizo.MaxHP)
    
    daño = daño - (daño * getResistenciaMagica(personaje))  'Resistencia magica
    
    If Not personaje.flags.Meditando Then
        EnviarPaquete Paquetes.HechizoFX, ITS(personaje.Char.charIndex) & ByteToString(hechizo.FXgrh) & ITS(hechizo.loops) & Chr$(hechizo.WAV), personaje.UserIndex, ToPCArea, personaje.pos.map
    End If
    
    personaje.Stats.minHP = personaje.Stats.minHP + daño
    
    If personaje.Stats.minHP > personaje.Stats.MaxHP Then personaje.Stats.minHP = personaje.Stats.MaxHP
    
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(9) & criatura.Name & "," & daño, personaje.UserIndex
    
    Call SendUserStatsBox(val(personaje.UserIndex))
    
ElseIf hechizo.SubeHP = 2 Then
        
        daño = RandomNumber(hechizo.minHP, hechizo.MaxHP)
       
        daño = daño - getAbosrcionTotalRsistenciaMagica(personaje, daño)
        
        'marche
        daño = daño - (daño * getResistenciaMagica(personaje))  'Resistencia magica
        
        If daño < 0 Then daño = 0
        
        EnviarPaquete Paquetes.WavSnd, Chr$(hechizo.WAV), personaje.UserIndex, ToPCArea
        
        If Not personaje.flags.Meditando Then EnviarPaquete Paquetes.HechizoFX, ITS(personaje.Char.charIndex) & ByteToString(hechizo.FXgrh) & ITS(hechizo.loops), personaje.UserIndex, ToPCArea, personaje.pos.map
        
        personaje.Stats.minHP = personaje.Stats.minHP - daño
        
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(9) & criatura.Name & "," & daño, personaje.UserIndex
        
        
        Call SubirSkill(personaje.UserIndex, eSkills.ResistenciaMagica)
        
        'Muere
        If personaje.Stats.minHP < 1 Then
        
            personaje.Stats.minHP = 0
            
            If criatura.NPCtype = NPCTYPE_GUARDIAS Then
                RestarCriminalidad (personaje.UserIndex)
            End If
            
            Call UserDie(personaje.UserIndex, False)
            
            If criatura.MaestroUser > 0 Then
                Call ContarMuerte(personaje, UserList(criatura.MaestroUser))
                Call UsuarioMataAUsuario(personaje.UserIndex, criatura.MaestroUser)
            End If
        Else
            Call SendUserVida(val(personaje.UserIndex)) 'Marche 3-8
        End If
End If
If hechizo.Paraliza = 1 Or hechizo.Inmoviliza = 1 Then
     If personaje.flags.Paralizado = 0 And personaje.flags.Mimetizado = 0 Then
        EnviarPaquete Paquetes.WavSnd, Chr$(hechizo.WAV), personaje.UserIndex, ToPCArea
       
       If Not personaje.flags.Meditando Then EnviarPaquete Paquetes.HechizoFX, ITS(personaje.Char.charIndex) & ByteToString(hechizo.FXgrh) & ITS(hechizo.loops), personaje.UserIndex, ToPCArea, personaje.pos.map
               
        personaje.flags.Paralizado = 1
        personaje.Counters.Paralisis = modPersonaje.getIntervaloParalizado(personaje)
        
        Call enviarParalizado(personaje)
    End If
End If
Call SubirSkill(personaje.UserIndex, 1)
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next
End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal slot As Integer)
Dim hIndex As Integer
Dim j As Integer

'agregar nacho

hIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex
If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(132), UserIndex
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(slot), 1)
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple, Chr$(133), UserIndex
End If
End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)
    Dim ind As String
    ind = UserList(UserIndex).Char.charIndex
    EnviarPaquete Paquetes.SaidMagicWords, ITS(ind) & s, UserIndex, ToPCArea
End Sub

Function puedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(UserIndex).flags.Muerto = 1 Then
   EnviarPaquete Paquetes.MensajeSimple, Chr$(137), UserIndex
   puedeLanzar = False
   Exit Function
End If

'Si esta en un evento. El evento permite que use esta magia?
If Not UserList(UserIndex).evento Is Nothing Then
    If Not UserList(UserIndex).evento.puedeLanzar(CByte(HechizoIndex)) Then
        EnviarPaquete Paquetes.mensajeinfo, "El evento en el cual estás participando no permite que utilices este hechizo.", UserIndex, ToIndex
        puedeLanzar = False
        Exit Function
    End If
End If

'Tiene l mana suficiente?
If UserList(UserIndex).Stats.MinMAN >= getManaRequeridoHechizoParaPersonaje(UserList(UserIndex), hechizos(HechizoIndex)) Then
    'Tiene los skills suficientes?
    If UserList(UserIndex).Stats.UserSkills(eSkills.Magia) >= hechizos(HechizoIndex).MinSkill Then
        'Tiene la energia suficiente?
        If UserList(UserIndex).Stats.MinSta >= hechizos(HechizoIndex).StaRequerido Then
            puedeLanzar = True
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr$(134), UserIndex
            puedeLanzar = False
            Exit Function
        End If
    Else
        EnviarPaquete Paquetes.MensajeSimple, Chr$(135), UserIndex
        puedeLanzar = False
        Exit Function
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple, Chr$(136), UserIndex
    puedeLanzar = False
    Exit Function
End If

If puedeLanzarHechizoObjetos(UserList(UserIndex), hechizos(HechizoIndex)) = False Then
    EnviarPaquete Paquetes.mensajeinfo, "No tienes el poder suficiente en tus manos para lanzar este hechizo.", UserIndex, ToIndex
    puedeLanzar = False
    Exit Function
End If

End Function


Public Function guardarMascota(ByRef personaje As User, mascotaType As Integer) As Boolean

Dim loopSlot As Byte

For loopSlot = 1 To MAXMASCOTAS
    If personaje.MascotasGuardadas(loopSlot) = 0 Then
        'Guardo el tipo de mascotas
        personaje.MascotasGuardadas(loopSlot) = mascotaType
        'Aumento la cantidad de mascotas guardadas
        personaje.NroMascotasGuardadas = personaje.NroMascotasGuardadas + 1
        guardarMascota = True
        Exit Function
    End If
Next

guardarMascota = False
                
End Function
Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)


If MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Trigger = 3 Then Exit Sub

Dim IndexNPC As Integer
Dim h As Integer
Dim i As Integer
Dim j As Integer
Dim guardar As Boolean


h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.hechizo)

If h = 39 Then  'Invocar mascotas
    'No tiene nada que guardar ni nada que sacar
    If UserList(UserIndex).NroMascotasGuardadas = 0 And UserList(UserIndex).NroMacotas = 0 Then Exit Sub
    If UserList(UserIndex).NroMascotasGuardadas = MAXMASCOTAS And UserList(UserIndex).NroMacotas = MAXMASCOTAS Then Exit Sub
        
    If UserList(UserIndex).NroMascotasGuardadas = 0 Then
        guardar = True
    End If
    
    'Tiene prioridad el guardado
    'Primero intenta guardar las mascotas
    If guardar Then
        For i = 1 To MAXMASCOTAS
            
            'Este ciclo se va a hacer mientras tengo menos de 3 mascotas guardadas
            If UserList(UserIndex).NroMascotasGuardadas = MAXMASCOTAS Then Exit For
            
            If UserList(UserIndex).MascotasIndex(i) > 0 Then
                If NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
                    'Busco un slot libre
                    Call guardarMascota(UserList(UserIndex), UserList(UserIndex).MascotasType(i))
                    
                    'Quito al npc
                    Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
                End If
            End If
        Next
    End If

    'Si no guardo inteto invocar
   If Not guardar Then
        For i = 1 To MAXMASCOTAS
            'Hasta que tenga 3 mascotas invocadas como maximo
            If UserList(UserIndex).NroMacotas = MAXMASCOTAS Then Exit For
            
            If UserList(UserIndex).MascotasGuardadas(i) > 0 Then
                'Creo el NPC
                IndexNPC = SpawnNpc(UserList(UserIndex).MascotasGuardadas(i), UserList(UserIndex).pos, True, False)
                            
                If IndexNPC > MAXNPCS Then
                    Exit Sub 'No puedo crear el npc
                End If
                            
                'Establezco la relacion usuario amo npc.
                Call NPCs.establecerAmo(UserIndex, IndexNPC)
                'Hago que siga a su amo
                Call FollowAmo(IndexNPC)
                'No tiene mas a esa mascota guardada
                UserList(UserIndex).MascotasGuardadas(i) = 0
                'Reduzco el número de mascotas guardadas
                UserList(UserIndex).NroMascotasGuardadas = UserList(UserIndex).NroMascotasGuardadas - 1
                
                ' Solo una mascota
                Exit For
            End If
        Next i
    End If
    
    Call InfoHechizo(UserIndex)
    b = True
    Exit Sub

End If ' Fin hechizo invocar mascotas


If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

If MapInfo(UserList(UserIndex).pos.map).AntiHechizosPts = 1 Then Exit Sub

Dim index As Integer
Dim TargetPos As WorldPos

TargetPos.map = UserList(UserIndex).flags.TargetMap
TargetPos.x = UserList(UserIndex).flags.TargetX
TargetPos.y = UserList(UserIndex).flags.TargetY


For j = 1 To hechizos(h).cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then

      '  If UserList(UserIndex).Clase = eclases.Druida Then
                If InStr(1, hechizos(h).nombre, "elemental") > 0 Then
                  If Not PuedeTirarElementos(UserIndex) Then Exit Sub
                ElseIf hechizos(h).nombre = "Implorar ayuda" Then
                  If Not PuedeTirarImplorar(UserIndex) Then Exit Sub
                ElseIf hechizos(h).nombre = "Ayuda del Espiritu Indomable" Then
                  If Not PuedeTirarIndomable(UserIndex) Then Exit Sub
                End If
       '   End If

        IndexNPC = SpawnNpc(hechizos(h).NumNpc, TargetPos, True, False)

        If Not IndexNPC > MAXNPCS Then

            Call NPCs.establecerAmo(UserIndex, IndexNPC)

            If UCase$(hechizos(h).nombre) = UCase$("Invocar Elemental de fuego") Then
                NpcList(IndexNPC).Contadores.TiempoExistencia = IntervaloInvocacionFuego
            ElseIf UCase$(hechizos(h).nombre) = UCase$("Invocar Elemental de tierra") Then
                NpcList(IndexNPC).Contadores.TiempoExistencia = IntervaloInvocacionTierra
            ElseIf UCase$(hechizos(h).nombre) = UCase$("Invocar Elemental de agua") Then
                NpcList(IndexNPC).Contadores.TiempoExistencia = IntervaloInvocacionAgua
            Else
                NpcList(IndexNPC).Contadores.TiempoExistencia = IntervaloInvocacion
            End If

            EnviarPaquete Paquetes.WavSnd, Chr$(hechizos(h).WAV), UserIndex, ToPCArea, UserList(UserIndex).pos.map

            NpcList(IndexNPC).GiveGLD = 0

            Call FollowAmo(IndexNPC)

        End If
    Else
        Exit For
    End If
Next j

Call InfoHechizo(UserIndex)
b = True
End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case hechizos(uh).tipo
    Case uInvocacion
        Call HechizoInvocacion(UserIndex, b)
    Case uEstado
        Call HechizoTerrenoEstado(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, eSkills.Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - getManaRequeridoHechizoParaPersonaje(UserList(UserIndex), hechizos(uh))
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendUserStatsBox(UserIndex)
End If
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean 'Es verdadero si lo ataque
Dim tindex As Integer

b = False

tindex = UserList(UserIndex).flags.TargetUser

Select Case hechizos(uh).tipo

    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
    
        Call HechizoEstadoUsuario(UserIndex, b)
    
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
    
        Call HechizoPropUsuario(UserList(UserIndex), UserList(tindex), hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.hechizo)), b)
       
End Select


If b Then 'Le pego?

    'Si la magia es Dardo magico, flecha electrica, flecha magica o misil magico se ve la magia volando
    If uh = 2 Or (uh >= 6 And uh <= 8) Then
        EnviarPaquete Paquetes.FXh, ITS(UserList(UserIndex).Char.charIndex) & ITS(UserList(tindex).Char.charIndex) & ByteToString(0), tindex, ToPCArea, UserList(UserIndex).pos.map
    End If
    
    'Intento subir el skill
    Call SubirSkill(UserIndex, eSkills.Magia)
        
    'Resto la mana
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - getManaRequeridoHechizoParaPersonaje(UserList(UserIndex), hechizos(uh))
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    
    'Resto la energia
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    
    'Le envio al cliente la actualizacion de su estado
    Call SendUserStatsBox(UserIndex)
    
    'A la victima
    Call SendUserStatsBox(UserList(UserIndex).flags.TargetUser)
    
    'Ya no apunta a ningun usuario
    UserList(UserIndex).flags.TargetUser = 0
End If

End Sub

Public Function getManaRequeridoHechizoParaPersonaje(ByRef personaje As User, ByRef hechizo As tHechizo) As Integer
    
    If personaje.clase = eClases.asesino Then
        getManaRequeridoHechizoParaPersonaje = hechizo.ManaRequeridoAsesino
    ElseIf personaje.clase = eClases.Paladin Then
        getManaRequeridoHechizoParaPersonaje = hechizo.ManaRequeridoPaladin
    ElseIf personaje.clase = eClases.Bardo Then
        getManaRequeridoHechizoParaPersonaje = hechizo.ManaRequeridoBardo
    Else
        getManaRequeridoHechizoParaPersonaje = hechizo.ManaRequerido
    End If
    
End Function



Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
Dim b As Boolean
Select Case hechizos(uh).tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
       If b And UserList(UserIndex).flags.TargetNPC > 0 And (uh = 2 Or (uh >= 6 And uh <= 8)) Then EnviarPaquete Paquetes.FXh, ITS(UserList(UserIndex).Char.charIndex) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & ByteToString(0), UserList(UserIndex).flags.TargetNPC, ToNPCArea, NpcList(UserList(UserIndex).flags.TargetNPC).pos.map
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
       If b And UserList(UserIndex).flags.TargetNPC > 0 And (uh = 2 Or (uh >= 6 And uh <= 8)) Then EnviarPaquete Paquetes.FXh, ITS(UserList(UserIndex).Char.charIndex) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & ByteToString(0), UserIndex, ToPCArea, NpcList(UserList(UserIndex).flags.TargetNPC).pos.map
End Select
If b Then
    Call SubirSkill(UserIndex, eSkills.Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - getManaRequeridoHechizoParaPersonaje(UserList(UserIndex), hechizos(uh))
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendUserStatsBox(UserIndex)
End If
End Sub


Private Sub penarLanzamientoErrado(ByRef personaje As User, hechizo As tHechizo)
    
If hechizo.manaPenalidad = 0 Then
    Exit Sub
End If
                
If Not (personaje.clase = eClases.Mago Or personaje.clase = eClases.Clerigo Or personaje.clase = eClases.Bardo Or personaje.clase = eClases.Druida) Then
    Exit Sub
End If

' Restamos
Call RestToVar(personaje.Stats.MinMAN, hechizo.manaPenalidad, 0)

' Actualizamos la man
Call SendUserMana(personaje.UserIndex)

' Le avisamos
EnviarPaquete Paquetes.mensajeinfo, "Has fallado y perdido " & hechizo.manaPenalidad & " puntos de mana.", personaje.UserIndex, ToIndex
                    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LanzarHechizo
' DateTime  : 18/02/2007 22:28
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub LanzarHechizo(index As Integer, ByRef personaje As User)
Dim uh As Integer
   
uh = personaje.Stats.UserHechizos(index)

If uh = 0 Then Exit Sub

If puedeLanzar(personaje.UserIndex, uh) = False Then
    Exit Sub
End If

Select Case hechizos(uh).Target
    Case uUsuarios
        
        ' ¿Esta apuntando a usuario?
        If personaje.flags.TargetUser > 0 Then
            Call HandleHechizoUsuario(personaje.UserIndex, uh)
        Else
            ' Fallo. Debemos penarlo
            If hechizos(uh).manaPenalidad > 0 Then
                Call penarLanzamientoErrado(personaje, hechizos(uh))
            Else
                EnviarPaquete Paquetes.MensajeSimple, Chr$(138), personaje.UserIndex
            End If
        End If
    Case uNPC
        If personaje.flags.TargetNPC > 0 Then
            Call HandleHechizoNPC(personaje.UserIndex, uh)
        Else
            ' Fallo. Debemos penarlo
            If hechizos(uh).manaPenalidad > 0 Then
                Call penarLanzamientoErrado(personaje, hechizos(uh))
            Else
                EnviarPaquete Paquetes.MensajeSimple, Chr$(139), personaje.UserIndex
            End If
        End If
    Case uUsuariosYnpc
        If personaje.flags.TargetUser > 0 Then
            Call HandleHechizoUsuario(personaje.UserIndex, uh)
        ElseIf personaje.flags.TargetNPC > 0 Then
            Call HandleHechizoNPC(personaje.UserIndex, uh)
        ElseIf hechizos(uh).Mimetiza = 1 And (personaje.flags.TargetObj = 147 Or personaje.flags.TargetObj = 148) Then
            Call HandleHechizoNPC(personaje.UserIndex, uh)
        Else
            ' Fallo. Debemos penarlo
            If hechizos(uh).manaPenalidad > 0 Then
                Call penarLanzamientoErrado(personaje, hechizos(uh))
            Else
                EnviarPaquete Paquetes.MensajeSimple, Chr$(140), personaje.UserIndex
            End If
        End If
    Case uTerreno
        Call HandleHechizoTerreno(personaje.UserIndex, uh)
End Select


End Sub


Private Sub puedeLanzarHechizoBueno(ByRef atacante As User, ByRef victima As User)

End Sub

' Usuario lanza una magia sobre un usuario
Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim h As Integer, TU As Integer

'Hechizo que estoy lanzando
h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.hechizo)
'Victima
TU = UserList(UserIndex).flags.TargetUser

If hechizos(h).Revivir = 1 Then
   
   'El objetivo esta muerto?
    If UserList(TU).flags.Muerto = 0 Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(26), UserIndex
        b = False
        Exit Sub
    End If
    
    'Si esta navegando o retando no puede ser revivido
    If UserList(TU).flags.Navegando = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes resucitar a una barca.", UserIndex, ToIndex
        b = False
        Exit Sub
    End If
    
    If Not puedeAyudar(UserList(UserIndex), UserList(TU)) Then
        EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no te permite ayudar a este personaje.", UserIndex, ToIndex
        b = False
        Exit Sub
    End If
            
    'No lo puede revivir si esta en modo combate
    If UserList(TU).flags.modoCombate = True Then
        EnviarPaquete Paquetes.mensajeinfo, "El usuario esta en Modo Combate. No puedes revivirlo.", UserIndex
        b = False
        Exit Sub
    End If
    
    ' Sufre una penalización del 40% de vida. Si queda con menos de 10 se cancela.
    If UserList(UserIndex).clase = eClases.asesino Or UserList(UserIndex).clase = eClases.Paladin Then
        If UserList(UserIndex).Stats.minHP <= 10 Then
            EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado débil para lanzar este hechizo.", UserIndex
            b = False
        End If
    Else
        If UserList(UserIndex).Stats.minHP <= 15 Then
            EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado débil para lanzar este hechizo.", UserIndex
            b = False
        End If
    End If
    
    'Si no es un criminal sube su reputacion
    'If Not criminal(UserIndex) Then Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)

    'Lo revivo
    If UserList(TU).evento Is Nothing Then
        Call modResucitar.agregarResucitacion(UserList(UserIndex), UserList(TU))
    Else
        Call RevivirUsuario(UserList(TU))
    End If
    
    'Le aviso
    EnviarPaquete Paquetes.MensajeSimple, Chr$(143), UserIndex
        
    Call InfoHechizo(UserIndex)
    b = True
End If


If hechizos(h).Invisibilidad = 1 Then

    'Si esta mimetizado no puede lanzar invisibilidad
    'If UserList(UserIndex).flags.Mimetizado = 1 Then Exit Sub
    
    If MapInfo(UserList(UserIndex).pos.map).AntiHechizosPts = 1 Then Exit Sub
    
    If Not puedeLanzarleInvisibilidad(UserList(UserIndex), UserList(TU)) Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes lanzar invisibilidad a este personaje. Debe pertenecer a tu clan o a tu party.", UserIndex, ToIndex
        b = False
        Exit Sub
    End If
        
    'No se puede poner invisible a un muerto
    If UserList(TU).flags.Muerto = 1 Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(88), UserIndex
        b = False
        Exit Sub
    End If
    
    'Si ya esta invisible no puede tirarle denuevo el hechizo
    If UserList(TU).flags.Invisible = 1 Then
       EnviarPaquete Paquetes.MensajeSimple2, Chr$(133), UserIndex
        b = False
        Exit Sub
    End If
    
    If Not puedeAyudar(UserList(UserIndex), UserList(TU)) Then
        EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no te permite ayudar a este personaje.", UserIndex, ToIndex
        b = False
        Exit Sub
    End If
    
   UserList(TU).flags.Invisible = 1
   EnviarPaquete Paquetes.Invisible, ITS(UserList(TU).Char.charIndex) & ByteToString(getInvisibilidadTipo(UserList(TU))), UserIndex, ToMap
   Call InfoHechizo(UserIndex)
   b = True
End If

If hechizos(h).Envenena = 1 Then
        If UserList(TU).flags.Envenenado = 1 Then Exit Sub
        
        If Not puedeAtacar(UserList(UserIndex), UserList(TU)) Then Exit Sub
        
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserList(UserIndex), UserList(TU))
            UserList(TU).flags.Envenenado = 1
            EnviarPaquete Paquetes.EstaEnvenenado, "", TU, ToIndex
            Call InfoHechizo(UserIndex)
            b = True
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr$(145), UserIndex
        End If
End If

If hechizos(h).CuraVeneno = 1 Then

        'Gorlok - No curar si no está/s envenenado.
        If UserList(TU).flags.Envenenado = 0 Then
            If UserIndex <> TU Then
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(103), UserIndex
            Else
               EnviarPaquete Paquetes.MensajeSimple2, Chr$(104), UserIndex
            End If
            Exit Sub
        End If

        If Not puedeAyudar(UserList(UserIndex), UserList(TU)) Then
            EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no te permite ayudar a este personaje.", UserIndex, ToIndex
            b = False
            Exit Sub
        End If

        UserList(TU).flags.Envenenado = 0
        EnviarPaquete Paquetes.EstaEnvenenado, "", TU, ToIndex
        
        Call InfoHechizo(UserIndex)
        b = True
End If

If hechizos(h).Paraliza = 1 Or hechizos(h).Inmoviliza = 1 Then
    'Prohibido paralizarse a si mismo - byGorlok 2005-03-25
    If UserIndex = TU Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(20), UserIndex
        Exit Sub
    End If
    
    If UserList(TU).flags.Paralizado = 1 Then
        Exit Sub
    End If
    
    If Not puedeAtacar(UserList(UserIndex), UserList(TU)) Then Exit Sub

    'If UserList(userIndex).flags.Invisible = 1 Then
    '    EnviarPaquete Paquetes.mensajeinfo, "Estas demasiado lejos para lanzar este hechizo.", userIndex, ToIndex
    '    Exit Sub
    'End If
    
    If UserList(TU).clase = eClases.Cazador And UserList(TU).flags.Oculto = 1 Then
        Call quitarOcultamiento(UserList(TU))
    End If
    
    Call UsuarioAtacadoPorUsuario(UserList(UserIndex), UserList(TU))
        
    UserList(TU).flags.Paralizado = 1
    UserList(TU).flags.paralizadoPor = UserIndex
    UserList(TU).Counters.Paralisis = modPersonaje.getIntervaloParalizado(UserList(TU))
    
    Call enviarParalizado(UserList(TU))
    
    Call InfoHechizo(UserIndex)
    b = True
End If

If hechizos(h).RemoverParalisis = 1 Then
    If Not puedeAyudar(UserList(UserIndex), UserList(TU)) Then
        EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no te permite ayudar a este personaje.", UserIndex, ToIndex
        b = False
        Exit Sub
    End If
    
    If UserList(TU).flags.Paralizado = 1 Then
        UserList(TU).flags.Paralizado = 0
        EnviarPaquete Paquetes.NoParalizado2, "", TU
        Call InfoHechizo(UserIndex)
        b = True
    End If
    
End If

If hechizos(h).Mimetiza = 1 Then
    ' Reglas:
    ' - Personaje
    '   a) No te podes mimetizar con vos mismo
    '   b) No te podés mimetizar si ya estas mimetizado
    ' - Mapa
    '   a) No valido en Zona Segura
    '   b) No valido en mapas especiales
    ' - Objetivo
    '   X) No te podes mimetizar con un personaje muerto
    '   b) No te podés mimetizar con GameMasters
    '   c) No te podés mimetizar con personajes que Navegan
    
    ' ¿Muerto?
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
    
    ' No se puede mimetizar con el mismo
    If UserIndex = TU Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes mimetizarte con vos mismo.", UserIndex, ToIndex
        Exit Sub
    End If
        
    ' No podés recargar el hechizo
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "Ya te encuentras transformado. El hechizo no ha tenido efecto.", UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No vale en Zona Segura
    If MapInfo(UserList(UserIndex).pos.map).Pk = False Then
        EnviarPaquete Paquetes.mensajeinfo, "No te puedes mimetizar en zona segura.", UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No vale en Mapas Especiales
    If MapInfo(UserList(UserIndex).pos.map).AntiHechizosPts = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "No te puedes mimetizar aquí.", UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No vale mimetizarte con personajes que estan navegando
    If UserList(TU).flags.Navegando = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "No logras ver el aspecto del personaje. Este se encuentra dentro de una barca.", UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No vale con un muerto
    ' If UserList(TU).flags.Muerto = 1 Then
    '   EnviarPaquete Paquetes.mensajeinfo, "No puedes mimetizarte con un espíritu.", UserIndex, ToIndex
    '    Exit Sub
    ' End If
    
    ' No vale con Game Masters
    If UserList(TU).flags.Privilegios > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes mimetizarte con los administradores del juego.", UserIndex, ToIndex
        Exit Sub
    End If

    ' Ejecutamos la mimetizacion
    Call modMimetismo.DoMimetizarConPersonaje(UserList(UserIndex), UserList(TU))
    
    ' Informacion del Hechizo
    Call InfoHechizo(UserIndex)
    b = True
End If

End Sub

Private Function puedeLanzarleInvisibilidad(ByRef lanzador As User, ByRef receptor As User) As Boolean
    
    If lanzador.UserIndex = receptor.UserIndex Then
        puedeLanzarleInvisibilidad = True
        Exit Function
    End If
    
    If lanzador.PartyIndex > 0 And lanzador.PartyIndex = receptor.PartyIndex Then
        puedeLanzarleInvisibilidad = True
        Exit Function
    End If
    
    If Not lanzador.ClanRef Is Nothing And lanzador.ClanRef Is receptor.ClanRef Then
        puedeLanzarleInvisibilidad = True
        Exit Function
    End If

    puedeLanzarleInvisibilidad = False

End Function

Sub HechizoEstadoNPC(ByVal npcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
'[Misery_Ezequiel 04/06/05]

If UserList(UserIndex).flags.TargetObj <> 147 And UserList(UserIndex).flags.TargetObj <> 148 Then
    
    If NpcList(npcIndex).InmuneAHechizos = 1 Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(87), UserIndex
        Exit Sub
    End If

    If hechizos(hIndex).Invisibilidad = 1 Then
        Call InfoHechizo(UserIndex)
        b = True
    End If

End If
  
If hechizos(hIndex).Mimetiza = 1 Then

    ' Reglas:
    ' - Personaje
    '   a) No te podes mimetizar con vos mismo
    ' - Mapa
    '   a) No valido en Zona Segura
    '   b) No valido en mapas especiales
    ' - Objetivo
    '
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

    ' No te podes mimetizar si ya estas mimetizado
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "Ya te encuentras transformado. El hechizo no ha tenido efecto", UserIndex, ToIndex
        Exit Sub
    End If

    ' Nov valido en zona segura
    If MapInfo(UserList(UserIndex).pos.map).Pk = False Then
        EnviarPaquete Paquetes.mensajeinfo, "No te puedes mimetizar en zona segura.", UserIndex, ToIndex
        Exit Sub
    End If

    ' No vale en Mapas Especiales
    If MapInfo(UserList(UserIndex).pos.map).AntiHechizosPts = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "No te puedes mimetizar aquí.", UserIndex, ToIndex
        Exit Sub
    End If
    
    'TODO Esto está por si se mimetiza con un arbol
    If npcIndex = 0 Then Exit Sub
    
    Call DoMimetizarConCriatura(UserList(UserIndex), NpcList(npcIndex))
    
    Call InfoHechizo(UserIndex)
    b = True
End If


If hechizos(hIndex).Paraliza = 1 Then
    If NpcList(npcIndex).flags.AfectaParalisis = 0 Then
    
        If Not AntiRoboNpc.puedeLucharContraELNPC(NpcList(npcIndex), UserList(UserIndex)) Then
            Exit Sub
        End If
        
        If Not PuedeAtacarNPC(UserList(UserIndex), NpcList(npcIndex)) Then
            Exit Sub
        End If
                
        Call InfoHechizo(UserIndex)
        
        NpcList(npcIndex).flags.Paralizado = 1
        NpcList(npcIndex).flags.Inmovilizado = 0
        NpcList(npcIndex).Contadores.Paralisis = IntervaloParalizadoNPC
        b = True
    Else
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(28), UserIndex
    End If
    
    
End If

If hechizos(hIndex).RemoverParalisis = 1 Then
    If (NpcList(npcIndex).flags.Paralizado = 1 Or NpcList(npcIndex).flags.Inmovilizado = 1) And (NpcList(npcIndex).MaestroUser = UserIndex Or UserList(UserIndex).flags.Privilegios > 0) Then
        Call InfoHechizo(UserIndex)
        NpcList(npcIndex).flags.Paralizado = 0
        NpcList(npcIndex).flags.Inmovilizado = 0
        NpcList(npcIndex).Contadores.Paralisis = 0
        b = True
    Else
        EnviarPaquete Paquetes.MensajeSimple, Chr$(146), UserIndex
    End If
End If
'[/wizard]

If hechizos(hIndex).Inmoviliza = 1 Then
   If NpcList(npcIndex).flags.AfectaParalisis = 0 Then
      If NpcList(npcIndex).flags.Paralizado = 1 Then Exit Sub
        NpcList(npcIndex).flags.Inmovilizado = 1
        NpcList(npcIndex).flags.Paralizado = 0
        NpcList(npcIndex).Contadores.Paralisis = IntervaloParalizadoNPC
        Call InfoHechizo(UserIndex)
        b = True
   Else
      EnviarPaquete Paquetes.MensajeSimple2, Chr$(28), UserIndex
   End If
End If
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal npcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)

Dim daño As Long
Dim otroUsuario As Integer
Dim nIndex As Integer
Dim AnguloNPC As Single

If NpcList(npcIndex).InmuneAHechizos = 1 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(144), UserIndex
    Exit Sub
End If

' Chequeo robo de NPC
If Not AntiRoboNpc.puedeLucharContraELNPC(NpcList(npcIndex), UserList(UserIndex)) Then
    Exit Sub
End If

'Salud

If hechizos(hIndex).SubeHP = 1 Then
    If NpcList(npcIndex).Stats.minHP = NpcList(npcIndex).Stats.MaxHP Then
       EnviarPaquete Paquetes.MensajeSimple2, Chr$(28), UserIndex
       Exit Sub
    Else
        daño = RandomNumber(hechizos(hIndex).minHP, hechizos(hIndex).MaxHP)
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        Call InfoHechizo(UserIndex)
        Call AddtoVar(NpcList(npcIndex).Stats.minHP, daño, NpcList(npcIndex).Stats.MaxHP)
        EnviarPaquete Paquetes.MensajeFight, "Has curado " & daño & " puntos de salud a la criatura.", UserIndex
        b = True
    End If
ElseIf hechizos(hIndex).SubeHP = 2 Then
    
    If Not PuedeAtacarNPC(UserList(UserIndex), NpcList(npcIndex)) Then
        b = False
        Exit Sub
    End If
    
    daño = RandomNumber(hechizos(hIndex).minHP, hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
         
    If hechizos(hIndex).StaffAffected Then
        daño = daño * calcularStaffAfected(UserList(UserIndex))
    End If
    
    Call InfoHechizo(UserIndex)
    
    b = True
    
    Call NpcAtacado(npcIndex, UserIndex)
    
    If NpcList(npcIndex).flags.Snd2 > 0 Then EnviarPaquete Paquetes.WavSnd, Chr$(NpcList(npcIndex).flags.Snd2), UserIndex, ToPCArea, UserList(UserIndex).pos.map
    
    EnviarPaquete Paquetes.TXA, ITS(NpcList(npcIndex).pos.x) & ITS(NpcList(npcIndex).pos.y) & ITS(daño) & ITS(distancia(NpcList(npcIndex).pos, UserList(UserIndex).pos)), UserIndex, ToPCArea, NpcList(npcIndex).pos.map
    NpcList(npcIndex).Stats.minHP = NpcList(npcIndex).Stats.minHP - daño
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(10) & daño, UserIndex, ToIndex
    
    Call CalcularDarExp(UserList(UserIndex), NpcList(npcIndex), daño)
    
    If NpcList(npcIndex).Stats.minHP < 1 Then
        NpcList(npcIndex).Stats.minHP = 0
        
        Call UsuarioMataNPC(UserList(UserIndex), NpcList(npcIndex))
        
        nIndex = MuereNpc(NpcList(npcIndex))
        
        ' Chequeo si en el mapa no hay mas de 10 usuarios para enviar el angulo del nuevo NPC.
        If nIndex > 0 Then
            If DeboEnviarAngulo(UserList(UserIndex).pos.map) Then
                AnguloNPC = Angulo(NpcList(nIndex).pos.x, NpcList(nIndex).pos.y, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)
                EnviarPaquete Paquetes.AnguloNPC, ITS(AnguloNPC), UserIndex, ToIndex
            End If
        End If
        
    End If
    
End If
End Sub

' Este valor es el multiplicador por el cual se multiplica el daño del hechizo dependiendo el poder del usuario con su herramienta/anillo
Private Function calcularStaffAfected(ByRef personaje As User) As Single

    If personaje.clase = eClases.Mago Then
        If personaje.Invent.WeaponEqpObjIndex > 0 Then
            calcularStaffAfected = (ObjData(personaje.Invent.WeaponEqpObjIndex).StaffDamageBonus + 70) / 100
        Else
            calcularStaffAfected = 0.8
        End If
        
        Exit Function
    ElseIf personaje.clase = eClases.Clerigo Or personaje.clase = eClases.Bardo Or personaje.clase = eClases.Druida Then
        
        If personaje.Invent.AnilloEqpObjIndex > 0 Then
            calcularStaffAfected = (ObjData(personaje.Invent.AnilloEqpObjIndex).StaffDamageBonus + 70) / 100
            Exit Function
        End If
            
    End If
    
    calcularStaffAfected = 1
End Function
Sub InfoHechizo(ByVal UserIndex As Integer)
    Dim h As Integer
    
    'Obtengo el hechizo
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.hechizo)
    
    'Muestro las palabras magicas
    Call DecirPalabrasMagicas(hechizos(h).PalabrasMagicas, UserIndex)
    
    'Carteles en consola y efectos
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
    
        If Not UserList(UserList(UserIndex).flags.TargetUser).flags.Meditando Then
            If UserList(UserIndex).flags.TargetUser <> UserIndex Then
                EnviarPaquete Paquetes.HechizoFX, ITS(UserList(UserList(UserIndex).flags.TargetUser).Char.charIndex) & ByteToString(hechizos(h).FXgrh) & ITS(hechizos(h).loops) & Chr$(hechizos(h).WAV), UserIndex, ToPCArea, UserList(UserIndex).pos.map
            Else
                EnviarPaquete Paquetes.HechizoFX, ITS(UserList(UserList(UserIndex).flags.TargetUser).Char.charIndex) & ByteToString(hechizos(h).FXgrh) & ITS(hechizos(h).loops) & Chr$(hechizos(h).WAV), UserIndex, ToPCArea, UserList(UserIndex).pos.map
            End If
        Else
            'Dejo de meditar
            EnviarPaquete Paquetes.Meditando, "", UserList(UserIndex).flags.TargetUser
            UserList(UserList(UserIndex).flags.TargetUser).flags.Meditando = False
            UserList(UserList(UserIndex).flags.TargetUser).Char.FX = 0
            UserList(UserList(UserIndex).flags.TargetUser).Char.loops = 0
            EnviarPaquete Paquetes.HechizoFX, ITS(UserList(UserList(UserIndex).flags.TargetUser).Char.charIndex) & ByteToString(0) & ITS(0), UserList(UserIndex).flags.TargetUser, ToPCArea, UserList(UserList(UserIndex).flags.TargetUser).pos.map
            
            'Le mando solamente el sonido
            EnviarPaquete Paquetes.WavSnd, Chr$(hechizos(h).WAV), UserIndex, ToPCArea, UserList(UserIndex).pos.map
        End If
        
        '¿Se tiro el hechizo asi mismo?
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            EnviarPaquete Paquetes.MensajeFight, hechizos(h).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name, UserIndex
            EnviarPaquete Paquetes.MensajeFight, UserList(UserIndex).Name & " " & hechizos(h).TargetMsg, UserList(UserIndex).flags.TargetUser
        Else
            EnviarPaquete Paquetes.MensajeFight, hechizos(h).PropioMsg, UserIndex
        End If
        
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        EnviarPaquete Paquetes.HechizoFX, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & ByteToString(hechizos(h).FXgrh) & ITS(hechizos(h).loops) & Chr$(hechizos(h).WAV), UserIndex, ToPCArea, UserList(UserIndex).pos.map
        EnviarPaquete Paquetes.MensajeFight, hechizos(h).HechizeroMsg & " la criatura.", UserIndex '
    End If
    
   
End Sub


Private Function puedeLanzarHechizoObjetos(ByRef personaje As User, ByRef hechizo As tHechizo) As Boolean

puedeLanzarHechizoObjetos = False


If personaje.clase = eClases.Druida Then
    
    If personaje.Invent.AnilloEqpObjIndex = ANILLO_PLATA_M2 Then
    
        puedeLanzarHechizoObjetos = True
        
    ElseIf personaje.Invent.AnilloEqpObjIndex = ANILLO_PLATA Then
        If hechizo.nombre = "Resucitar" Then Exit Function
        If hechizo.nombre = "Invocar elemental de fuego" Then Exit Function
        If hechizo.nombre = "Invocar elemental de agua" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Mimetismo" Then Exit Function
        If hechizo.nombre = "Ayuda del Espiritu Indomable" Then Exit Function
        If hechizo.nombre = "Implorar ayuda" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
        
    ElseIf personaje.Invent.AnilloEqpObjIndex = ANILLO_PLATA_M1 Then
        If hechizo.nombre = "Resucitar" Then Exit Function
        If hechizo.nombre = "Ayuda del Espiritu Indomable" Then Exit Function
        If hechizo.nombre = "Implorar ayuda" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
    Else
        If hechizo.nombre = "Inmovilizar" Then Exit Function
        If hechizo.nombre = "Tormenta de fuego" Then Exit Function
        If hechizo.nombre = "Descarga Eléctrica" Then Exit Function
        If hechizo.nombre = "Apocalípsis" Then Exit Function
        If hechizo.nombre = "Resucitar" Then Exit Function
        If hechizo.nombre = "Invocar elemental de fuego" Then Exit Function
        If hechizo.nombre = "Invocar elemental de agua" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Invisibilidad" Then Exit Function
        If hechizo.nombre = "Invocar mascotas" Then Exit Function
        If hechizo.nombre = "Mimetismo" Then Exit Function
        If hechizo.nombre = "Ayuda del Espiritu Indomable" Then Exit Function
        If hechizo.nombre = "Implorar ayuda" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
    End If
    
ElseIf personaje.clase = eClases.Bardo Then

    If personaje.Invent.AnilloEqpObjIndex = LAUDMAGICO_M1 Then
    
        puedeLanzarHechizoObjetos = True
        
    ElseIf personaje.Invent.AnilloEqpObjIndex = LAUDMAGICO Then
    
        If hechizo.nombre = "Resucitar" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
        
    ElseIf personaje.Invent.AnilloEqpObjIndex = FLAUTA_MAGICA Then
    
        If hechizo.nombre = "Invocar elemental de fuego" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Resucitar" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
        
    Else
    
        If hechizo.nombre = "Descarga Eléctrica" Then Exit Function
        If hechizo.nombre = "Apocalípsis" Then Exit Function
        If hechizo.nombre = "Resucitar" Then Exit Function
        If hechizo.nombre = "Invocar elemental de fuego" Then Exit Function
        If hechizo.nombre = "Invocar elemental de agua" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Inmovilizar" Then Exit Function
        If hechizo.nombre = "Tormenta de fuego" Then Exit Function
        If hechizo.nombre = "Invisibilidad" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
        
    End If
    
ElseIf personaje.clase = eClases.Clerigo Then
    
    If personaje.Invent.AnilloEqpObjIndex = CRUZTEJO Then
        puedeLanzarHechizoObjetos = True
    ElseIf personaje.Invent.AnilloEqpObjIndex = CRUZMADERA Then
        If hechizo.nombre = "Invocar elemental de fuego" Then Exit Function
        If hechizo.nombre = "Invocar elemental de agua" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Apocalípsis" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
    Else   ' No tiene nada o tiene otro anillo
        If hechizo.nombre = "Invocar elemental de fuego" Then Exit Function
        If hechizo.nombre = "Invocar elemental de agua" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Descarga Eléctrica" Then Exit Function
        If hechizo.nombre = "Apocalípsis" Then Exit Function
        If hechizo.nombre = "Resucitar" Then Exit Function
                
        puedeLanzarHechizoObjetos = True
    End If
    
ElseIf personaje.clase = eClases.Mago Then

    If personaje.Invent.WeaponEqpObjIndex = VARA_FRESNO Then

        If hechizo.nombre = "Invocar elemental de agua" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Apocalípsis" Then Exit Function
        If hechizo.nombre = "Resucitar" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
        
    ElseIf personaje.Invent.WeaponEqpObjIndex = BASTON_NUDOSO Then

        If hechizo.nombre = "Resucitar" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
        
    ElseIf personaje.Invent.WeaponEqpObjIndex = BACULO_ENGARZADO Then
    
        puedeLanzarHechizoObjetos = True
        
    Else
        If hechizo.nombre = "Inmovilizar" Then Exit Function
        If hechizo.nombre = "Tormenta de fuego" Then Exit Function
        If hechizo.nombre = "Descarga Eléctrica" Then Exit Function
        If hechizo.nombre = "Apocalípsis" Then Exit Function
        If hechizo.nombre = "Invocar elemental de fuego" Then Exit Function
        If hechizo.nombre = "Invocar elemental de agua" Then Exit Function
        If hechizo.nombre = "Invocar elemental de tierra" Then Exit Function
        If hechizo.nombre = "Resucitar" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
    End If

ElseIf personaje.clase = eClases.Paladin Then
    If personaje.Invent.BrasaleteEqpObjIndex = ANILLO_RESISTENCIA Then
        puedeLanzarHechizoObjetos = True
    Else
        If hechizo.nombre = "Resucitar" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
    End If
ElseIf personaje.clase = eClases.asesino Then
    If personaje.Invent.BrasaleteEqpObjIndex = ANILLO_RESISTENCIA Or personaje.Invent.BrasaleteEqpObjIndex = ANILLO_PROTECCION Then
        puedeLanzarHechizoObjetos = True
    Else
        If hechizo.nombre = "Resucitar" Then Exit Function
        
        puedeLanzarHechizoObjetos = True
    End If
Else
    puedeLanzarHechizoObjetos = True
End If



End Function


Sub HechizoPropUsuario(ByRef atacante As User, ByRef victima As User, ByRef hechizo As tHechizo, ByRef b As Boolean)

Dim h As Integer
Dim daño As Integer
Dim tempChr As Integer



'Restaura el hambre
If hechizo.SubeHam = 1 Then

    daño = RandomNumber(hechizo.minham, hechizo.MaxHam)
    Call AddtoVar(victima.Stats.minham, daño, victima.Stats.MaxHam)
    
    Call InfoHechizo(atacante.UserIndex)
    
    If atacante.UserIndex <> victima.UserIndex Then
        EnviarPaquete Paquetes.MensajeFight, "Le has restaurado " & daño & " puntos de hambre a " & victima.Name, atacante.UserIndex
        EnviarPaquete Paquetes.MensajeFight, atacante.Name & " te ha restaurado " & daño & " puntos de hambre.", victima.UserIndex
    Else
        EnviarPaquete Paquetes.MensajeFight, "Te has restaurado " & daño & " puntos de hambre.", atacante.UserIndex
    End If
    
    Call EnviarHambreYsed(victima.UserIndex)
    
    b = True
'Provocar hambre
ElseIf hechizo.SubeHam = 2 Then

    If MapInfo(atacante.pos.map).AntiHechizosPts = 1 Then Exit Sub
    If Not puedeAtacar(atacante, victima) Then Exit Sub

    If atacante.UserIndex <> victima.UserIndex Then
        Call UsuarioAtacadoPorUsuario(atacante, victima)
    Else
        'No se puede tirar hambre a si mismo
        'Tipo que sino seria re ghandi! :D
        EnviarPaquete Paquetes.MensajeSimple, Chr$(145), atacante.UserIndex
        Exit Sub
    End If
    
    Call InfoHechizo(atacante.UserIndex)
    daño = RandomNumber(hechizo.minham, hechizo.MaxHam)

    victima.Stats.minham = victima.Stats.minham - daño
    If victima.Stats.minham < 0 Then victima.Stats.minham = 0
    If atacante.UserIndex <> victima.UserIndex Then
        EnviarPaquete Paquetes.MensajeFight, "Le has quitado " & daño & " puntos de hambre a " & victima.Name, atacante.UserIndex
        EnviarPaquete Paquetes.MensajeFight, atacante.Name & " te ha quitado " & daño & " puntos de hambre.", victima.UserIndex
    Else
        EnviarPaquete Paquetes.MensajeFight, "Te has quitado " & daño & " puntos de hambre.", atacante.UserIndex
    End If
     
    If victima.Stats.minham < 1 Then
        victima.Stats.minham = 0
        victima.flags.Hambre = 1
    End If
    
    Call EnviarHambreYsed(victima.UserIndex)
    b = True
End If

'Restaura la sed
If hechizo.SubeSed = 1 Then

    Call InfoHechizo(atacante.UserIndex)
    Call AddtoVar(victima.Stats.minAgu, daño, victima.Stats.MaxAGU)

    If atacante.UserIndex <> victima.UserIndex Then
        EnviarPaquete Paquetes.MensajeFight, "Le has restaurado " & daño & " puntos de sed a " & victima.Name, atacante.UserIndex
        EnviarPaquete Paquetes.MensajeFight, atacante.Name & " te ha restaurado " & daño & " puntos de sed.", victima.UserIndex
    Else
        EnviarPaquete Paquetes.MensajeFight, "Te has restaurado " & daño & " puntos de sed.", atacante.UserIndex
    End If
    
    Call EnviarHambreYsed(victima.UserIndex)
    b = True

'Provocar sed
ElseIf hechizo.SubeSed = 2 Then

    If Not puedeAtacar(atacante, victima) Then Exit Sub
    
    If atacante.UserIndex <> victima.UserIndex Then
        Call UsuarioAtacadoPorUsuario(atacante, victima)
    End If
    
    Call InfoHechizo(atacante.UserIndex)
    
    victima.Stats.minAgu = victima.Stats.minAgu - daño
    
    If atacante.UserIndex <> victima.UserIndex Then
        EnviarPaquete Paquetes.MensajeFight, "Le has quitado " & daño & " puntos de sed a " & victima.Name, atacante.UserIndex
        EnviarPaquete Paquetes.MensajeFight, atacante.Name & " te ha quitado " & daño & " puntos de sed.", victima.UserIndex
    Else
        EnviarPaquete Paquetes.MensajeFight, "Te has quitado " & daño & " puntos de sed.", atacante.UserIndex
    End If
    
    If victima.Stats.minAgu < 1 Then
            victima.Stats.minAgu = 0
            victima.flags.Sed = 1
    End If
    
    b = True
End If
' <-------- Agilidad ---------->
If hechizo.SubeAgilidad = 1 Then

   If victima.flags.Muerto = 1 Then Exit Sub
     
    If atacante.UserIndex <> victima.UserIndex Then
        If Not puedeAyudar(atacante, victima) Then
            EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no te permite ayudar a este personaje.", atacante.UserIndex, ToIndex
            b = False
            Exit Sub
        End If
    End If
    
    Call InfoHechizo(atacante.UserIndex)
    
    daño = RandomNumber(hechizo.MinAgilidad, hechizo.MaxAgilidad)

    Call modPersonaje.incrementarAgilidad(victima, daño, IntervaloDuracionPociones)

    b = True
    
ElseIf hechizo.SubeAgilidad = 2 Then

    If Not puedeAtacar(atacante, victima) Then Exit Sub
        
    If atacante.UserIndex <> victima.UserIndex Then
        Call UsuarioAtacadoPorUsuario(atacante, victima)
    End If
    
    Call InfoHechizo(atacante.UserIndex)

    daño = RandomNumber(hechizo.MinAgilidad, hechizo.MaxAgilidad)
    
    Call modPersonaje.reducirAgilidad(atacante, daño, 25000)
    
    b = True
    
End If
' <-------- Fuerza ---------->
'Subir fuerza
If hechizo.SubeFuerza = 1 Then
   
    If victima.flags.Muerto = 1 Then Exit Sub
  
    If victima.UserIndex <> atacante.UserIndex Then
        If Not puedeAyudar(atacante, victima) Then
            EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no te permite ayudar a este personaje.", atacante.UserIndex, ToIndex
            b = False
            Exit Sub
        End If
    End If
    
    Call InfoHechizo(atacante.UserIndex)
    
    'Calculo cuanto le va a aumentar
    daño = RandomNumber(hechizo.MinFuerza, hechizo.MaxFuerza)
    
    Call modPersonaje.incrementarFuerza(victima, daño, IntervaloDuracionPociones)
      
    b = True
ElseIf hechizo.SubeFuerza = 2 Then
    ' Restar fuerza
    If Not puedeAtacar(atacante, victima) Then Exit Sub

    If atacante.UserIndex <> victima.UserIndex Then
        Call UsuarioAtacadoPorUsuario(atacante, victima)
    End If
    Call InfoHechizo(atacante.UserIndex)

    daño = RandomNumber(hechizo.MinFuerza, hechizo.MaxFuerza)
    
    Call modPersonaje.reducirFuerza(victima, daño, IntervaloDuracionPociones)
    
    b = True
End If

'Salud
If hechizo.SubeHP = 1 Then

    If victima.flags.Muerto = 1 Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(88), atacante.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' Gorlok - No curar si esta/s con toda la vida.
    If victima.Stats.minHP >= victima.Stats.MaxHP Then
        If atacante.UserIndex <> victima.UserIndex Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(29), atacante.UserIndex
        Else
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(117), atacante.UserIndex
        End If
        Exit Sub
    End If
    
    If atacante.UserIndex <> victima.UserIndex Then
        If Not puedeAyudar(atacante, victima) Then
            EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no te permite ayudar a este personaje.", atacante.UserIndex, ToIndex
            b = False
            Exit Sub
        End If
    End If
  
    
    daño = RandomNumber(hechizo.minHP, hechizo.MaxHP)
    daño = daño + Porcentaje(daño, 3 * atacante.Stats.ELV)

    If hechizo.StaffAffected Then
        daño = daño * calcularStaffAfected(atacante)
    End If
    
    If daño < 0 Then daño = 2
  
    Call InfoHechizo(atacante.UserIndex)
    Call AddtoVar(victima.Stats.minHP, daño, _
         victima.Stats.MaxHP)
         
    
    If atacante.UserIndex <> victima.UserIndex Then
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(11) & daño & "," & victima.Name, atacante.UserIndex
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(12) & atacante.Name & "," & daño, victima.UserIndex
    Else
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(13) & daño, atacante.UserIndex
    End If
    b = True
ElseIf hechizo.SubeHP = 2 Then

    If atacante.UserIndex = victima.UserIndex Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(145), atacante.UserIndex
        Exit Sub
    End If
        
    If Not puedeAtacar(atacante, victima) Then Exit Sub

    daño = RandomNumber(hechizo.minHP, hechizo.MaxHP)
    
    daño = daño + Porcentaje(daño, 3 * atacante.Stats.ELV)
    
    ' Penalizacion
    If victima.Stats.ELV >= 25 And victima.Invent.ArmourEqpObjIndex > 0 Then
        If ObjData(victima.Invent.ArmourEqpObjIndex).MaxDef < 15 Then
            daño = daño * 1.15
        End If
    End If
    
    
    ' Items magicos
    If hechizo.StaffAffected Then
        daño = daño * calcularStaffAfected(atacante)
    End If
        
    If atacante.clase <> eClases.Mago Then daño = daño * 1.05
    
    'Resistencia magica
    daño = daño - getAbosrcionTotalRsistenciaMagica(victima, daño)
  
    daño = daño - (daño * getResistenciaMagica(victima))
    
    If daño < 0 Then daño = 0
    
    If atacante.UserIndex <> victima.UserIndex Then
        Call UsuarioAtacadoPorUsuario(atacante, victima)
    End If
    
    Call InfoHechizo(atacante.UserIndex)
    
    victima.Stats.minHP = victima.Stats.minHP - daño
    EnviarPaquete Paquetes.TXA, ITS(victima.pos.x) & ITS(victima.pos.y) & ITS(daño) & ITS(distancia(atacante.pos, victima.pos)), victima.UserIndex, ToPCArea, victima.pos.map
    EnviarPaquete Paquetes.MensajeFight, "Le has quitado " & daño & " puntos de vida a " & victima.Name, atacante.UserIndex
    EnviarPaquete Paquetes.MensajeFight, atacante.Name & " te ha quitado " & daño & " puntos de vida.", victima.UserIndex
    
    ' Incremento Skill
    Call SubirSkill(victima.UserIndex, eSkills.ResistenciaMagica)
    
    If victima.Stats.minHP < 1 Then
        Call ContarMuerte(victima, atacante)
        victima.Stats.minHP = 0
        Call UsuarioMataAUsuario(victima.UserIndex, atacante.UserIndex)
    End If
    b = True
End If

' <--------Ilimited All---------->
If hechizo.AgiUpAndFuer = 1 Then
    Call InfoHechizo(atacante.UserIndex)
    
    daño = RandomNumber(hechizo.MinAgiFuer, hechizo.MaxAgiFuer)
      
    Call AddtoVar(victima.Stats.UserAtributos(eAtributos.Agilidad), daño, MAXATRIBUTOS)
    Call AddtoVar(victima.Stats.UserAtributos(eAtributos.Fuerza), daño, MAXATRIBUTOS)

    Call modPersonaje.incrementarAgilidad(victima, daño, 1200)
    Call modPersonaje.incrementarFuerza(victima, daño, 1200)

    b = True
ElseIf hechizo.AgiUpAndFuer = 2 Then
    If Not puedeAtacar(atacante, victima) Then Exit Sub
    
    If atacante.UserIndex <> victima.UserIndex Then
        Call UsuarioAtacadoPorUsuario(atacante, victima)
    End If
    
    Call InfoHechizo(atacante.UserIndex)

    daño = RandomNumber(hechizo.MinAgiFuer, hechizo.MaxAgiFuer)

    Call modPersonaje.reducirAgilidad(victima, daño, 700)
    Call modPersonaje.reducirFuerza(victima, daño, 700)
    
    b = True
End If

If hechizo.id = 38 Then

    If Not atacante.UserIndex = victima.UserIndex Then
        EnviarPaquete Paquetes.mensajeinfo, "Este hechizo solo te lo puedes lanzar a ti mismo.", atacante.UserIndex, ToIndex
        Exit Sub
    End If
    
    Call modPersonaje.incrementarAgilidad(victima, MAXATRIBUTOS, IntervaloDuracionPociones)
    Call modPersonaje.incrementarFuerza(victima, MAXATRIBUTOS, IntervaloDuracionPociones)
    
    Call InfoHechizo(atacante.UserIndex)
    
    b = True
End If

End Sub

Private Function getAbosrcionTotalRsistenciaMagica(ByRef personaje As User, daño As Integer) As Integer
    getAbosrcionTotalRsistenciaMagica = 0
    
    If personaje.Invent.CascoEqpObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.CascoEqpObjIndex), daño)
    End If
    
    If personaje.Invent.HerramientaEqpObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.HerramientaEqpObjIndex), daño)
    End If
    
    If personaje.Invent.AnilloEqpObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.AnilloEqpObjIndex), daño)
    End If
    
    If personaje.Invent.BrasaleteEqpObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.BrasaleteEqpObjIndex), daño)
    End If
    
    If personaje.Invent.ArmourEqpObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.ArmourEqpObjIndex), daño)
    End If
    
    If personaje.Invent.BarcoObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.BarcoObjIndex), daño)
    End If
    
    If personaje.Invent.EscudoEqpObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.EscudoEqpObjIndex), daño)
    End If
    
    If personaje.Invent.WeaponEqpObjIndex > 0 Then
        getAbosrcionTotalRsistenciaMagica = getAbosrcionTotalRsistenciaMagica + getResistenciaMagicaObjeto(ObjData(personaje.Invent.WeaponEqpObjIndex), daño)
    End If
    
End Function

Private Function getResistenciaMagicaObjeto(ByRef objeto As ObjData, danio As Integer) As Integer

    If objeto.DefensaMagicaMax > 0 Then
        getResistenciaMagicaObjeto = RandomNumber((danio * (objeto.DefensaMagicaMin / 100)), (danio * (objeto.DefensaMagicaMax / 100)))
    Else
        getResistenciaMagicaObjeto = 0
    End If

End Function
Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)
'Call LogTarea("Sub UpdateUserHechizos")
Dim loopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then
    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, slot, UserList(UserIndex).Stats.UserHechizos(slot))
    Else
        Call ChangeUserHechizo(UserIndex, slot, 0)
    End If
Else
'Actualiza todos los slots
For loopC = 1 To MAXUSERHECHIZOS
        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(loopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, loopC, UserList(UserIndex).Stats.UserHechizos(loopC))
        Else
            Call ChangeUserHechizo(UserIndex, loopC, 0)
        End If
Next loopC
End If
End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal hechizo As Integer)
'Call LogTarea("ChangeUserHechizo")
UserList(UserIndex).Stats.UserHechizos(slot) = hechizo
If hechizo > 0 And hechizo < NumeroHechizos + 1 Then
    EnviarPaquete CambiarHechizo, Chr$(slot) & Chr(hechizo) & hechizos(hechizo).nombre, UserIndex
Else
   EnviarPaquete CambiarHechizo, Chr$(slot), UserIndex
End If
End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub
Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(149), UserIndex
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(149), UserIndex
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        Call UpdateUserHechizos(False, UserIndex, CualHechizo + 1)
    End If
End If
Call UpdateUserHechizos(False, UserIndex, CualHechizo)
End Sub

Sub NpcLanzaSpellSobreNpc(ByVal npcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

Dim daño As Integer
Dim nIndex As Integer
Dim AnguloNPC As Single

If hechizos(Spell).SubeHP = 2 Then
        daño = RandomNumber(hechizos(Spell).minHP, hechizos(Spell).MaxHP)
        EnviarPaquete Paquetes.HechizoFX, ITS(NpcList(TargetNPC).Char.charIndex) & ByteToString(hechizos(Spell).FXgrh) & ITS(hechizos(Spell).loops) & Chr$(hechizos(Spell).WAV), TargetNPC, ToNPCArea, NpcList(TargetNPC).pos.map
        NpcList(TargetNPC).Stats.minHP = NpcList(TargetNPC).Stats.minHP - daño
        'Muere
        If NpcList(TargetNPC).Stats.minHP < 1 Then
            NpcList(TargetNPC).Stats.minHP = 0
            
            If NpcList(npcIndex).MaestroUser > 0 Then
                Call UsuarioMataNPC(UserList(NpcList(npcIndex).MaestroUser), NpcList(TargetNPC))
            End If
            
            nIndex = MuereNpc(NpcList(TargetNPC))
            
            If nIndex > 0 Then
                If DeboEnviarAngulo(UserList(NpcList(npcIndex).MaestroUser).pos.map) Then
                    AnguloNPC = Angulo(NpcList(nIndex).pos.x, NpcList(nIndex).pos.y, UserList(NpcList(npcIndex).MaestroUser).pos.x, UserList(NpcList(npcIndex).MaestroUser).pos.y)
                    EnviarPaquete Paquetes.AnguloNPC, ITS(AnguloNPC), NpcList(npcIndex).MaestroUser, ToIndex
                End If
            End If
            
        End If
          
End If
End Sub

'[Misery_Ezequiel 26/06/05]
Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim h As Integer
Dim TempX As Integer
Dim TempY As Integer

    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.hechizo)
    
    If hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If SV_PosicionesValidas.existePosicionMundo(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 _
                            And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            EnviarPaquete Paquetes.HechizoFX, ITS(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.charIndex) & ByteToString(hechizos(h).FXgrh) & ITS(hechizos(h).loops), UserIndex, ToPCArea, UserList(UserIndex).pos.map
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If
End Sub

'Marche nuevo
Private Function PuedeTirarIndomable(UserIndex As Integer) As Boolean
Dim i As Integer
Dim CantidadDeFuego As Integer
PuedeTirarIndomable = False
For i = 1 To MAXMASCOTAS
If UserList(UserIndex).MascotasIndex(i) > 0 Then
If UCase$(NpcList(UserList(UserIndex).MascotasIndex(i)).Name) = "ESPIRITU INDOMABLE" Then Exit Function
If UCase$(NpcList(UserList(UserIndex).MascotasIndex(i)).Name) = "FUEGO FATUO" Then Exit Function
If InStr(1, NpcList(UserList(UserIndex).MascotasIndex(i)).Name, "Elemental") > 0 Then CantidadDeFuego = CantidadDeFuego + 1
End If
Next
If CantidadDeFuego > 1 Then Exit Function
PuedeTirarIndomable = True
End Function

Private Function PuedeTirarImplorar(UserIndex As Integer) As Boolean
PuedeTirarImplorar = False
If UserList(UserIndex).NroMacotas = 0 Then PuedeTirarImplorar = True
End Function
Private Function PuedeTirarElementos(UserIndex As Integer) As Boolean
Dim i As Integer
Dim CantidadDeFuego As Integer
Dim tieneindomable As Boolean
PuedeTirarElementos = False
tieneindomable = False

For i = 1 To MAXMASCOTAS
If UserList(UserIndex).MascotasIndex(i) > 0 Then
If UCase$(NpcList(UserList(UserIndex).MascotasIndex(i)).Name) = "ESPIRITU INDOMABLE" Then tieneindomable = True
If UCase$(NpcList(UserList(UserIndex).MascotasIndex(i)).Name) = "FUEGO FATUO" Then Exit Function
If InStr(1, UCase$(NpcList(UserList(UserIndex).MascotasIndex(i)).Name), "FUEGO") > 0 Or InStr(1, UCase$(NpcList(UserList(UserIndex).MascotasIndex(i)).Name), "TIERRA") > 0 Then CantidadDeFuego = CantidadDeFuego + 1
End If
Next
                  
If tieneindomable And CantidadDeFuego = 1 Then Exit Function
PuedeTirarElementos = True
End Function
'marche nuevo

