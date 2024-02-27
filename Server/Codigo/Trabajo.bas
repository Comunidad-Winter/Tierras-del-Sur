Attribute VB_Name = "Trabajo"
Option Explicit


Public Sub DoNavega(ByRef personaje As User, ByRef Barco As ObjData, ByVal slot As Integer)

Dim ModNave As Long

' ¿Puede usar Barca?
If personaje.Stats.ELV < 25 Then
    If personaje.clase <> eClases.Pescador And personaje.clase <> eClases.Pirata Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(33), personaje.UserIndex
        Exit Sub
    End If
End If
            
ModNave = ModNavegacion(personaje.clase)

' ¿Puede usar esta barca?
If personaje.Stats.UserSkills(Navegacion) / ModNave < Barco.MinSkill Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(94), personaje.UserIndex
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(35) & (Barco.MinSkill * ModNave), personaje.UserIndex
    Exit Sub
End If

' ¿Eh?
If personaje.Invent.HerramientaEqpObjIndex = RED_PESCA Then Call Desequipar(personaje.UserIndex, personaje.Invent.HerramientaEqpSlot)

' Establecemos los datos
personaje.Invent.BarcoObjIndex = personaje.Invent.Object(slot).ObjIndex
personaje.Invent.BarcoSlot = slot

' Sino está navegando, arranca
If personaje.flags.Navegando = 0 Then

    ' ¿Hay agua cerca?
    If (MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger And eTriggers.Navegable) = 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes acercarte a aguas navegables.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    personaje.flags.Navegando = 1
Else

    If (MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger And eTriggers.NoCaminable) > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes acercarte a tierra firme.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
        
    personaje.flags.Navegando = 0
    
    Call WarpMascotas(personaje.UserIndex, True)
End If

' Le damos la apareciencia correspondiente
Call modPersonaje.DarAparienciaCorrespondiente(personaje)

' Actualizamos el char
Call modPersonaje_TCP.ActualizarEstetica(personaje)

' Toggle el Navegar
EnviarPaquete Paquetes.Navega, "", personaje.UserIndex
End Sub



Function quitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Long, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")
Dim i As Integer


With UserList(UserIndex)

    For i = 1 To .Stats.MaxItems
        If .Invent.Object(i).ObjIndex = ItemIndex Then
            'Si esta equipado lo desequipa
            If .Invent.Object(i).Equipped = 1 Then Call Desequipar(UserIndex, i)
            
            ' Si la cantidad es menor a lo que tengo en el slot
            If cant < .Invent.Object(i).Amount Then
                ' Se lo quito y la cnatidad a restar es 0
                .Invent.Object(i).Amount = .Invent.Object(i).Amount - cant
                cant = 0
            Else
                ' Si la cantidad es igual o mayor, el slot pasa a estar en 0 y resto la cantidad
                cant = cant - .Invent.Object(i).Amount
                ' Blanqueamos el slot
                .Invent.Object(i).Amount = 0
                .Invent.Object(i).ObjIndex = 0
            End If
            
            Call UpdateUserInv(False, UserIndex, i)
        
            If (cant = 0) Then
                quitarObjetos = True
                Exit Function
            End If
        End If
    Next i
    
End With

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer)

    Call quitarObjetosa(UserList(UserIndex), ObjData(ItemIndex).recursosNecesarios, cantidad)

End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer)

    Call quitarObjetosa(UserList(UserIndex), ObjData(ItemIndex).recursosNecesarios, cantidad)

End Sub

Public Sub quitarObjetosa(ByRef personaje As User, ByRef objetos() As ObjectoNecesario, Optional ByVal cantidad As Integer = 1)

    Dim loopObjeto As Integer
    
    For loopObjeto = LBound(objetos) To UBound(objetos)
    
        If objetos(loopObjeto).cantidad > 0 Then
        
            Call quitarObjetos(objetos(loopObjeto).ObjIndex, objetos(loopObjeto).cantidad * cantidad, personaje.UserIndex)
        
        End If
    Next loopObjeto

End Sub

Private Function tieneObjetosNecesariosConstruccion(personaje As User, objeto As ObjData, ByVal cantidad As Integer) As Boolean

    If objeto.recursosNecesarios(1).ObjIndex = 0 Then
        tieneObjetosNecesariosConstruccion = False
        Exit Function
    End If
    
    Dim loopItem As Byte
    
    For loopItem = LBound(objeto.recursosNecesarios) To UBound(objeto.recursosNecesarios)
        Dim Total As Long
        
        Total = objeto.recursosNecesarios(loopItem).cantidad * cantidad
    
        If Not TieneObjetos(objeto.recursosNecesarios(loopItem).ObjIndex, Total, personaje) Then
            tieneObjetosNecesariosConstruccion = False
            Exit Function
        End If
        
    Next
    
    tieneObjetosNecesariosConstruccion = True
End Function

Function CarpinteroTieneMateriales(personaje As User, objeto As ObjData, ByVal cantidad As Integer) As Boolean
    If objeto.recursosNecesarios(1).ObjIndex = 0 Then
        CarpinteroTieneMateriales = False
        Exit Function
    End If
    
    If tieneObjetosNecesariosConstruccion(personaje, objeto, cantidad) = False Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(96), personaje.UserIndex
        CarpinteroTieneMateriales = False
        Exit Function
    End If
        
    CarpinteroTieneMateriales = True
End Function
 
Function HerreroTieneMateriales(personaje As User, objeto As ObjData, ByVal cantidad As Integer) As Boolean
    Dim Total As Long
        
    If objeto.recursosNecesarios(1).ObjIndex = 0 Then
        HerreroTieneMateriales = False
        Exit Function
    End If
    
    If tieneObjetosNecesariosConstruccion(personaje, objeto, cantidad) = False Then
        EnviarPaquete Paquetes.mensajeinfo, "No tienes la cantidad suficiente de recursos para construir.", personaje.UserIndex
        HerreroTieneMateriales = False
        Exit Function
    End If
    
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruirObjetoMetales(personaje As User, ObjetoIndex As Integer) As Boolean

' ¿La cantidad de Skills en Herreria es suficiente para construir este objeto?
PuedeConstruirObjetoMetales = False

If Not PuedeConstruirHerreria(ObjetoIndex) Then
    EnviarPaquete Paquetes.mensajeinfo, "Este objeto no se puede construir a través de la herreria.", personaje.UserIndex
    Exit Function
End If

If Not personaje.Stats.UserSkills(eSkills.Herreria) >= ObjData(ObjetoIndex).SkHerreria * ModHerreriA(personaje.clase) Then
    EnviarPaquete Paquetes.mensajeinfo, "No tienes los suficientes conocimientos en Herreria para construir este objeto.", personaje.UserIndex
    Exit Function
End If

PuedeConstruirObjetoMetales = True
End Function

Public Function PuedeConstruirHerreria(ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i

For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i

PuedeConstruirHerreria = False
End Function

Public Sub DoHerreria(personaje As User)
Dim cantidad As Integer
Dim tieneEnergia As Boolean
Dim MiObj As obj
Dim ItemIndex As Integer
Dim mensaje As Byte

ItemIndex = personaje.Trabajo.modo

If personaje.clase = eClases.Herrero Then
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoTalarLeñador)
Else
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoTalarGeneral)
End If

If Not tieneEnergia Then
    ' Le avisamos que esta cansado
    EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado cansado. Esperá un poco antes de seguir trabajando.", personaje.UserIndex, ToIndex
    ' Dejamos de trabajar
    Call modPersonaje.DejarDeTrabajar(personaje)
    Exit Sub
End If

' ¿Cuantos hacemos?
' En TyTrabajoCant se guarda la cnatidad pendiente
' y en Suerte la cantidad que puede hacer al mismo tiempo
If personaje.Trabajo.cantidad - personaje.Trabajo.modificador < 0 Then
    cantidad = personaje.Trabajo.cantidad ' La cantidad restante
    personaje.Trabajo.cantidad = 0
Else
    cantidad = personaje.Trabajo.modificador
    personaje.Trabajo.cantidad = personaje.Trabajo.cantidad - personaje.Trabajo.modificador
End If

' Se puede construir? tiene los objetos y la capacidad?
If Not HerreroTieneMateriales(personaje, ObjData(ItemIndex), cantidad) Then
    Call DejarDeTrabajar(personaje)
    Exit Sub
End If

' Seteamos el Objeto
MiObj.Amount = cantidad
MiObj.ObjIndex = ItemIndex

' ¿Le entra?
If Not InvUsuario.tieneLugar(personaje, MiObj) Then
    EnviarPaquete Paquetes.mensajeinfo, "No tienes más lugar para guardar objetos.", personaje.UserIndex, ToIndex
    Call DejarDeTrabajar(personaje)
    Exit Sub
End If
    
' Quito minerales
Call HerreroQuitarMateriales(personaje.UserIndex, ItemIndex, cantidad)

' Le agrego los lingotes
Call MeterItemEnInventario(personaje.UserIndex, MiObj)

Call UpdateUserInv(True, personaje.UserIndex, 0)

Call SubirSkill(personaje.UserIndex, eSkills.Herreria)

' Envio el mensaje
If ObjData(ItemIndex).subTipo = OBJTYPE_WEAPON Then
    mensaje = 100
ElseIf ObjData(ItemIndex).subTipo = OBJTYPE_ESCUDO Then
    mensaje = 101
ElseIf ObjData(ItemIndex).subTipo = OBJTYPE_CASCO Then
    mensaje = 103
ElseIf ObjData(ItemIndex).subTipo = OBJTYPE_ARMADURA Then
    mensaje = 103
End If

If Not personaje.flags.UltimoMensaje = mensaje Then
    personaje.flags.UltimoMensaje = mensaje
    EnviarPaquete Paquetes.MensajeSimple, Chr$(mensaje), personaje.UserIndex
End If
          
EnviarPaquete Paquetes.WavSnd, Chr$(MARTILLOHERRERO), personaje.UserIndex
    
' ¿Termino de Trabajar?
If personaje.Trabajo.cantidad <= 0 Then Call DejarDeTrabajar(personaje)

End Sub

Public Function esObjetoConstruible_Carpinteria(ByVal ItemIndex As Integer) As Boolean

    Dim i As Long
    
    For i = 1 To UBound(ObjCarpintero)
        If ObjCarpintero(i) = ItemIndex Then
            esObjetoConstruible_Carpinteria = True
            Exit Function
        End If
    Next i
    
    esObjetoConstruible_Carpinteria = False

End Function
Public Function PuedeConstruirCarpintero(personaje As User, ByVal ItemIndex As Integer) As Boolean

PuedeConstruirCarpintero = False

' ¿Este objeto se puede construir?
If Not esObjetoConstruible_Carpinteria(ItemIndex) Then
    EnviarPaquete Paquetes.mensajeinfo, "No se puede construir este objeto a través de la Carpinteria.", personaje.UserIndex, ToIndex
    Exit Function
End If

' ¿Tiene los skils suficientes para construir este objeto?
If Not personaje.Stats.UserSkills(eSkills.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria * ModCarpinteria(personaje.clase) Then
    EnviarPaquete Paquetes.mensajeinfo, "No tienes la suficiente habilidad en Carpinteria para construir " & ObjData(ItemIndex).Name & ".", personaje.UserIndex, ToIndex
    Exit Function
End If
  
' ¿Tiene serrucho?
If Not personaje.Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then
    EnviarPaquete Paquetes.mensajeinfo, "Debes equipar el serrucho para construir un objeto de madera.", personaje.UserIndex, ToIndex
    Exit Function
End If

' ¿Es de tejo?
'If ObjData(ItemIndex).MaderaT > 0 Then
'    ' Se fija si puede laburar con eso cagada
'    If Not personaje.Stats.UserAtributos(eAtributos.Inteligencia) * personaje.Stats.UserSkills(eSkills.Magia) >= 525 Then
'        EnviarPaquete Paquetes.MensajeSimple2, Chr$(99), personaje.userIndex
'        Exit Function
'    End If
'End If

PuedeConstruirCarpintero = True

End Function

Public Sub DoCarpinteria(personaje As User)
Dim cantidad As Integer
Dim tieneEnergia As Boolean
Dim MiObj As obj
Dim ItemIndex As Integer

ItemIndex = personaje.Trabajo.modo

If personaje.clase = eClases.Herrero Then
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoTalarLeñador)
Else
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoTalarGeneral)
End If

If Not tieneEnergia Then
    ' Le avisamos que esta cansado
    EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado cansado. Esperá un poco antes de seguir trabajando.", personaje.UserIndex, ToIndex
    ' Dejamos de trabajar
    Call modPersonaje.DejarDeTrabajar(personaje)
    Exit Sub
End If

' Calculamos cuantos puede hacer, dependiendo la cantidad que falta hacer y cuantos
' puede hacer al mismo tiempo
If personaje.Trabajo.cantidad - personaje.Trabajo.modificador < 0 Then
    cantidad = personaje.Trabajo.cantidad
    personaje.Trabajo.cantidad = 0 ' Terminamos
Else
    cantidad = personaje.Trabajo.modificador
    personaje.Trabajo.cantidad = personaje.Trabajo.cantidad - personaje.Trabajo.modificador
End If

' ¿ Tiene los materiales?
If Not CarpinteroTieneMateriales(personaje, ObjData(ItemIndex), cantidad) Then
    Call modPersonaje.DejarDeTrabajar(personaje)
    Exit Sub
End If

MiObj.ObjIndex = ItemIndex
MiObj.Amount = cantidad

 ' ¿Le entra?
If Not InvUsuario.tieneLugar(personaje, MiObj) Then
    EnviarPaquete Paquetes.mensajeinfo, "No tienes más lugar para guardar objetos.", personaje.UserIndex, ToIndex
    Call DejarDeTrabajar(personaje)
    Exit Sub
End If

Call CarpinteroQuitarMateriales(personaje.UserIndex, ItemIndex, cantidad)

Call MeterItemEnInventario(personaje.UserIndex, MiObj)

Call UpdateUserInv(True, personaje.UserIndex, 0)

Call SubirSkill(personaje.UserIndex, eSkills.Carpinteria)
    
EnviarPaquete Paquetes.WavSnd, Chr$(LABUROCARPINTERO), personaje.UserIndex, ToPCArea

' ¿Terminamos?
If personaje.Trabajo.cantidad <= 0 Then Call DejarDeTrabajar(personaje)

End Sub


Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.hierrocrudo
            MineralesParaLingote = 13
        Case iMinerales.platacruda
            MineralesParaLingote = 25
        Case iMinerales.orocrudo
            MineralesParaLingote = 50
        Case Else
            MineralesParaLingote = 10001
    End Select
End Function

Public Sub DoFundirMineral(personaje As User)
Dim slot As Integer
Dim obji As Integer
Dim cantidad As Integer
Dim MiObj As obj

slot = personaje.Invent.HerramientaEqpSlot
obji = personaje.Invent.HerramientaEqpObjIndex
    
' ¿Tiene miinerales?
If personaje.Invent.Object(slot).Amount = 0 Then
    'No tiene nada en el slot
    EnviarPaquete Paquetes.MensajeSimple, Chr$(105), personaje.UserIndex
    
    personaje.Invent.HerramientaEqpSlot = 0
    personaje.Invent.HerramientaEqpObjIndex = 0
    
    Call UpdateUserInv(False, personaje.UserIndex, slot)
    Call DejarDeTrabajar(personaje)
    Exit Sub
End If

' Obtenemos la mayor catidad que puede hacer teniendo en cuenta la cantidad que tiene
' y su habilitad
cantidad = Int(personaje.Invent.Object(slot).Amount / MineralesParaLingote(obji))
cantidad = mini(cantidad, personaje.Trabajo.modificador * 2) 'TODO. X2 para acelerarlo.

' ¿Hacemos algo?
If cantidad = 0 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(105), personaje.UserIndex
    Call DejarDeTrabajar(personaje)
    Exit Sub
End If

'Creo el item
MiObj.Amount = cantidad
MiObj.ObjIndex = ObjData(obji).LingoteIndex

If Not InvUsuario.tieneLugar(personaje, MiObj) Then
    ' Avisamos
    EnviarPaquete Paquetes.mensajeinfo, "No tienes más lugar para guardar lingotes.", personaje.UserIndex, ToIndex
    ' Dejamos de trabajar
    Call DejarDeTrabajar(personaje)
    ' Salimos
    Exit Sub
End If

'Le quito los minerales
Call QuitarUserInvItem(personaje.UserIndex, slot, MineralesParaLingote(obji) * cantidad)
Call UpdateUserInv(False, personaje.UserIndex, slot)
    
' Agrego los lingotes
Call InvUsuario.MeterItemEnInventario(personaje.UserIndex, MiObj)

'Le mando un mensaje
If cantidad = 1 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(106), personaje.UserIndex
Else
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(20) & cantidad, personaje.UserIndex
End If
    
End Sub
Function ModNavegacion(ByVal clase As eClases) As Integer

Select Case clase
    Case eClases.Pirata
        ModNavegacion = 1
    Case eClases.Pescador
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function

Function ModFundicion(ByVal clase As eClases) As Integer

Select Case clase
    Case eClases.Minero
        ModFundicion = 1
    Case eClases.Herrero
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal clase As eClases) As Integer

Select Case clase
    Case eClases.Carpintero
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal clase As eClases) As Integer
Select Case clase
    Case eClases.Herrero
        ModHerreriA = 1
    Case eClases.Minero
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select
End Function

Function ModDomar(ByVal clase As eClases) As Integer

Select Case clase
    Case eClases.Druida
        ModDomar = 6
    Case Else
        ModDomar = 11
End Select

End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long
    CalcularPoderDomador = (UserList(UserIndex).Stats.UserSkills(Domar) * UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) / ModDomar(UserList(UserIndex).clase)
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
'Call LogTarea("Sub FreeMascotaIndex")
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) = 0 Then
        FreeMascotaIndex = j
        Exit Function
    End If
Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal npcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then

    If NpcList(npcIndex).MaestroUser = UserIndex Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(108), UserIndex
        Exit Sub
    End If
    
    If NpcList(npcIndex).MaestroNpc > 0 Or NpcList(npcIndex).MaestroUser > 0 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(109), UserIndex
        Exit Sub
    End If
    
    If NpcList(npcIndex).flags.Domable <= CalcularPoderDomador(UserIndex) Then
           
        If Int(RandomNumber(0, 3)) = 2 Then
        
            'Se doma
            Call NPCs.establecerAmo(UserIndex, npcIndex)
            'Hago que los npcs siguen al su amo
            Call FollowAmo(npcIndex)
            'Como este lo dome, pongo un nuevo npc.
            Call CrearNPC(NpcList(npcIndex).numero, UserList(UserIndex).pos, NpcList(npcIndex).Orig, False)
            
            EnviarPaquete Paquetes.MensajeSimple, Chr$(110), UserIndex
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr$(111), UserIndex
        End If
        
    Else
    
         'Para no estar repitiendo siempre lo mismo
          If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(111), UserIndex
            UserList(UserIndex).flags.UltimoMensaje = 5
          End If
          
    End If
    
    'Dome o no sube skils
    Call SubirSkill(UserIndex, eSkills.Domar)
    
Else
    EnviarPaquete Paquetes.MensajeSimple, Chr$(112), UserIndex
End If


End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.Body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Body = 0
        UserList(UserIndex).Char.Head = 0
    Else
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).Char.Body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
    End If
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
Dim Suerte As Byte
Dim exito As Byte
Dim obj As obj

If MapInfo(UserList(UserIndex).pos.map).Pk = False Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(86), UserIndex
    Exit Sub
End If

If Not esPosicionJugable(x, y) Then Exit Sub

If MapData(map, x, y).OBJInfo.Amount < 3 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(113), UserIndex
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 34 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 35 Then
            Suerte = 1
End If
exito = RandomNumber(1, Suerte)
If exito = 1 Then
    obj.ObjIndex = FOGATA_APAG
    obj.Amount = MapData(map, x, y).OBJInfo.Amount / 3
    If obj.Amount > 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "Has hecho " & obj.Amount & " fogatas.", UserIndex
    Else
        EnviarPaquete Paquetes.MensajeSimple, Chr$(114), UserIndex
    End If
    Call MakeObj(ToMap, 0, map, obj, map, x, y)

Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(115), UserIndex
        UserList(UserIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If
Call SubirSkill(UserIndex, Supervivencia)
End Sub


Public Function DoApuñalar(ByRef atacante As User, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer) As Integer
Dim Suerte As Integer
Dim res As Integer

If atacante.clase = eClases.asesino Then
    Suerte = (atacante.Stats.UserSkills(Apuñalar) \ 10) * 2
ElseIf atacante.clase = eClases.Guerrero Then
    Suerte = (atacante.Stats.UserSkills(Apuñalar) \ 10) * 0.5
Else
    Suerte = (atacante.Stats.UserSkills(Apuñalar) \ 10) * 1
End If

res = RandomNumber(0, 100)

If res > Suerte Then
    'EnviarPaquete Paquetes.MensajeSimple, Chr$(123), atacante.userIndex
    Exit Function
End If

If VictimUserIndex <> 0 Then
    DoApuñalar = Int(daño * 1.5)
Else
    NpcList(VictimNpcIndex).Stats.minHP = NpcList(VictimNpcIndex).Stats.minHP - Int(daño * 2)
    
    DoApuñalar = Int(daño * 2) + daño

    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(6) & Int(daño * 2) & "," & Int(daño * 3), atacante.UserIndex
    Call SubirSkill(atacante.UserIndex, Apuñalar)
    Call CalcularDarExp(atacante, NpcList(VictimNpcIndex), Int(daño * 2))
End If

End Function


Public Sub DoMeditar(ByRef personaje As User, ByVal tiempoTranscurrido As Long)

Dim cant As Integer
Dim ticks As Byte
Dim intervaloMeditacion As Integer

' Si tiene maxima mana, deja de laburar.
If personaje.Stats.MinMAN >= personaje.Stats.MaxMAN Then
    EnviarPaquete Paquetes.Meditando, "", personaje.UserIndex
    personaje.flags.Meditando = False
    personaje.Char.FX = 0
    personaje.Char.loops = 0
    EnviarPaquete Paquetes.HechizoFX, ITS(personaje.Char.charIndex) & ByteToString(0) & ITS(0), personaje.UserIndex, ToMap, personaje.pos.map
    Exit Sub
End If

' Calculamos el tiempo cada cuanto medita el usuario
' Para todos es el mismo tiempo de meditacion
intervaloMeditacion = 500

personaje.Counters.Meditacion = personaje.Counters.Meditacion + tiempoTranscurrido
 
' El tiempo que paso es menor?
If personaje.Counters.Meditacion >= intervaloMeditacion Then
    ' Por las dudas de que haya un lagaso contamos cuantas veces le tendría que haber subido la meditacion
    ticks = personaje.Counters.Meditacion / intervaloMeditacion
    
    ' Le resto el tiempo que paso
    personaje.Counters.Meditacion = personaje.Counters.Meditacion - (intervaloMeditacion * ticks)

    If personaje.Stats.UserSkills(eSkills.Meditar) <= 90 Then
        cant = (personaje.Stats.MaxMAN / 33) * ticks
    ElseIf personaje.Stats.UserSkills(eSkills.Meditar) <= 99 Then
       cant = (personaje.Stats.MaxMAN / 25) * ticks
    Else
       cant = (personaje.Stats.MaxMAN / 20) * ticks
    End If
    
    Call AddtoVar(personaje.Stats.MinMAN, cant, personaje.Stats.MaxMAN)
    
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(19) & cant, personaje.UserIndex, ToIndex
    Call SendUserMana(personaje.UserIndex)
    Call SubirSkill(personaje.UserIndex, eSkills.Meditar)
End If

End Sub

Sub VolverCriminal2(ByVal UserIndex As Integer)
If MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Trigger = 6 Then Exit Sub
If UserList(UserIndex).flags.Privilegios < 2 Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    UserList(UserIndex).faccion.ArmadaReal = 1
    UserList(UserIndex).faccion.CiudadanosMatados = 1
    Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    If UserList(UserIndex).faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If
End Sub

Sub VolverCiudadano2(ByVal UserIndex As Integer)
If MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Trigger = 6 Then Exit Sub
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).faccion.CiudadanosMatados = 0
Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlASALTO, MAXREP)
End Sub



Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Wresterling) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) <= 99 _
   And UserList(UserIndex).Stats.UserSkills(Wresterling) >= 91 Then
                    Suerte = 8
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(91), UserIndex
        EnviarPaquete Paquetes.MensajeFight, "Tu oponente te ha desarmado!", VictimIndex
End If
End Sub


Public Sub personajeHerreria(personaje As User, datos As String)
    
    Dim ObjetoIndex As Integer
    Dim cantidad As Integer

    ' ¿Ya esta trabajando?
    If personaje.flags.Trabajando = True Then Exit Sub
    
    ObjetoIndex = DeCodify(Right(datos, Len(datos) - 2))
    cantidad = STI(datos, 1) 'cantidad a hacer
        
    If ObjetoIndex <= 0 Or ObjetoIndex > UBound(ObjData) Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes seleccionar un objeto para construir.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
   
    If cantidad <= 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes seleccionar la cantidad de objetos de este tipo que deseas construir.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' ¿Este objeto lo puede contruir el usuario?
    If Not PuedeConstruirObjetoMetales(personaje, ObjetoIndex) Then
        Call DejarDeTrabajar(personaje)
        Exit Sub
    End If

    ' Configuramos el trabajo
    personaje.Trabajo.tipo = eTrabajos.Herreria
    personaje.Trabajo.modo = ObjetoIndex
    personaje.Trabajo.cantidad = cantidad
    Call Trabajo.CalcularModificador(personaje)
    
    ' Agrego
    personaje.flags.Trabajando = True
    Call TrabajadoresGroup.agregar(personaje.UserIndex)
    
    ' Informo
    EnviarPaquete Paquetes.EmpiezaTrabajo, "", personaje.UserIndex, ToIndex
End Sub

Public Sub personajeCarpinteria(personaje As User, datos As String)
    Dim ObjetoIndex As Long
    Dim cantidad As Integer
    
    ' ¿Ya esta trabajando?
    If personaje.flags.Trabajando Then Exit Sub
    
    ObjetoIndex = DeCodify(Right(datos, Len(datos) - 2)) ' objeto que desea hacer
    cantidad = STI(datos, 1) 'cantidad a hacer
    
    If ObjetoIndex <= 0 Or ObjetoIndex > UBound(ObjData) Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes seleccionar un objeto para construir.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
   
    If cantidad <= 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes seleccionar la cantidad de objetos de este tipo que deseas construir.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' ¿Es un Objeto construible? ¿Y lo puede construir el usuairo?
    If Not PuedeConstruirCarpintero(personaje, ObjetoIndex) Then
        Exit Sub
    End If

    ' Seteamos los parametros
    personaje.Trabajo.tipo = eTrabajos.Carpinteria
    personaje.Trabajo.modo = ObjetoIndex
    personaje.Trabajo.cantidad = cantidad
    
    personaje.flags.Trabajando = True
      
    '
    Call Trabajo.CalcularModificador(personaje)
    
    Call TrabajadoresGroup.agregar(personaje.UserIndex)
    EnviarPaquete Paquetes.EmpiezaTrabajo, "", personaje.UserIndex, ToIndex
End Sub

       
Public Sub personajeTrabajar(personaje As User)
        
    Dim auxind As Integer
    Dim wpaux As WorldPos
                
    ' ¿Ya esta trabajando?
    If personaje.flags.Trabajando Then Call DejarDeTrabajar(personaje)
        
    ' ¿Tiene alguna herramienta?
    If Not personaje.Invent.HerramientaEqpObjIndex > 0 Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(130), personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' ¿Tiene energia?
    If Not personaje.Stats.MinSta > 0 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(167), personaje.UserIndex, ToIndex
        Exit Sub
    End If
        
    ' Es Newbie el objeto pero el usuario no lo es?
    If ObjData(personaje.Invent.HerramientaEqpObjIndex).Newbie = 1 And Not EsNewbie(personaje.UserIndex) Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(287 - 255), personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' Dependiendo la herramenta es lo que quiere hacer
    Select Case personaje.Invent.HerramientaEqpObjIndex

    Case OBJTYPE_CAÑA, RED_PESCA
        
        ' No se puede trabajar bajo techo
        If (MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger And eTriggers.BajoTecho) Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(234), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' Distancia al agua
        wpaux.map = personaje.pos.map
        wpaux.x = personaje.flags.TargetX
        wpaux.y = personaje.flags.TargetY
                    
        If distancia(wpaux, personaje.pos) > 5 Then
            EnviarPaquete Paquetes.mensajeinfo, "Tu herramienta no es lo suficientemente larga para llegar al agua.", personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' ¿hay agua?
        If Not HayAgua(personaje.pos.map, personaje.flags.TargetX, personaje.flags.TargetY) Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(235), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' Cargamos los modificadores
        personaje.Trabajo.tipo = eTrabajos.Pesca
        
        If personaje.Invent.HerramientaEqpObjIndex = OBJTYPE_CAÑA Then
            personaje.Trabajo.modo = OBJTYPE_CAÑA
        Else
            personaje.Trabajo.modo = RED_PESCA
        End If
        
        Call Trabajo.CalcularModificador(personaje)
        
        ' Lo agregamos en la lista de trabajadores y lo marcamos como trabajando
        Call TrabajadoresGroup.agregar(personaje.UserIndex)
        personaje.flags.Trabajando = True
        
        EnviarPaquete Paquetes.EmpiezaTrabajo, "", personaje.UserIndex, ToIndex
    
    Case HACHA_LEÑADOR, HACHA_DORADA
        
        ' ¿Esta cliceando un objeto sobre una posicion valida?
        If Not SV_PosicionesValidas.existePosicionMundo(personaje.pos.map, personaje.flags.TargetObjX, personaje.flags.TargetObjY) Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(241), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' ¿Hay algo en el suelo?
        auxind = MapData(personaje.pos.map, personaje.flags.TargetObjX, personaje.flags.TargetObjY).OBJInfo.ObjIndex
        
        If auxind = 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(241), personaje.UserIndex, ToIndex
            Exit Sub
        End If
                
        ' Distancia a lo que esta clickeando
        wpaux.map = personaje.pos.map
        wpaux.x = personaje.flags.TargetObjX
        wpaux.y = personaje.flags.TargetObjY
        
        If distancia(wpaux, personaje.pos) > 2 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(5), personaje.UserIndex, ToIndex
            Exit Sub
        ElseIf distancia(wpaux, personaje.pos) = 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(240), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' ¿Es un Arbol?
        If Not ObjData(auxind).ObjType = OBJTYPE_ARBOLES Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes extraer leña de ahi.", personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        'Si talar en un arbol de tejo tenes que tener hacha dorada
        If auxind = Objetos_Constantes.ARBOL_DE_TEJO And Not personaje.Invent.HerramientaEqpObjIndex = HACHA_DORADA Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(278 - 255), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' Si no talas tejo tenes que tener un hacha comun
        If Not auxind = Objetos_Constantes.ARBOL_DE_TEJO And Not personaje.Invent.HerramientaEqpObjIndex = HACHA_LEÑADOR Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(279 - 255), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        If ObjData(auxind).LeñaIndex = 0 Then
            EnviarPaquete Paquetes.mensajeinfo, "Este ärbol no es talable", personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' Establecemos los parametros del trabajo que llevar adelante el usuario
        personaje.Trabajo.tipo = eTrabajos.Tala
        personaje.Trabajo.modo = ObjData(auxind).LeñaIndex
        
        Call Trabajo.CalcularModificador(personaje)
        
        ' Agregamos al trabajador
        TrabajadoresGroup.agregar (personaje.UserIndex)
        personaje.flags.Trabajando = True
        
        ' Le avisamos
        EnviarPaquete Paquetes.EmpiezaTrabajo, "", personaje.UserIndex, ToIndex

    Case PIQUETE_MINERO, PIQUETE_DE_ORO
        
        ' ¿Esta cliceando un objeto sobre una posicion valida?
        If Not SV_PosicionesValidas.existePosicionMundo(personaje.pos.map, personaje.flags.TargetObjX, personaje.flags.TargetObjY) Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(242), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' ¿Esta seleccionando un yacimiento?
        auxind = MapData(personaje.pos.map, personaje.flags.TargetObjX, personaje.flags.TargetObjY).OBJInfo.ObjIndex
        
        If auxind = 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(242), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        If Not ObjData(auxind).ObjType = OBJTYPE_YACIMIENTO Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(242), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' Distancia al elemento seleccionado
        wpaux.map = personaje.pos.map
        wpaux.x = personaje.flags.TargetObjX
        wpaux.y = personaje.flags.TargetObjY
                    
        If distancia(wpaux, personaje.pos) > 2 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(5), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' Para minar oro solo con piquete de oro
        If ObjData(auxind).MineralIndex = iMinerales.orocrudo And Not personaje.Invent.HerramientaEqpObjIndex = PIQUETE_DE_ORO Then
            EnviarPaquete Paquetes.mensajeinfo, "Para minar oro necesitas un piquete de oro", personaje.UserIndex
            Exit Sub
        End If
           
        ' Cargo datos
        personaje.Trabajo.tipo = eTrabajos.Mineria
        personaje.Trabajo.modo = ObjData(auxind).MineralIndex
        
        Call Trabajo.CalcularModificador(personaje)
                        
        Call TrabajadoresGroup.agregar(personaje.UserIndex)
        personaje.flags.Trabajando = True
            
        EnviarPaquete Paquetes.EmpiezaTrabajo, "", personaje.UserIndex, ToIndex
            
    Case hierrocrudo, orocrudo, platacruda ' Esto es Fundicion
    
        ' ¿Esta cliceando un objeto sobre una posicion valida?
        If Not SV_PosicionesValidas.existePosicionMundo(personaje.pos.map, personaje.flags.TargetObjX, personaje.flags.TargetObjY) Then
            EnviarPaquete Paquetes.MensajeSimple, Chr(247), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' ¿Hay algo?
        auxind = MapData(personaje.pos.map, personaje.flags.TargetObjX, personaje.flags.TargetObjY).OBJInfo.ObjIndex
        
        If auxind = 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr(247), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' ¿Es una fragua?
        If Not ObjData(auxind).ObjType = OBJTYPE_FRAGUA Then
            EnviarPaquete Paquetes.MensajeSimple, Chr(247), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' Distancia al elemento seleccionado
        wpaux.map = personaje.pos.map
        wpaux.x = personaje.flags.TargetObjX
        wpaux.y = personaje.flags.TargetObjY
                    
        If distancia(wpaux, personaje.pos) > 5 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(5), personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        ' ¿Tengo los Skills para manipularlo?
        If Not (personaje.Stats.UserSkills(eSkills.Mineria) / ModFundicion(personaje.clase) > ObjData(personaje.Invent.HerramientaEqpObjIndex).MinSkill) Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(95), personaje.UserIndex
            Exit Sub
        End If
        
        ' Seteamos
        personaje.Trabajo.tipo = eTrabajos.Fundicion
        personaje.Trabajo.modo = 0
        
        Call Trabajo.CalcularModificador(personaje)
                
        ' Agregamos a la lista de trabajadores
        TrabajadoresGroup.agregar (personaje.UserIndex)
        personaje.flags.Trabajando = True
        
        EnviarPaquete Paquetes.EmpiezaTrabajo, "", personaje.UserIndex, ToIndex
    
    End Select
    
End Sub


Public Sub CalcularModificador(personaje As User)
Dim Suerte As Integer

Select Case personaje.Trabajo.tipo

Case eTrabajos.Pesca  'PESCAR
        
    Select Case personaje.Stats.UserSkills(eSkills.Pesca)
        Case 0:         Suerte = 200
        Case 1 To 10:   Suerte = 195
        Case 11 To 20:  Suerte = 190
        Case 21 To 30:  Suerte = 180
        Case 31 To 40:  Suerte = 170
        Case 41 To 50:  Suerte = 160
        Case 51 To 60:  Suerte = 150
        Case 61 To 70:  Suerte = 140
        Case 71 To 80:  Suerte = 130
        Case 81 To 90:  Suerte = 120
        Case 91 To 99:  Suerte = 110
        Case Else:      Suerte = 100
    End Select
    
    personaje.Trabajo.modificador = Suerte
    
Case eTrabajos.Tala  'TALAR
    
    personaje.Trabajo.modificador = modTalar.calcularModificadorTalar(personaje)
    personaje.Trabajo.rangoGeneracion = modTalar.calcularRangoExtraccionTalar(personaje)
    
Case eTrabajos.Mineria  'MINAR
    
    personaje.Trabajo.modificador = modMineria.calcularModificadorMineria(personaje)
    personaje.Trabajo.rangoGeneracion = modMineria.calcularRangoExtraccionMineria(personaje)

Case eTrabajos.Fundicion  'LINGOTEAR
        
        If personaje.Stats.UserSkills(eSkills.Mineria) <= 25 Then
            Suerte = 1
        ElseIf personaje.Stats.UserSkills(eSkills.Mineria) <= 50 Then
            Suerte = 2
        ElseIf personaje.Stats.UserSkills(eSkills.Mineria) <= 75 Then
            Suerte = 3
        Else
            Suerte = 4
        End If
        
        personaje.Trabajo.modificador = Suerte
        
Case eTrabajos.Carpinteria, eTrabajos.Herreria   'HERRERIA y CARPINTERIA

        If personaje.Stats.ELV <= 5 Then
            Suerte = 1
        ElseIf personaje.Stats.ELV < 14 Then
            Suerte = 2
        ElseIf personaje.Stats.ELV < 24 Then
            Suerte = 3
        Else
            Suerte = 4
        End If
        
        personaje.Trabajo.modificador = Suerte * 2
End Select



End Sub
