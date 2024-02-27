Attribute VB_Name = "SistemaCombate"
Option Explicit

Public Const MAXDISTANCIAARCO = 18
Public Const MAXDISTANCIAMAGIA = 18



Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal npcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1
If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(UserList(UserIndex))
    Else
        PoderAtaque = PoderAtaqueArma(UserList(UserIndex))
    End If
Else 'Peleando con puños
    PoderAtaque = PoderAtaqueWresterling(UserIndex)
End If
ProbExito = HelperMatematicas.maxs(10, HelperMatematicas.mins(90, 50 + ((PoderAtaque - NpcList(npcIndex).PoderEvasion) * 0.4)))
UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(UserIndex, proyectiles)
       Else
            Call SubirSkill(UserIndex, Armas)
       End If
    Else
        Call SubirSkill(UserIndex, Wresterling)
    End If
End If
End Function

Public Function NpcImpacto(ByVal npcIndex As Integer, ByVal UserIndex As Integer) As Boolean
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(UserList(UserIndex))
NpcPoderAtaque = NpcList(npcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(UserList(UserIndex))
SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkills.tacticas)
SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkills.Defensa)
'Esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
ProbExito = HelperMatematicas.maxs(10, HelperMatematicas.mins(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
' el usuario esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
   If NpcImpacto = False Then
      ProbRechazo = HelperMatematicas.maxs(10, HelperMatematicas.mins(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas + 1))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
         EnviarPaquete Paquetes.WavSnd, Chr$(SND_ESCUDO), UserIndex, ToPCArea
         EnviarPaquete Paquetes.COMBRechEsc, "", UserIndex
         EnviarPaquete Paquetes.AnimEscu, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea
         Call SubirSkill(UserIndex, Defensa)
      End If
   End If
End If
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal npcIndex As Integer = 0) As Long
Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim DañoMaxArma As Long
Dim DañoExtraCazador As Integer

DañoExtraCazador = 0

'¿Tiene un arma?
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    ' Ataca a un npc?
    If npcIndex > 0 Then
        'Usa la mata dragones?
        If Arma.subTipo = MATADRAGONES Then ' Usa la matadragones?
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).ClaseNumero)
                If NpcList(npcIndex).Name = "Dragon rojo" Or NpcList(npcIndex).Name = "Gran Dragón Rojo" Then  'Ataca dragon?
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                UserList(UserIndex).Stats.MinSta = 0
            Else ' Sino es dragon daño es 0
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(92), UserIndex
                DañoArma = 0
                DañoMaxArma = 0
                Exit Function
            End If
        Else
            '¿Es un arma que lanza proyectiles?
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).ClaseNumero)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                DañoExtraCazador = DañoExtra(UserIndex)
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                'Es un arma de combate cuerpo a cuerpo
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).ClaseNumero)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    Else ' Ataca usuario
        If Arma.subTipo = MATADRAGONES Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(92), UserIndex
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).ClaseNumero)
            DañoArma = 0 ' Si usa la espada matadragones daño es 0
            DañoMaxArma = 0
            Exit Function
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).ClaseNumero)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                    'Dañoextracazador =
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).ClaseNumero)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    End If
End If

DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHIT)

CalcularDaño = (((3 * DañoArma) + ((DañoMaxArma / 5) * HelperMatematicas.maxs(0, (UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + DañoUsuario) * ModifClase) + DañoExtraCazador

End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal npcIndex As Integer)
Dim daño As Long
Dim nIndex As Integer
Dim AnguloNPC As Single

daño = CalcularDaño(UserIndex, npcIndex)

'esta navegando? si es asi le sumamos el daño del barco
If UserList(UserIndex).flags.Navegando = 1 Then _
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)

daño = daño - NpcList(npcIndex).Stats.Def

If daño < 0 Then daño = 0

If daño = 0 Then
    EnviarPaquete Paquetes.MensajeFight, "No lográs causarle daño a la criatura", UserIndex, ToIndex
    Exit Sub
End If

NpcList(npcIndex).Stats.minHP = NpcList(npcIndex).Stats.minHP - daño
EnviarPaquete Paquetes.COMBUserImpcNpc, Codify(daño), UserIndex

'q feo esto
If NpcList(npcIndex).Stats.minHP > 0 Then
    If PuedeApuñalar(UserIndex) Then
       Dim danioApu As Integer
       
       danioApu = DoApuñalar(UserList(UserIndex), npcIndex, 0, daño)
       
       Call SubirSkill(UserIndex, Apuñalar)
       If danioApu > 0 Then
            EnviarPaquete Paquetes.TXAII, ITS(NpcList(npcIndex).pos.x) & ITS(NpcList(npcIndex).pos.y) & ITS(danioApu) & ITS(distancia(NpcList(npcIndex).pos, UserList(UserIndex).pos)), UserIndex, ToIndex, NpcList(npcIndex).pos.map
       Else
            EnviarPaquete Paquetes.TXA, ITS(NpcList(npcIndex).pos.x) & ITS(NpcList(npcIndex).pos.y) & ITS(daño) & ITS(distancia(NpcList(npcIndex).pos, UserList(UserIndex).pos)), UserIndex, ToIndex, NpcList(npcIndex).pos.map
       End If
    Else
        EnviarPaquete Paquetes.TXA, ITS(NpcList(npcIndex).pos.x) & ITS(NpcList(npcIndex).pos.y) & ITS(daño) & ITS(distancia(NpcList(npcIndex).pos, UserList(UserIndex).pos)), UserIndex, ToIndex, NpcList(npcIndex).pos.map
    End If
Else
    If daño < 32000 Then EnviarPaquete Paquetes.TXA, ITS(NpcList(npcIndex).pos.x) & ITS(NpcList(npcIndex).pos.y) & ITS(daño) & ITS(distancia(NpcList(npcIndex).pos, UserList(UserIndex).pos)), UserIndex, ToIndex, NpcList(npcIndex).pos.map
End If

Call CalcularDarExp(UserList(UserIndex), NpcList(npcIndex), daño)

If NpcList(npcIndex).Stats.minHP <= 0 Then
    ' Si era un Dragon perdemos la espada matadragones
    If NpcList(npcIndex).NPCtype = DRAGON Then
        Call quitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
    End If
    ' Para que las mascotas no sigan intentando luchar y
    ' comiencen a seguir al amo
    Call UsuarioMataNPC(UserList(UserIndex), NpcList(npcIndex))
    
    nIndex = MuereNpc(NpcList(npcIndex))
    
    If nIndex > 0 Then
        If DeboEnviarAngulo(UserList(UserIndex).pos.map) Then
            AnguloNPC = Angulo(NpcList(nIndex).pos.x, NpcList(nIndex).pos.y, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)
            EnviarPaquete Paquetes.AnguloNPC, ITS(AnguloNPC), UserIndex, ToIndex
        End If
    End If
    
End If

End Sub

Public Sub NpcDaño(ByVal npcIndex As Integer, ByVal UserIndex As Integer)
Dim daño As Integer, Lugar As Integer, absorbido As Integer
Dim antdaño As Integer, defbarco As Integer
Dim obj As ObjData

daño = RandomNumber(NpcList(npcIndex).Stats.MinHIT, NpcList(npcIndex).Stats.MaxHIT)
antdaño = daño
If UserList(UserIndex).flags.Navegando = 1 Then
    obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(obj.MinDef, obj.MaxDef)
End If
Lugar = RandomNumber(1, 6)
Select Case Lugar
  Case bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
           obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
           obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
           absorbido = absorbido + defbarco
         End If
         
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        obj = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
        absorbido = absorbido + RandomNumber(obj.MinDef, obj.MaxDef)
        End If
        
        daño = daño - absorbido
        If daño < 1 Then daño = 1
        
End Select

EnviarPaquete Paquetes.COMBNpcHIT, Chr$(Lugar) & Codify(daño), UserIndex, ToIndex
EnviarPaquete Paquetes.TXA, ITS(UserList(UserIndex).pos.x) & ITS(UserList(UserIndex).pos.y) & ITS(daño) & ITS(distancia(UserList(UserIndex).pos, NpcList(npcIndex).pos)), UserIndex, ToIndex, UserList(UserIndex).pos.map

If UserList(UserIndex).flags.Privilegios = 0 Then UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.minHP - daño

'Muere el usuario
If UserList(UserIndex).Stats.minHP <= 0 Then
    EnviarPaquete Paquetes.COMBMuereUser, "", UserIndex

    If NpcList(npcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(NpcList(npcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If NpcList(npcIndex).Hostil = False Then
            Call RestoreOldMovement(npcIndex)
        End If
    End If
    Call UserDie(UserIndex, False)
End If
End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)
    'If UserList(UserIndex).Reputacion.AsesinoRep > 0 Then
    '     UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep - vlASESINO
    '     If UserList(UserIndex).Reputacion.AsesinoRep < 0 Then UserList(UserIndex).Reputacion.AsesinoRep = 0
    'Else
    If UserList(UserIndex).Reputacion.BandidoRep > 0 Then
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep - vlASALTO
        If UserList(UserIndex).Reputacion.BandidoRep <= 0 Then
        UserList(UserIndex).Reputacion.BandidoRep = 0
        End If
    End If

     If UserList(UserIndex).Reputacion.LadronesRep > 0 Then
            UserList(UserIndex).Reputacion.LadronesRep = UserList(UserIndex).Reputacion.LadronesRep - (vlCAZADOR * 10)
            If UserList(UserIndex).Reputacion.LadronesRep < 0 Then
             UserList(UserIndex).Reputacion.LadronesRep = 0
            End If
    End If
End Sub

Public Sub CheckPets(ByVal npcIndex As Integer, ByVal UserIndex As Integer)
'Anti robo de npcs
Dim otroUsuario As Integer

If NpcList(npcIndex).MaestroUser = 0 And MapInfo(NpcList(npcIndex).pos.map).PermiteRoboNPC = 0 Then
  otroUsuario = estaLuchando(NpcList(npcIndex))

  If Not otroUsuario = UserIndex And otroUsuario > 0 Then
      If Not AntiRoboNpc.puedePegarleAlNpc(UserIndex, otroUsuario) Then
          EnviarPaquete Paquetes.mensajeinfo, "Tu mascotas no pueden atacar a esta criatura por que esta está peleando con " & UserList(otroUsuario).Name, UserIndex, ToIndex
          Exit Sub
      End If
  Else
      If UserList(UserIndex).LuchandoNPC <> npcIndex And UserList(UserIndex).LuchandoNPC > 0 Then
          ' Si antes le estaba pegando a otro npc, libero a ese npc
          Call AntiRoboNpc.resetearLuchador(NpcList(UserList(UserIndex).LuchandoNPC))
      End If
  NpcList(npcIndex).UltimoGolpe = GetTickCount()
  NpcList(npcIndex).UserIndexLucha = UserIndex
  UserList(UserIndex).LuchandoNPC = npcIndex
  End If
End If

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
        If UserList(UserIndex).MascotasIndex(j) <> npcIndex Then
            'Balance. Si es el ele de fuego de tierra no ataca npcs, esto se podria hacer con una variable desde los dats, pero me parece q es medio al pedo total se toma solo en este sub.
            If Not (NpcList(UserList(UserIndex).MascotasIndex(j)).numero = ELEMENTALFUEGO Or NpcList(UserList(UserIndex).MascotasIndex(j)).numero = ELEMENTALTIERRA) Then
                  NpcList(UserList(UserIndex).MascotasIndex(j)).TargetNPCID = npcIndex
                  NpcList(UserList(UserIndex).MascotasIndex(j)).Movement = NPCDEFENSA
            End If
        End If
    End If
Next j
Exit Sub
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
    End If
Next j
End Sub

Public Sub NpcAtacaUser(ByVal npcIndex As Integer, ByVal UserIndex As Integer)

If UserList(UserIndex).flags.Mimetizado = 0 Then
    Call CheckPets(npcIndex, UserIndex)
    'If NpcList(NpcIndex).TargetUserID = 0 Then NpcList(NpcIndex).TargetUserID = UserList(UserIndex).id
Else
    Exit Sub
End If

If NpcList(npcIndex).flags.Snd1 > 0 Then EnviarPaquete Paquetes.WavSnd, Chr$(NpcList(npcIndex).flags.Snd1), UserIndex, ToPCArea, UserList(UserIndex).pos.map
If NpcImpacto(npcIndex, UserIndex) Then
    EnviarPaquete Paquetes.WavSnd, Chr$(SND_IMPACTO), UserIndex, ToPCArea
    If UserList(UserIndex).flags.Meditando = False Then
        If UserList(UserIndex).flags.Navegando = 0 And Not UserList(UserIndex).flags.Meditando Then EnviarPaquete Paquetes.SangraUser, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea
    End If
    Call NpcDaño(npcIndex, UserIndex)
    '¿Puede envenenar?
    
    If UserList(UserIndex).flags.Meditando = True Then
    EnviarPaquete Paquetes.Meditando, "", UserIndex
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    EnviarPaquete Paquetes.HechizoFX, ITS(UserList(UserIndex).Char.charIndex) & ByteToString(0) & ITS(0), UserIndex, ToMap, UserList(UserIndex).pos.map
    End If


    If UserList(UserIndex).Stats.minHP > 0 And NpcList(npcIndex).Veneno = 1 And UserList(UserIndex).flags.Envenenado = 0 Then Call NpcEnvenenarUser(UserIndex)

Else
    'EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), UserIndex, ToPCArea
    EnviarPaquete Paquetes.COMBNpcFalla, "", UserIndex
End If
'-----Tal vez suba los skills------
Call SubirSkill(UserIndex, tacticas)
Call SendUserStatsBox(UserIndex)
End Sub

Function NpcImpactoNpc(ByVal atacante As Integer, ByVal victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long
Dim ProbExito As Long

PoderAtt = NpcList(atacante).PoderAtaque
PoderEva = NpcList(victima).PoderEvasion
ProbExito = HelperMatematicas.maxs(10, HelperMatematicas.mins(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDañoNpc(ByVal atacante As Integer, ByVal victima As Integer)
Dim daño As Integer
Dim nIndex As Integer
Dim AnguloNPC As Single

daño = RandomNumber(NpcList(atacante).Stats.MinHIT, NpcList(atacante).Stats.MaxHIT)
NpcList(victima).Stats.minHP = NpcList(victima).Stats.minHP - daño

' Si es una mascota, entonce le da experiencia al maestro por cada bife que le pone.
If NpcList(atacante).MaestroUser <> 0 Then
    Call CalcularDarExp(UserList(NpcList(atacante).MaestroUser), NpcList(victima), daño)
End If

If NpcList(victima).Stats.minHP < 1 Then
    ' Si tiene mascota es como si lo hubiese matado el
    If NpcList(atacante).MaestroUser > 0 Then
        Call UsuarioMataNPC(UserList(NpcList(atacante).MaestroUser), NpcList(victima))
    End If
    
    nIndex = MuereNpc(NpcList(victima))
    
    If nIndex > 0 Then
        If DeboEnviarAngulo(UserList(NpcList(atacante).MaestroUser).pos.map) Then
            AnguloNPC = Angulo(NpcList(nIndex).pos.x, NpcList(nIndex).pos.y, UserList(NpcList(atacante).MaestroUser).pos.x, UserList(NpcList(atacante).MaestroUser).pos.y)
            EnviarPaquete Paquetes.AnguloNPC, ITS(AnguloNPC), NpcList(atacante).MaestroUser, ToIndex
        End If
    End If
    
End If
End Sub

Public Sub NpcAtacaNpc(ByVal atacante As Integer, ByVal victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
' El npc puede atacar ???

If NpcList(victima).MaestroUser > 0 Then
    Call CheckPets(atacante, NpcList(victima).MaestroUser)
End If

If NpcList(atacante).MaestroUser > 0 And MapInfo(NpcList(atacante).pos.map).PermiteRoboNPC = 0 Then
    'Anti robo de npcs
    Dim otroUsuario As Integer
    Dim UserIndex As Integer

    If NpcList(victima).MaestroUser = 0 Then
        otroUsuario = estaLuchando(NpcList(victima))
        UserIndex = NpcList(atacante).MaestroUser
    
        If Not otroUsuario = UserIndex And otroUsuario > 0 Then
            If Not AntiRoboNpc.puedePegarleAlNpc(UserIndex, otroUsuario) Then
                Call FollowAmo(atacante)
                Exit Sub
            End If
        Else
            If UserList(UserIndex).LuchandoNPC <> victima And UserList(UserIndex).LuchandoNPC > 0 Then
                'Si antes le estaba pegando a otro npc, libero a ese npc
                Call AntiRoboNpc.resetearLuchador(NpcList(UserList(UserIndex).LuchandoNPC))
            End If

            NpcList(victima).UltimoGolpe = GetTickCount()
            NpcList(victima).UserIndexLucha = UserIndex
            UserList(UserIndex).LuchandoNPC = victima
        End If
    End If
End If

If NpcList(atacante).flags.Snd1 > 0 Then EnviarPaquete Paquetes.WavSnd, Chr$(NpcList(atacante).flags.Snd1), atacante, ToNPCArea, NpcList(atacante).pos.map
If NpcImpactoNpc(atacante, victima) Then
    If NpcList(victima).flags.Snd2 > 0 Then
        EnviarPaquete Paquetes.WavSnd, Chr$((NpcList(victima).flags.Snd2)), victima, ToNPCArea, NpcList(victima).pos.map
    Else
        EnviarPaquete Paquetes.WavSnd, Chr$(SND_IMPACTO2), victima, ToNPCArea, NpcList(victima).pos.map
    End If
    
    If NpcList(atacante).MaestroUser > 0 Then
        EnviarPaquete Paquetes.WavSnd, Chr$(SND_IMPACTO), atacante, ToNPCArea, NpcList(atacante).pos.map
    Else
        EnviarPaquete Paquetes.WavSnd, Chr$(SND_IMPACTO), victima, ToNPCArea, NpcList(victima).pos.map
    End If
    Call NpcDañoNpc(atacante, victima)
Else
    If NpcList(atacante).MaestroUser > 0 Then
        EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), atacante, ToNPCArea, NpcList(atacante).pos.map
    Else
        EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), victima, ToNPCArea, NpcList(victima).pos.map
    End If
End If
End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal npcIndex As Integer)


If distancia(UserList(UserIndex).pos, NpcList(npcIndex).pos) > MAXDISTANCIAARCO Then
   EnviarPaquete Paquetes.MensajeSimple, Chr$(6), UserIndex
   Exit Sub
End If

If Not PuedeAtacarNPC(UserList(UserIndex), NpcList(npcIndex)) Then
    Exit Sub
End If
    
'Anti robo de npcs
If Not AntiRoboNpc.puedeLucharContraELNPC(NpcList(npcIndex), UserList(UserIndex)) Then
    Exit Sub
End If

Call NpcAtacado(npcIndex, UserIndex)

If UserImpactoNpc(UserIndex, npcIndex) Then
    EnviarPaquete Paquetes.AnimGolpe, Codify(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea, UserList(UserIndex).pos.map
    If NpcList(npcIndex).flags.Snd2 > 0 Then
        EnviarPaquete Paquetes.WavSnd, Chr$(NpcList(npcIndex).flags.Snd2), UserIndex, ToPCArea
    Else
        EnviarPaquete Paquetes.WavSnd, Chr$(SND_IMPACTO2), UserIndex, ToPCArea
    End If
    Call UserDañoNpc(UserIndex, npcIndex)
Else
    EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), UserIndex, ToPCArea
    EnviarPaquete Paquetes.AnimGolpe, Codify(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea, UserList(UserIndex).pos.map
    EnviarPaquete Paquetes.COMBUserFalla, "", UserIndex, ToIndex, 0
End If
End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

Dim AttackPos As WorldPos
Dim loquebaja As Integer


If IntervaloPermiteAtacar(UserIndex) Then

    'Quitamos la energia
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia = 0 Then
            loquebaja = RandomNumber(1, 10)
            
            If UserList(UserIndex).Stats.MinSta - loquebaja <= 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(11), UserIndex
                Exit Sub
            Else
                Call QuitarSta(UserIndex, loquebaja)
            End If
        Else
             
            If UserList(UserIndex).Stats.MinSta >= ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia Then
                Call QuitarSta(UserIndex, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia)
            Else
                EnviarPaquete Paquetes.MensajeSimple, Chr$(11), UserIndex
                Exit Sub
            End If
        End If
    Else
        Call QuitarSta(UserIndex, RandomNumber(1, 10))
    End If


    AttackPos = UserList(UserIndex).pos
    Call HeadtoPos(UserList(UserIndex).Char.heading, AttackPos)

    'Exit if not legal
    If AttackPos.x < X_MINIMO_USABLE Or AttackPos.x > X_MAXIMO_USABLE Or AttackPos.y < Y_MINIMO_USABLE Or AttackPos.y > Y_MAXIMO_USABLE Then
        EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), UserIndex, ToPCArea
        EnviarPaquete Paquetes.AnimGolpe, Codify(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea, UserList(UserIndex).pos.map
        Exit Sub
    End If
    
    Dim index As Integer
    index = MapData(AttackPos.map, AttackPos.x, AttackPos.y).UserIndex
    
    If index > 0 Then
        Call UsuarioAtacaUsuario(UserIndex, index)
        Call SendUserStatsBox(index)
        Exit Sub
    End If

    If MapData(AttackPos.map, AttackPos.x, AttackPos.y).npcIndex > 0 Then
       
       '[eLwE 19/05/05]Comenta que hiciste O_O
        If NpcList(MapData(AttackPos.map, AttackPos.x, AttackPos.y).npcIndex).Attackable Then
            If NpcList(MapData(AttackPos.map, AttackPos.x, AttackPos.y).npcIndex).MaestroUser > 0 Then
                If MapInfo(NpcList(MapData(AttackPos.map, AttackPos.x, AttackPos.y).npcIndex).pos.map).Pk = False Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(173), UserIndex
                    Exit Sub
                End If
            End If
            Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.map, AttackPos.x, AttackPos.y).npcIndex)
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr$(174), UserIndex
        End If
    Else
        EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), UserIndex, ToPCArea
        EnviarPaquete Paquetes.AnimGolpe, Codify(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea
    End If
End If
Call SendUserStatsBox(UserIndex)
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(tacticas)
SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(Defensa)
Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
If Arma > 0 Then
    proyectil = ObjData(Arma).proyectil = 1
Else
    proyectil = False
End If
'Calculamos el poder de evasion...
UserPoderEvasion = PoderEvasion(UserList(VictimaIndex))

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(UserList(VictimaIndex))
   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(UserList(AtacanteIndex))
    Else
        PoderAtaque = PoderAtaqueArma(UserList(AtacanteIndex))
    End If
    ProbExito = HelperMatematicas.maxs(10, HelperMatematicas.mins(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
Else
    PoderAtaque = PoderAtaqueWresterling(AtacanteIndex)
    ProbExito = HelperMatematicas.maxs(10, HelperMatematicas.mins(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
End If
UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
' el usuario esta usando un escudo ???
If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = HelperMatematicas.maxs(10, HelperMatematicas.mins(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
              EnviarPaquete Paquetes.WavSnd, Chr$(SND_ESCUDO), AtacanteIndex, ToPCArea
              EnviarPaquete Paquetes.COMBRechEsc, "", VictimaIndex
              EnviarPaquete Paquetes.COMBEnemEscu, "", AtacanteIndex
              EnviarPaquete Paquetes.AnimEscu, ITS(UserList(VictimaIndex).Char.charIndex), VictimaIndex, ToPCArea
              Call SubirSkill(VictimaIndex, Defensa)
      End If
    End If
End If
If UsuarioImpacto Then
   If Arma > 0 Then
           If Not proyectil Then
                  Call SubirSkill(AtacanteIndex, Armas)
           Else
                  Call SubirSkill(AtacanteIndex, proyectiles)
           End If
   Else
        Call SubirSkill(AtacanteIndex, Wresterling)
   End If
End If
End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not puedeAtacar(UserList(AtacanteIndex), UserList(VictimaIndex)) Then Exit Sub

If distancia(UserList(AtacanteIndex).pos, UserList(VictimaIndex).pos) > MAXDISTANCIAARCO Then
   EnviarPaquete Paquetes.MensajeSimple, Chr$(6), AtacanteIndex
   Exit Sub
End If

Call UsuarioAtacadoPorUsuario(UserList(AtacanteIndex), UserList(VictimaIndex))

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then

    EnviarPaquete Paquetes.WavSnd, Chr$(SND_IMPACTO), AtacanteIndex, ToPCArea
    
    If UserList(VictimaIndex).flags.Navegando = 0 And UserList(VictimaIndex).flags.Meditando = False Then EnviarPaquete Paquetes.SangraUser, ITS(UserList(VictimaIndex).Char.charIndex), VictimaIndex, ToPCArea
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
    
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex = 0 Then
        Call Desarmar(AtacanteIndex, VictimaIndex)
    End If
    
Else
    EnviarPaquete Paquetes.AnimGolpe, ITS(UserList(AtacanteIndex).Char.charIndex), AtacanteIndex, ToPCArea
    EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), AtacanteIndex, ToPCArea
    EnviarPaquete Paquetes.COMBUserFalla, "", AtacanteIndex
    EnviarPaquete Paquetes.COMBEnemFalla, UserList(AtacanteIndex).Name, VictimaIndex
End If

If UserList(AtacanteIndex).clase = eClases.Ladron Then Call Desarmar(AtacanteIndex, VictimaIndex)

End Sub


Private Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim daño As Long, antdaño As Integer
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer
Dim obj As ObjData
Dim Resist As Byte

daño = CalcularDaño(AtacanteIndex)

antdaño = daño

Call UserEnvenena(AtacanteIndex, VictimaIndex)

If UserList(AtacanteIndex).flags.Navegando = 1 Then
     obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     daño = daño + RandomNumber(obj.MinHIT, obj.MaxHIT)
End If

If UserList(VictimaIndex).flags.Navegando = 1 Then
     obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(obj.MinDef, obj.MaxDef)
End If

If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    Resist = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Refuerzo
End If

Lugar = RandomNumber(1, 6)

'¿Donde le pego?
Select Case Lugar
  Case bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
           obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
           absorbido = absorbido + defbarco - (absorbido * Resist * 0.01)
           daño = daño - absorbido
           If daño < 0 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
           obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
           absorbido = absorbido + defbarco - (absorbido * Resist * 0.01)
        End If
        
        ' Penalizacion por Equiparse mal
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
            If UserList(VictimaIndex).Stats.ELV >= 25 And ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex).MaxDef < 15 Then
                daño = daño * 1.15
            End If
        End If
        
        If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
            obj = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
            absorbido = absorbido + RandomNumber(obj.MinDef, obj.MaxDef)
        End If
        
        daño = daño - absorbido
        
        If daño < 0 Then daño = 1
End Select

'Efectos
EnviarPaquete Paquetes.COMBUserHITUser, Chr$(Lugar) & ITS(daño) & UserList(VictimaIndex).Name, AtacanteIndex
EnviarPaquete Paquetes.COMBEnemHitUs, Chr$(Lugar) & ITS(daño) & UserList(AtacanteIndex).Name, VictimaIndex
EnviarPaquete Paquetes.AnimGolpe, ITS(UserList(AtacanteIndex).Char.charIndex), AtacanteIndex, ToPCArea

Dim danioApu As Integer

danioApu = 0

If PuedeApuñalar(AtacanteIndex) Then
    danioApu = DoApuñalar(UserList(AtacanteIndex), 0, VictimaIndex, daño)
    Call SubirSkill(AtacanteIndex, Apuñalar)
End If

    
' Actualizo la vida
Dim danioTotal As Integer


If danioApu > 0 Then

    If UserList(AtacanteIndex).clase = eClases.asesino Then
        If daño + danioApu > UserList(VictimaIndex).Stats.minHP + 40 Then
            ' Lo hizo pelota, lo mata mal
           danioTotal = daño + danioApu
        ElseIf daño + danioApu >= UserList(VictimaIndex).Stats.minHP Then
            ' Lo podria matar, pero no lo mata, queda en uno de vida.
            danioApu = UserList(VictimaIndex).Stats.minHP - 1 - daño
            danioTotal = UserList(VictimaIndex).Stats.minHP - 1
        Else
            danioTotal = daño + danioApu ' El golpe real
        End If
    Else
        danioTotal = daño + danioApu
    End If
    
    If danioApu > 0 Then
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(5) & UserList(VictimaIndex).Name & "," & danioApu, AtacanteIndex
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(4) & UserList(AtacanteIndex).Name & "," & danioApu, VictimaIndex

        EnviarPaquete Paquetes.TXAII, ITS(UserList(VictimaIndex).pos.x) & ITS(UserList(VictimaIndex).pos.y) & ITS(danioTotal) & ITS(distancia(UserList(VictimaIndex).pos, UserList(AtacanteIndex).pos)), VictimaIndex, ToIndex
        EnviarPaquete Paquetes.TXAII, ITS(UserList(VictimaIndex).pos.x) & ITS(UserList(VictimaIndex).pos.y) & ITS(danioTotal) & ITS(distancia(UserList(VictimaIndex).pos, UserList(AtacanteIndex).pos)), AtacanteIndex, ToIndex
    
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(8) & Int(danioTotal), AtacanteIndex
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(7) & Int(danioTotal), VictimaIndex
    End If
Else
    danioTotal = daño
    EnviarPaquete Paquetes.TXA, ITS(UserList(VictimaIndex).pos.x) & ITS(UserList(VictimaIndex).pos.y) & ITS(daño) & ITS(distancia(UserList(VictimaIndex).pos, UserList(AtacanteIndex).pos)), VictimaIndex, ToIndex
    EnviarPaquete Paquetes.TXA, ITS(UserList(VictimaIndex).pos.x) & ITS(UserList(VictimaIndex).pos.y) & ITS(daño) & ITS(distancia(UserList(VictimaIndex).pos, UserList(AtacanteIndex).pos)), AtacanteIndex, ToIndex
End If

' Definitivamnte le resto la vida
UserList(VictimaIndex).Stats.minHP = UserList(VictimaIndex).Stats.minHP - danioTotal
 
If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then
    'Si usa un arma quizas suba "Combate con armas"
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call SubirSkill(AtacanteIndex, Armas)
    Else
        'sino tal vez lucha libre
        Call SubirSkill(AtacanteIndex, Wresterling)
    End If
    
    Call SubirSkill(VictimaIndex, tacticas)
End If

'¿Murio la victima?
If UserList(VictimaIndex).Stats.minHP <= 0 Then
     Call ContarMuerte(UserList(VictimaIndex), UserList(AtacanteIndex))
     ' Para que las mascotas no sigan intentando luchar y
     ' comiencen a seguir al amo
     Dim j As Integer
     For j = 1 To MAXMASCOTAS
        If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
            If NpcList(UserList(AtacanteIndex).MascotasIndex(j)).TargetUserID = UserList(VictimaIndex).id Then NpcList(UserList(AtacanteIndex).MascotasIndex(j)).TargetUserID = 0
            Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
        End If
     Next j
     Call UsuarioMataAUsuario(VictimaIndex, AtacanteIndex)
End If

' Enviamos estadisticas
Call SendUserStatsBox(VictimaIndex)
End Sub

Public Sub UsuarioAtacadoPorUsuario(ByRef atacante As User, ByRef victima As User)
        
' Si el personaje está saliendo naturalmente, cancelamos la salida
If victima.flags.Saliendo = eTipoSalida.SaliendoNaturalmente Then
    victima.flags.Saliendo = eTipoSalida.NoSaliendo
    victima.Counters.Salir = 0
End If

' Si la Victima esta Meditando, deja de hacerlo
If victima.flags.Meditando = True Then
    victima.flags.Meditando = False
    victima.Char.FX = 0
    victima.Char.loops = 0
    
    EnviarPaquete Paquetes.Meditando, "", victima.UserIndex
    EnviarPaquete Paquetes.HechizoFX, ITS(victima.Char.charIndex) & ByteToString(0) & ITS(0), victima.UserIndex, ToPCArea, victima.pos.map
End If

' ¿Acá se puede atacar sin consecuencias?
If EsPosicionParaAtacarSinPenalidad(atacante.pos) And EsPosicionParaAtacarSinPenalidad(victima.pos) Then Exit Sub

' Las mascotas se ponen en situacion de ataque
Call AllMascotasAtacanUser(atacante.UserIndex, victima.UserIndex)
Call AllMascotasAtacanUser(victima.UserIndex, atacante.UserIndex)

End Sub

Sub AllMascotasAtacanUser(ByVal Victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer
For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            NpcList(UserList(Maestro).MascotasIndex(iCount)).TargetUserID = UserList(Victim).id
            NpcList(UserList(Maestro).MascotasIndex(iCount)).Movement = NPCDEFENSA
    End If
Next iCount
End Sub

Public Function PuedeAtacarNPC(ByRef Usuario As User, ByRef criatura As npc) As Boolean

If criatura.Attackable = 0 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(144), Usuario.UserIndex
    Exit Function
End If
    
If Not criatura.faccion = eAlineaciones.indefinido And Not criatura.faccion = eAlineaciones.Neutro And criatura.faccion = Usuario.faccion.alineacion Then
    EnviarPaquete Paquetes.MensajeFight, "No puedes atacar criaturas de tu facción.", Usuario.UserIndex
    PuedeAtacarNPC = False
    Exit Function
End If

If criatura.MaestroUser > 0 Then
    If MapInfo(Usuario.pos.map).Pk = False Then
        EnviarPaquete Paquetes.MensajeFight, "No puedes atacar mascotas en zonas seguras.", Usuario.UserIndex
        Exit Function
    End If
    
    If Not Usuario.faccion.alineacion = eAlineaciones.Neutro Then
        If Usuario.faccion.alineacion = UserList(criatura.MaestroUser).faccion.alineacion Then
            EnviarPaquete Paquetes.MensajeFight, "No puedes atacar masotas de integrantes de tu ejército!.", Usuario.UserIndex
            PuedeAtacarNPC = False
            Exit Function
        End If
    End If
End If

If Usuario.flags.Muerto = 1 Then
    EnviarPaquete Paquetes.mensajeinfo, "No podés atacar porque estas muerto.", Usuario.UserIndex
    PuedeAtacarNPC = False
    Exit Function
End If

If Usuario.flags.Privilegios = 1 Then
    PuedeAtacarNPC = False
    Exit Function
End If

PuedeAtacarNPC = True
End Function


Private Sub entregarExperienciaGolpeCriatura(ByVal expDar As Long, ByRef personaje As User, ByRef criatura As npc)

If personaje.PartyIndex > 0 Then
    Call mdParty.ObtenerExito(personaje.UserIndex, criatura, expDar)
Else
    Dim expFinal As Long
    ' Penalizador por diferencia de nivel
    If personaje.Stats.ELV > criatura.Nivel Then
        expFinal = expDar * PENALIZACION_CRIATURA_MENOR_NIVEL_USUARIO
    Else
        expFinal = expDar
    End If

    Call modUsuarios.agregarExperiencia(personaje.UserIndex, expFinal)
    Call modPersonaje_TCP.actualizarExperiencia(personaje)
    
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(3) & expFinal, personaje.UserIndex
End If

End Sub

Public Sub CalcularDarExpUltimoGolpe(ByRef personaje As User, ByRef criatura As npc)

Dim ExpaDar As Long

If criatura.flags.ExpCount = 0 Then Exit Sub

ExpaDar = criatura.flags.ExpCount

Call entregarExperienciaGolpeCriatura(ExpaDar, personaje, criatura)

criatura.flags.ExpCount = 0
    
End Sub

' Este metodo calcular y le otorga la experiencia que le da al usuario golpear a la criatura.
Public Sub CalcularDarExp(ByRef personaje As User, ByRef criatura As npc, ByVal ElDaño As Long)

Dim ExpNPC As Long
Dim ExpaDar As Long
Dim TotalNpcVida As Long
Dim ExpSinMorir As Long

If criatura.flags.ExpCount = 0 Then Exit Sub

ExpNPC = criatura.GiveEXP

ExpSinMorir = (2 * criatura.GiveEXP) / 3
TotalNpcVida = criatura.Stats.MaxHP

If TotalNpcVida <= 0 Then
    Exit Sub
End If

If ElDaño > criatura.Stats.minHP Then
    ElDaño = criatura.Stats.minHP
End If

If ElDaño < 0 Then
    ElDaño = 0
End If

'totalnpcvida _____ ExpSinMorir
'daño         _____ (daño * ExpSinMorir) / totalNpcVida
'ExpaDar = CLng((ElDaño) * (ExpSinMorir / TotalNpcVida))
' [Cada vez que se golpea a un npc da la misma exp sin importar que se el ultimo]
Dim danioSingle As Single
danioSingle = ElDaño

ExpaDar = CLng((danioSingle * ExpSinMorir) / TotalNpcVida)

If ExpaDar <= 0 Then
    Exit Sub
End If

If ExpaDar > criatura.flags.ExpCount Then
    ExpaDar = criatura.flags.ExpCount
    criatura.flags.ExpCount = 0
Else
    criatura.flags.ExpCount = criatura.flags.ExpCount - ExpaDar
End If

Call entregarExperienciaGolpeCriatura(ExpaDar, personaje, criatura)

End Sub

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
ObjInd = 0
If ArmaObjInd > 0 Then
    If ObjData(ArmaObjInd).proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
    End If
    If ObjInd > 0 Then
        If (ObjData(ObjInd).Envenena = 1) Then
            num = RandomNumber(1, 100)
            If num < 60 Then
                UserList(VictimaIndex).flags.Envenenado = 1
                EnviarPaquete Paquetes.EstaEnvenenado, "", VictimaIndex, ToIndex
                EnviarPaquete Paquetes.MensajeFight, "Has envenenado a " & UserList(VictimaIndex).Name & "!!", AtacanteIndex
                EnviarPaquete Paquetes.MensajeFight, UserList(AtacanteIndex).Name & " te ha envenenado!!", VictimaIndex
            End If
        End If
    End If
End If
End Sub

Private Function DañoExtra(ByVal UserIndex As Integer) As Integer
Dim NombreDelObjeto As String

If UserList(UserIndex).clase = eClases.Cazador Then
    NombreDelObjeto = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name
    If InStr(1, NombreDelObjeto, "reforzado") > 0 Or NombreDelObjeto = "Arco de Cazador" Then
        DañoExtra = UserList(UserIndex).Stats.ELV
        EnviarPaquete Paquetes.MensajeFight, "Has echo un golpe critico por " & DañoExtra & ".", UserIndex, ToIndex
    Else
        DañoExtra = 0
    End If
Else
    DañoExtra = 0
End If
End Function

