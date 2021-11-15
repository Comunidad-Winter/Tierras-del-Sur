Attribute VB_Name = "modHechizos"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

'********************Misery_Ezequiel 28/05/05********************'
Option Explicit
'[Misery_Ezequiel 10/06/05]
Public Const SUPERANILLO = 649
'[\]Misery_Ezequiel 10/06/05]

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)
If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
'[Wizard 03/09/05] Este sub fue modificado para que al meditar los hechizos de daño te desconcentren y los otros no se vean.
If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Mimetizado = 1 Then Exit Sub

If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
     daño = daño - (daño * (UserList(UserIndex).Stats.UserSkills(1) / 2000)) 'Resistencia magica
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    If Not UserList(UserIndex).flags.Meditando Then Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + daño
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    Call Senddata(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Call SendUserStatsBox(val(UserIndex))
ElseIf Hechizos(Spell).SubeHP = 2 Then
    If UserList(UserIndex).flags.Privilegios = 0 And UserList(UserIndex).flags.Mimetizado = 0 Then
        If UserList(UserIndex).flags.Meditando Then
            UserList(UserIndex).flags.Meditando = False
            UserList(UserIndex).Counters.bPuedeMeditar = False
            Call Senddata(ToIndex, UserIndex, 0, "Y377")
            Call Senddata(ToIndex, UserIndex, 0, "MEDOK")
        End If
        
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        daño = daño - (daño * (UserList(UserIndex).Stats.UserSkills(1) / 2000)) 'Resistencia magica
        'marche
        If daño < 0 Then daño = 0
        Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
        If Not UserList(UserIndex).flags.Meditando Then Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
        Call Senddata(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
        Call SendUserVida(val(UserIndex)) 'Marche 3-8
        'Muere
        If UserList(UserIndex).Stats.MinHP < 1 Then
            UserList(UserIndex).Stats.MinHP = 0
            If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                RestarCriminalidad (UserIndex)
            End If
            Call UserDie(UserIndex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    End If
End If
If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 And UserList(UserIndex).flags.Mimetizado = 0 Then
        Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
       If Not UserList(UserIndex).flags.Meditando Then Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        '[Misery_Ezequiel 10/06/05]
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
            'Call Senddata(ToIndex, UserIndex, 0, "||Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
            Call Senddata(ToIndex, UserIndex, 0, "Y348")
            Exit Sub
        End If
        '[\]Misery_Ezequiel 10/06/05]
        UserList(UserIndex).flags.Paralizado = 1
        UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
        
       ' If EncriptarProtocolosCriticos Then
        '  Call SendCryptedData(ToIndex, UserIndex, 0, "PARADOK")
        'Else
          Call Senddata(ToIndex, UserIndex, 0, "PARADOK")
        'End If
        Call Senddata(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y) 'Gorlok
    End If
End If
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
On Error GoTo errhandler
Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next
Exit Function
errhandler:
End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer

'agregar nacho

hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex
If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y132")
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call Senddata(ToIndex, UserIndex, 0, "Y133")
End If
End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)
On Error Resume Next
    Dim ind As String
    ind = UserList(UserIndex).Char.charindex
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbCyan & "°" & s & "°" & ind)
    Exit Sub
End Sub

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
If UserList(UserIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY
         'marche
        If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call Senddata(ToIndex, UserIndex, 0, "||Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call Senddata(ToIndex, UserIndex, 0, "||Necesitas un baculo para lanzar este hechizo!!." & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
    'marche
    If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(UserIndex).Stats.UserSkills(Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(UserIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y134")
                PuedeLanzar = False
            End If
        Else
            Call Senddata(ToIndex, UserIndex, 0, "Y135")
            PuedeLanzar = False
        End If
    Else
            Call Senddata(ToIndex, UserIndex, 0, "Y136")
            PuedeLanzar = False
    End If
Else
   Call Senddata(ToIndex, UserIndex, 0, "Y137")
   PuedeLanzar = False
End If
End Function

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)
  'marcelo
  
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 3 Then Exit Sub
  
   
'Call LogTarea("HechizoInvocacion")
If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos
TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

For j = 1 To Hechizos(H).cant
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind <= MAXNPCS Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            Index = FreeMascotaIndex(UserIndex)
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero
            Npclist(ind).MaestroUser = UserIndex
            
            If UCase$(Hechizos(H).Nombre) = UCase$("Invocar Elemental de fuego") Then
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion * 0.667
            Else
              If UCase$(Hechizos(H).Nombre) = UCase$("Invocar Elemental de tierra") Then
              Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion * 0.667
                Else
                Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
                End If
            End If
            
            Npclist(ind).GiveGLD = 0
            Call FollowAmo(ind)
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
'[Misery_Ezequiel 26/06/05]
Select Case Hechizos(uh).Tipo
    Case uInvocacion '
        Call HechizoInvocacion(UserIndex, b)
    Case uEstado
        Call HechizoTerrenoEstado(UserIndex, b)
End Select
'[\]Misery_Ezequiel 26/06/05]
If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendUserStatsBox(UserIndex)
End If
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
End Select
If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendUserStatsBox(UserIndex)
    Call SendUserStatsBox(UserList(UserIndex).flags.TargetUser)
    UserList(UserIndex).flags.TargetUser = 0
End If
End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
End Select
If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendUserStatsBox(UserIndex)
End If
End Sub

Sub LanzarHechizo(Index As Integer, UserIndex As Integer)
Dim uh As Integer
Dim exito As Boolean

uh = UserList(UserIndex).Stats.UserHechizos(Index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        Case uUsuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y138")
            End If
        Case uNPC
            If UserList(UserIndex).flags.TargetNPC > 0 Or UserList(UserIndex).flags.TargetObj = 147 Or UserList(UserIndex).flags.TargetObj = 148 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y139")
            End If
        Case uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y140")
            End If
        Case uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
End If
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim H As Integer, TU As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
TU = UserList(UserIndex).flags.TargetUser

If Hechizos(H).Revivir = 1 Then
   If UserList(TU).flags.Muerto = 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y281")
        b = False
     Exit Sub
    Else
    
        If Criminal(TU) And Not Criminal(UserIndex) Then
                If UserList(UserIndex).flags.Seguro Then
                    Call Senddata(ToIndex, UserIndex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos" & FONTTYPE_INFO)
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)
                End If
        End If
    
         If UserList(TU).flags.ModoCombate = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y281")
            b = False
            Exit Sub
        End If
        
    UserList(TU).Stats.MinMAN = 0
        Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)
        Call Senddata(ToIndex, UserIndex, 0, "Y143")
 
   
           'revisamos si necesita vara
        If UCase$(UserList(UserIndex).Clase) = "a" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y282")
                    b = False
                    Exit Sub
                End If
            End If
        '[\]Misery_Ezequiel 10/06/05]
        ElseIf UCase$(UserList(UserIndex).Clase) = "BARDO" Then
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> LAUDMAGICO Then
                Call Senddata(ToIndex, UserIndex, 0, "Y344")
                b = False
                Exit Sub
            End If
        ElseIf UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> ANILLOMAGICODRUIDA Then
            '[Misery_Ezequiel 26/06/05]
                Call Senddata(ToIndex, UserIndex, 0, "Y352")
            '[\]Misery_Ezequiel 26/06/05]
                b = False
                Exit Sub
            End If
        End If
        '[\]Misery_Ezequiel 10/06/05]
       '[eLwE 20/05/05]Si tira a un criminal le dice que se saque el seguro.
        If Not Criminal(UserIndex) And Criminal(TU) Then
            If UserList(UserIndex).flags.Seguro = True Then
                Call Senddata(ToIndex, UserIndex, 0, "Y277")
                Exit Sub
            Else
                VolverCriminal (UserIndex)
            End If
        End If
       '[\]eLwE 19/05/05]
       '[Misery_Ezequiel 26/06/05]
        UserList(TU).Stats.MinAGU = 0
        UserList(TU).Stats.MinHam = 0
        UserList(TU).flags.Hambre = 1
        UserList(TU).flags.Sed = 1
        Call RevivirUsuario(TU)
        Call EnviarHambreYsed(TU)
        '[\]Misery_Ezequiel 26/06/05]
    End If
    Call InfoHechizo(UserIndex)
    b = True
End If
'[Misery_Ezequiel 05/06/05]
If UserList(UserIndex).flags.Muerto = 1 Then
   Call Senddata(ToIndex, UserIndex, 0, "Y343")
   Exit Sub
End If
'[\]Misery_Ezequiel 05/06/05]
If Hechizos(H).Invisibilidad = 1 Then

     If Criminal(TU) And Not Criminal(UserIndex) Then
            If UserList(UserIndex).flags.Seguro Then
                Call Senddata(ToIndex, UserIndex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos" & FONTTYPE_INFO)
                Exit Sub
            Else
                Call VolverCriminal(UserIndex)
            End If
        End If
   'marcelo

    If UserList(TU).flags.Muerto = 1 Then
        Call Senddata(ToIndex, UserIndex, 0, "||¡Está muerto!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If UserList(TU).flags.Invisible = 1 Or UserList(TU).flags.Oculto = 1 Then Exit Sub
   UserList(TU).flags.Invisible = 1
   'If EncriptarProtocolosCriticos Then
    '  Call SendCryptedData(ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.charindex & ",1")
  ' Else
    Call Senddata(ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.charindex & ",1")
  ' End If
   Call InfoHechizo(UserIndex)
   b = True
End If
If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        b = True
End If
If Hechizos(H).CuraVeneno = 1 Then
        If UserList(TU).flags.Envenenado = 0 Then 'Gorlok - No curar si no está/s envenenado.
            If UserIndex <> TU Then
                Call Senddata(ToIndex, UserIndex, 0, "||No está envenenado." & FONTTYPE_INFO)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "||No estás envenenado." & FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        If Criminal(TU) And Not Criminal(UserIndex) Then
            If UserList(UserIndex).flags.Seguro Then
                Call Senddata(ToIndex, UserIndex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos" & FONTTYPE_INFO)
                Exit Sub
            Else
                Call VolverCriminal(UserIndex)
            End If
        End If
        '[eLwE 20/05/05]Si tira a un criminal le dice que se saque el seguro.
        If Not Criminal(UserIndex) And Criminal(TU) Then
            If UserList(UserIndex).flags.Seguro = True Then
                Call Senddata(ToIndex, UserIndex, 0, "Y277")
                Exit Sub
            Else
                VolverCriminal (UserIndex)
            End If
        End If
      '[\]eLwE 20/05/05]
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True
End If
If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If
If Hechizos(H).RemoverMaldicion = 1 Then
        '[eLwE 20/05/05]Si tira a un criminal le dice que se saque el seguro.
        If Not Criminal(UserIndex) And Criminal(TU) Then
            If UserList(UserIndex).flags.Seguro = True Then
                Call Senddata(ToIndex, UserIndex, 0, "Y277")
                Exit Sub
            Else
                VolverCriminal (UserIndex)
            End If
        End If
        '[\]eLwE 20/05/05]
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If
If Hechizos(H).Bendicion = 1 Then
        '[eLwE 20/05/05]Si tira a un criminal le dice que se saque el seguro.
        If Not Criminal(UserIndex) And Criminal(TU) Then
            If UserList(UserIndex).flags.Seguro = True Then
                Call Senddata(ToIndex, UserIndex, 0, "Y277")
                Exit Sub
            Else
                VolverCriminal (UserIndex)
            End If
        End If
        '[\]eLwE 19/05/05]
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If
If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
    If UserIndex = TU Then
        Call Senddata(ToIndex, UserIndex, 0, "Y275")
        Exit Sub 'Prohibido paralizarse a si mismo - byGorlok 2005-03-25
    End If
    If UserList(TU).flags.Paralizado = 0 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        
        Call InfoHechizo(UserIndex)
        b = True
        '[Misery_Ezequiel 10/06/05]
        If UserList(TU).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
            'Call Senddata(ToIndex, TU, 0, "||Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
            Call Senddata(ToIndex, TU, 0, "Y348")
            'Call Senddata(ToIndex, UserIndex, 0, "||¡El hechizo no tiene efecto!" & FONTTYPE_FIGHT)
            Call Senddata(ToIndex, UserIndex, 0, "Y349")
            Exit Sub
        End If
        '[\]Misery_Ezequiel 10/06/05]
        UserList(TU).flags.Paralizado = 1
        UserList(TU).Counters.Paralisis = IntervaloParalizado
       ' If EncriptarProtocolosCriticos Then
        '    Call SendCryptedData(ToIndex, TU, 0, "PARADOK")
       ' Else
            Call Senddata(ToIndex, TU, 0, "PARADOK")
        'End If
        Call Senddata(ToIndex, TU, 0, "PU" & UserList(TU).Pos.X & "," & UserList(TU).Pos.Y) 'Gorlok
    End If
End If
'[Misery_Ezequiel 26/06/05]
If Hechizos(H).RemoverEstupidez = 1 Then
    If Criminal(TU) And Not Criminal(UserIndex) Then
            If UserList(UserIndex).flags.Seguro Then
                Call Senddata(ToIndex, UserIndex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos" & FONTTYPE_INFO)
                Exit Sub
            Else
                Call VolverCriminal(UserIndex)
            End If
        End If
    If Not UserList(TU).flags.Estupidez = 0 Then
            UserList(TU).flags.Estupidez = 0
            'no need to crypt this
            Call Senddata(ToIndex, TU, 0, "NESTUP")
            Call InfoHechizo(UserIndex)
            b = True
    End If
End If
'[\]Misery_Ezequiel 26/06/05]
If Hechizos(H).RemoverParalisis = 1 Then
    '[eLwE 20/05/05]Si tira a un criminal le dice que se saque el seguro.
    If Not Criminal(UserIndex) And Criminal(TU) Then
        If UserList(UserIndex).flags.Seguro = True Then
            Call Senddata(ToIndex, UserIndex, 0, "Y277")
            Exit Sub
        Else
            VolverCriminal (UserIndex)
        End If
    End If
   '[\]eLwE 19/05/05]
    If UserList(TU).flags.Paralizado = 1 Then
        UserList(TU).flags.Paralizado = 0
        'no need to crypt this
        Call Senddata(ToIndex, TU, 0, "PARADOK")
        Call InfoHechizo(UserIndex)
        b = True
    End If
End If
If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        Call Senddata(ToIndex, TU, 0, "CEGU")
        Call InfoHechizo(UserIndex)
        b = True
End If
If Hechizos(H).Estupidez = 1 Then
          'marcelo
        
        If UserList(TU).flags.Estupidez = 1 Then
        Exit Sub
        End If
        
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado - 15
       ' If EncriptarProtocolosCriticos Then
        '    Call SendCryptedData(ToIndex, TU, 0, "DUMB")
        'Else
            Call Senddata(ToIndex, TU, 0, "DUMB")
        'End If
        Call InfoHechizo(UserIndex)
        b = True
End If

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
'[Misery_Ezequiel 04/06/05]
 If UserList(UserIndex).flags.TargetObj <> 147 And UserList(UserIndex).flags.TargetObj <> 148 Then
If Npclist(NpcIndex).InmuneAHechizos = 1 Then
Call Senddata(ToIndex, UserIndex, 0, "Y342")
Exit Sub
End If
'[\]Misery_Ezequiel 04/06/05]
If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

  End If
If Hechizos(hIndex).Mimetiza = 1 Then
    If UserList(UserIndex).flags.Muerto = 1 Then
        Exit Sub
    End If
    
   
   If UserList(UserIndex).flags.Mimetizado = 1 Then
        Call Senddata(ToIndex, UserIndex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
        Exit Sub
   End If
    
    'copio el char original al mimetizado
    
    UserList(UserIndex).CharMimetizado.Body = UserList(UserIndex).Char.Body
    UserList(UserIndex).CharMimetizado.Head = UserList(UserIndex).Char.Head
    UserList(UserIndex).CharMimetizado.CascoAnim = UserList(UserIndex).Char.CascoAnim
    UserList(UserIndex).CharMimetizado.ShieldAnim = UserList(UserIndex).Char.ShieldAnim
    UserList(UserIndex).CharMimetizado.WeaponAnim = UserList(UserIndex).Char.WeaponAnim
    
   UserList(UserIndex).flags.Mimetizado = 1
   UserList(UserIndex).flags.Invisible = 1
   If UserList(UserIndex).flags.TargetObj = 147 Or UserList(UserIndex).flags.TargetObj = 148 Then
      UserList(UserIndex).Char.Body = 25
    UserList(UserIndex).Char.Head = 0
    UserList(UserIndex).Char.CascoAnim = 0
    UserList(UserIndex).Char.ShieldAnim = 0
    UserList(UserIndex).Char.WeaponAnim = 0
   
   Else
    'ahora pongo local el del enemigo
    UserList(UserIndex).Char.Body = Npclist(NpcIndex).Char.Body
    UserList(UserIndex).Char.Head = Npclist(NpcIndex).Char.Head
    UserList(UserIndex).Char.CascoAnim = Npclist(NpcIndex).Char.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = Npclist(NpcIndex).Char.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = Npclist(NpcIndex).Char.WeaponAnim
End If
    Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "CP" & UserList(UserIndex).Char.charindex & "," & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & UserList(UserIndex).Char.CascoAnim)
   
   Call InfoHechizo(UserIndex)
   b = True

End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y144")
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If
If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If
If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y144")
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 1
   b = True
End If
If Hechizos(hIndex).Maldicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If
If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If
If Hechizos(hIndex).Paraliza = 1 Then
   If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).flags.Inmovilizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            b = True
   Else
      Call Senddata(ToIndex, UserIndex, 0, "Y283")
   End If
End If
'[Barrin 16-2-04]
'[Wizard 03/09/05] Arregla el bug de el remover paralisis; pasa ENTERO El hechizo.
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If (Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1) And (Npclist(NpcIndex).MaestroUser = UserIndex) Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).flags.Inmovilizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call Senddata(ToIndex, UserIndex, 0, "Y146")
   End If
End If
'[/wizard]
'[/Barrin]
If Hechizos(hIndex).Inmoviliza = 1 Then
   If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
      If Npclist(NpcIndex).flags.Paralizado = 1 Then Exit Sub
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(UserIndex)
        b = True
   Else
      Call Senddata(ToIndex, UserIndex, 0, "Y283")
   End If
End If
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
Dim daño As Long
'[Misery_Ezequiel 04/06/05]
If Npclist(NpcIndex).InmuneAHechizos = 1 Then
Call Senddata(ToIndex, UserIndex, 0, "Y342")
Exit Sub
End If
'[\]Misery_Ezequiel 04/06/05]
'Salud

If Hechizos(hIndex).SubeHP = 1 Then
'[Misery_Ezequiel 04/06/05]
If Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
   Call Senddata(ToIndex, UserIndex, 0, "Y280")
   Exit Sub
Else
'[\]Misery_Ezequiel 04/06/05]
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    Call InfoHechizo(UserIndex)
    Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, daño, Npclist(NpcIndex).Stats.MaxHP)
    Call Senddata(ToIndex, UserIndex, 0, "||Has curado " & daño & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
End If
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    If Npclist(NpcIndex).Attackable = 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y144")
        Exit Sub
    End If
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
     'marche
    If Hechizos(hIndex).StaffAffected Then
        If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta daño segun el staff-
                'Daño = (Daño* (80 + BonifBáculo)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 80% del original
            End If
        End If
    End If
    '[Misery_Ezequiel 10/06/05]
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.04  'laud magico de los bardos
    Else
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = ANILLOMAGICODRUIDA Then
            daño = daño * 1.04  'anillo mágico de los druidas
        End If
    End If
    '[\]Misery_Ezequiel 10/06/05]
    Call InfoHechizo(UserIndex)
    b = True
    Call NpcAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    Senddata ToIndex, UserIndex, 0, "||Le has causado " & daño & " puntos de daño a la criatura!" & FONTTYPE_FIGHT
    Call CalcularDarExp(UserIndex, NpcIndex, daño)
    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, UserIndex)
    End If
End If
End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)
    Dim H As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(H).WAV)
    If UserList(UserIndex).flags.TargetUser > 0 Then
        
        If Not UserList(UserList(UserIndex).flags.TargetUser).flags.Meditando Then Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserList(UserIndex).flags.TargetUser).Char.charindex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call Senddata(ToPCArea, UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Map, "CFX" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    End If
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call Senddata(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name & FONTTYPE_FIGHT)
            Call Senddata(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call Senddata(ToIndex, UserIndex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim H As Integer
Dim daño As Integer
Dim tempChr As Integer
    
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tempChr = UserList(UserIndex).flags.TargetUser
'Hambre
If Hechizos(H).SubeHam = 1 Then
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    Call AddtoVar(UserList(tempChr).Stats.MinHam, _
         daño, UserList(tempChr).Stats.MaxHam)
    If UserIndex <> tempChr Then
        Call Senddata(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    Call EnviarHambreYsed(tempChr)
    b = True
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    Else
        '[Wizard] No se puede tirar hambre a si mismo
        'Tipo que sino seria re ghandi! :D
        Call Senddata(ToIndex, UserIndex, 0, "Y145")
        Exit Sub
    End If
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    If UserIndex <> tempChr Then
        Call Senddata(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    Call EnviarHambreYsed(tempChr)
    b = True
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
End If
'Sed
If Hechizos(H).SubeSed = 1 Then
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinAGU, daño, _
         UserList(tempChr).Stats.MaxAGU)
    If UserIndex <> tempChr Then
      Call Senddata(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
      Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call Senddata(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeSed = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    If UserIndex <> tempChr Then
        Call Senddata(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    b = True
End If
' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then

   If UserList(tempChr).flags.Muerto = 1 Then
     Exit Sub
    End If
     
   If UserIndex <> tempChr Then
    If Not Criminal(UserIndex) And Criminal(tempChr) Then
        If UserList(UserIndex).flags.Seguro = True Then
            Call Senddata(ToIndex, UserIndex, 0, "Y277")
            Exit Sub
        Else
            VolverCriminal (UserIndex)
        End If
    End If
End If
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 1200
'[Misery_Ezequiel 12/06/05]
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), daño, MAXATRIBUTOS)
    If UserList(tempChr).Stats.UserAtributos(Agilidad) > 2 * UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) Then UserList(tempChr).Stats.UserAtributos(Agilidad) = 2 * UserList(tempChr).Stats.UserAtributosBackUP(Agilidad)
    UserList(tempChr).flags.DuracionEfecto = 1200

'[\]Misery_Ezequiel 12/06/05]
    'Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), daño, MAXATRIBUTOS)
    UserList(tempChr).flags.TomoPocion = True
    b = True
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    Call InfoHechizo(UserIndex)
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Agilidad) = UserList(tempChr).Stats.UserAtributos(Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Agilidad) = MINATRIBUTOS
    b = True
End If
' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
   
   If UserList(tempChr).flags.Muerto = 1 Then
     Exit Sub
     End If
  
  If tempChr <> UserIndex Then
    If Not Criminal(UserIndex) And Criminal(tempChr) Then
        If UserList(UserIndex).flags.Seguro = True Then
            Call Senddata(ToIndex, UserIndex, 0, "Y277")
            Exit Sub
        Else
            VolverCriminal (UserIndex)
        End If
    End If
End If

    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 1200
'[Misery_Ezequiel 12/06/05]
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), daño, MAXATRIBUTOS)
    If UserList(tempChr).Stats.UserAtributos(Fuerza) > 2 * UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) Then UserList(tempChr).Stats.UserAtributos(Fuerza) = 2 * UserList(tempChr).Stats.UserAtributosBackUP(Fuerza)
'[\]Misery_Ezequiel 12/06/05]
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).flags.TomoPocion = True
    b = True
ElseIf Hechizos(H).SubeFuerza = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    Call InfoHechizo(UserIndex)
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Fuerza) = UserList(tempChr).Stats.UserAtributos(Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Fuerza) = MINATRIBUTOS
    b = True
End If
'Salud
If Hechizos(H).SubeHP = 1 Then
    If UserList(tempChr).Stats.MinHP >= UserList(tempChr).Stats.MaxHP Then 'Gorlok - No curar si esta/s con toda la vida.
        If UserIndex <> tempChr Then
            Call Senddata(ToIndex, UserIndex, 0, "Y284")
        Else
            Call Senddata(ToIndex, UserIndex, 0, "Y372")
        End If
        Exit Sub
    End If
    
    If UserIndex <> tempChr Then
      If Not Criminal(UserIndex) And Criminal(tempChr) Then
        If UserList(UserIndex).flags.Seguro = True Then
            Call Senddata(ToIndex, UserIndex, 0, "Y277")
            Exit Sub
        Else
            VolverCriminal (UserIndex)
        End If
    End If
    End If
  
    If UserList(tempChr).flags.Muerto = 1 Then Exit Sub
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
     'marche
     ' If Hechizos(H).StaffAffected Then
      '  If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
       '     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        '    daño = 200
         '       daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
          '  Else
           ' daño = 200
            '    daño = daño * 0.7 'Baja daño a 70% del original
           ' End If
        'End If
    'End If
    '[Misery_Ezequiel 10/06/05]
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.04  'laud magico de los bardos
    End If
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = ANILLOMAGICODRUIDA Then
        daño = daño * 1.04  'anillo mágico de los druidas
    End If

    If daño < 0 Then daño = 2

    '[\]Misery_Ezequiel 10/06/05]
    Call SubirSkill(tempChr, 1)
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinHP, daño, _
         UserList(tempChr).Stats.MaxHP)
    If UserIndex <> tempChr Then
        Call Senddata(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then

        If UserIndex <> tempChr Then
        Else
        Call Senddata(ToIndex, UserIndex, 0, "Y145")
        Exit Sub
        End If
        
        


    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
            daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
     If UserList(tempChr).flags.Meditando = True Then
        UserList(tempChr).flags.Meditando = False
        UserList(UserIndex).Counters.bPuedeMeditar = False
        Call Senddata(ToIndex, tempChr, 0, "Y377")
        Call Senddata(ToIndex, tempChr, 0, "MEDOK")
     End If
     
     'marche
      If Hechizos(H).StaffAffected Then
        If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    '[Misery_Ezequiel 10/06/05]
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.04  'laud magico de los bardos
    End If
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = ANILLOMAGICODRUIDA Then
        daño = daño * 1.04  'laud magico de los bardos
    End If
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax + 1)
    End If
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax + 1)
    End If
    daño = daño - (daño * (UserList(tempChr).Stats.UserSkills(1) / 2000)) 'Resistencia magica

    If daño < 0 Then daño = 0
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    '[\]Misery_Ezequiel 10/06/05]
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    Call InfoHechizo(UserIndex)
    Call SubirSkill(tempChr, 1)
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    Call Senddata(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
    Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        'Call UserDie(tempChr)
    End If
    b = True
End If
'Mana
If Hechizos(H).SubeMana = 1 Then
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinMAN, daño, UserList(tempChr).Stats.MaxMAN)
    If UserIndex <> tempChr Then
        Call Senddata(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    Call InfoHechizo(UserIndex)
    If UserIndex <> tempChr Then
        Call Senddata(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
End If
'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinSta, daño, _
         UserList(tempChr).Stats.MaxSta)
    If UserIndex <> tempChr Then
         Call Senddata(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
         Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    Call InfoHechizo(UserIndex)
    If UserIndex <> tempChr Then
        Call Senddata(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If
'[Misery_Ezequiel 26/06/05]
' <--------Ilimited All---------->
If Hechizos(H).AgiUpAndFuer = 1 Then
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinAgiFuer, Hechizos(H).MaxAgiFuer)
    UserList(tempChr).flags.DuracionEfecto = 1200
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), daño, MAXATRIBUTOS)
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), daño, MAXATRIBUTOS)
    UserList(tempChr).flags.TomoPocion = True
    b = True
ElseIf Hechizos(H).AgiUpAndFuer = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    Call InfoHechizo(UserIndex)
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgiFuer, Hechizos(H).MaxAgiFuer)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Fuerza) = UserList(tempChr).Stats.UserAtributos(Fuerza) - daño
    UserList(tempChr).Stats.UserAtributos(Agilidad) = UserList(tempChr).Stats.UserAtributos(Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Fuerza) = MINATRIBUTOS
    If UserList(tempChr).Stats.UserAtributos(Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Agilidad) = MINATRIBUTOS
    b = True
End If
'[\]Misery_Ezequiel 26/06/05]
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'Call LogTarea("Sub UpdateUserHechizos")
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then
    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If
Else
'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS
        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If
Next LoopC
End If
End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
'Call LogTarea("ChangeUserHechizo")
UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    Call Senddata(ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)
Else
    Call Senddata(ToIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")
End If
End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub
Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y149")
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call Senddata(ToIndex, UserIndex, 0, "Y149")
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

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!
If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call Senddata(ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call Senddata(ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "CFX" & Npclist(TargetNPC).Char.charindex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
End If
End Sub

'[Misery_Ezequiel 26/06/05]
Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer

    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.charindex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If
End Sub
'[\]Misery_Ezequiel 26/06/05]
'********************Misery_Ezequiel 28/05/05********************'
