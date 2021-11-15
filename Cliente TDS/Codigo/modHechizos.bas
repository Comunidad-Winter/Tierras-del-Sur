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


Option Explicit

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.Invisible = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + daño
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Call SendUserStatsBox(val(UserIndex))

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(UserIndex).flags.Privilegios = 0 Then
    
       
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        'marche
        If daño < 0 Then daño = 0
        
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
        
        Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
        Call SendUserStatsBox(val(UserIndex))
        
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

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 Then
        UserList(UserIndex).flags.Paralizado = 1
        UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        If EncriptarProtocolosCriticos Then
          Call SendCryptedData(ToIndex, UserIndex, 0, "PARADOK")
        Else
          Call SendData(ToIndex, UserIndex, 0, "PARADOK")
        End If
        Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y) 'Gorlok
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
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(ToIndex, UserIndex, 0, "Y132")
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "Y133")
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(UserIndex).Char.charindex
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbCyan & "°" & s & "°" & ind)
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
                    Call SendData(ToIndex, UserIndex, 0, "||Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Necesitas un baculo para lanzar este hechizo!!." & FONTTYPE_INFO)
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
                Call SendData(ToIndex, UserIndex, 0, "Y134")
                PuedeLanzar = False
            End If
                
        Else
            Call SendData(ToIndex, UserIndex, 0, "Y135")
            PuedeLanzar = False
        End If
    Else
            Call SendData(ToIndex, UserIndex, 0, "Y136")
            PuedeLanzar = False
    End If
Else
   Call SendData(ToIndex, UserIndex, 0, "Y137")
   PuedeLanzar = False
End If

End Function

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)

'Call LogTarea("HechizoInvocacion")
If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).Cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind <= MAXNPCS Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            Index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
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

Select Case Hechizos(uh).Tipo
    Case uInvocacion '
       Call HechizoInvocacion(UserIndex, b)
End Select

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
       Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNpc, uh, b, UserIndex)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNpc, UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNpc = 0
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
                Call SendData(ToIndex, UserIndex, 0, "Y138")
            End If
        Case uNPC
            If UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "Y139")
            End If
        Case uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "Y140")
            End If
        Case uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = False
'[/Barrin 30-11-03]

End Sub


Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

Dim H As Integer, TU As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
TU = UserList(UserIndex).flags.TargetUser

If Hechizos(H).Invisibilidad = 1 Then
   
     If UserList(TU).flags.Invisible = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡El usuario ya está invisible!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If UserList(TU).flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡Está muerto!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
   
   UserList(TU).flags.Invisible = 1
   If EncriptarProtocolosCriticos Then
      Call SendCryptedData(ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.charindex & ",1")
   Else
    Call SendData(ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.charindex & ",1")
   End If
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
                Call SendData(ToIndex, UserIndex, 0, "||No está envenenado." & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No estás envenenado." & FONTTYPE_INFO)
            End If
            Exit Sub
        End If
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
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Then
    If UserIndex = TU Then
        Call SendData(ToIndex, UserIndex, 0, "Y275")
        Exit Sub 'Prohibido paralizarse a si mismo - byGorlok 2005-03-25
    End If
    If UserList(TU).flags.Paralizado = 0 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        
        UserList(TU).flags.Paralizado = 1
        UserList(TU).Counters.Paralisis = IntervaloParalizado
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(ToIndex, TU, 0, "PARADOK")
        Else
            Call SendData(ToIndex, TU, 0, "PARADOK")
        End If
        Call SendData(ToIndex, TU, 0, "PU" & UserList(TU).Pos.X & "," & UserList(TU).Pos.Y) 'Gorlok
        Call InfoHechizo(UserIndex)
        b = True
    End If
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
                UserList(TU).flags.Paralizado = 0
                'no need to crypt this
                Call SendData(ToIndex, TU, 0, "PARADOK")
                Call InfoHechizo(UserIndex)
                b = True
    End If
End If

If Hechizos(H).Revivir = 1 Then

   If UserList(TU).flags.Muerto = 0 Then
   
                  
            Call SendData(ToIndex, UserIndex, 0, "¡¡El usuario ya está vivo!!" & FONTTYPE_INFO)
            
           b = False
     Exit Sub
      Else

        If Not Criminal(TU) Then
                If TU <> UserIndex Then
                    Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)
                    Call SendData(ToIndex, UserIndex, 0, "Y143")
                End If
        End If
   
        
        
           'revisamos si necesita vara
        If UCase$(UserList(UserIndex).Clase) = "a" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call SendData(ToIndex, UserIndex, 0, "||Necesitas un mejor báculo para este hechizo" & FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
         End If
        
        Call RevivirUsuario(TU)
        
        UserList(TU).Stats.MinAGU = 0
        UserList(TU).Stats.MinHam = 0
        UserList(TU).flags.Hambre = 1
        UserList(TU).flags.Sed = 1
        
    End If
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        Call SendData(ToIndex, TU, 0, "CEGU")
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(ToIndex, TU, 0, "DUMB")
        Else
            Call SendData(ToIndex, TU, 0, "DUMB")
        End If
        Call InfoHechizo(UserIndex)
        b = True
End If

End Sub


Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)

If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "Y144")
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
        Call SendData(ToIndex, UserIndex, 0, "Y144")
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 1
   b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
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
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "Y145")
   End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "Y146")
   End If
End If
'[/Barrin]
 


End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)

Dim daño As Long


'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, daño, Npclist(NpcIndex).Stats.MaxHP)
    Call SendData(ToIndex, UserIndex, 0, "||Has curado " & daño & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "Y144")
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
    
    Call InfoHechizo(UserIndex)
    b = True
    Call NpcAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    SendData ToIndex, UserIndex, 0, "||Le has causado " & daño & " puntos de daño a la criatura!" & FONTTYPE_FIGHT
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
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(H).WAV)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserList(UserIndex).flags.TargetUser).Char.charindex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(ToPCArea, UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).Pos.Map, "CFX" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
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
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
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
      Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
      Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
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
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), daño, MAXATRIBUTOS)
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
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), daño, MAXATRIBUTOS)
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
            Call SendData(ToIndex, UserIndex, 0, "||No está herido." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No estás herido." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
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
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax + 1)
    End If
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHP, daño, _
         UserList(tempChr).Stats.MaxHP)
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If UserIndex = tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "Y148")
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
    Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinMAN, daño, UserList(tempChr).Stats.MaxMAN)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
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
         Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
         Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If


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

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)

Else

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")

End If


End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "Y149")
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(ToIndex, UserIndex, 0, "Y149")
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

