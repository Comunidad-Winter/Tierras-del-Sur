Attribute VB_Name = "Trabajo"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
Option Explicit

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Exit Sub
End If

If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then Suerte = Suerte + 50




res = RandomNumber(1, Suerte)

If res > 9 Then
   UserList(UserIndex).flags.Oculto = 0
   UserList(UserIndex).flags.Invisible = 0
   'no hace falta encriptar este (se jode el gil que bypassea esto)
   Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
   Call SendData(ToIndex, UserIndex, 0, "||�Has vuelto a ser visible!" & FONTTYPE_INFO)
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub
Public Sub DoOcultarse(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Suerte = 7
End If

If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then Suerte = Suerte + 50

res = RandomNumber(1, Suerte)

If res <= 5 Then
   UserList(UserIndex).flags.Oculto = 1
   UserList(UserIndex).flags.Invisible = 1
   If EncriptarProtocolosCriticos Then
        Call SendCryptedData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",1")
   Else
        Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",1")
   End If
   Call SendData(ToIndex, UserIndex, 0, "Y93")
   Call SubirSkill(UserIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
      Call SendData(ToIndex, UserIndex, 0, "Y1")
      UserList(UserIndex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]

Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long
ModNave = ModNavegacion(UserList(UserIndex).Clase)

If UserList(UserIndex).Stats.UserSkills(Navegacion) / ModNave < Barco.MinSkill Then
    Call SendData(ToIndex, UserIndex, 0, "Y94")
    Call SendData(ToIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
    Exit Sub
End If

UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
UserList(UserIndex).Invent.BarcoSlot = Slot

If UserList(UserIndex).flags.Navegando = 0 Then
    
    UserList(UserIndex).Char.Head = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Body = Barco.Ropaje
    Else
        UserList(UserIndex).Char.Body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).flags.Navegando = 1
    
Else
    
    UserList(UserIndex).flags.Navegando = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
            
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If

End If

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendData(ToIndex, UserIndex, 0, "NAVEG")

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
'Call LogTarea("Sub FundirMineral")

If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).ObjType = OBJTYPE_MINERALES And ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(Mineria) / ModFundicion(UserList(UserIndex).Clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call SendData(ToIndex, UserIndex, 0, "Y95")
   End If

End If

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).objIndex = ItemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).objIndex = ItemIndex Then
        
        Call Desequipar(UserIndex, i)
        
        UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - Cant
        If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
            UserList(UserIndex).Invent.Object(i).Amount = 0
            UserList(UserIndex).Invent.Object(i).objIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, UserIndex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Le�a, ObjData(ItemIndex).Madera, UserIndex)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Le�a, ObjData(ItemIndex).Madera, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "Y96")
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "Y97")
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "Y98")
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingP, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "Y99")
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(Herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
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


Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'Call LogTarea("Sub HerreroConstruirItem")
If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(UserIndex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).ObjType = OBJTYPE_WEAPON Then
        Call SendData(ToIndex, UserIndex, 0, "Y100")
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ESCUDO Then
        Call SendData(ToIndex, UserIndex, 0, "Y101")
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_CASCO Then
        Call SendData(ToIndex, UserIndex, 0, "Y102")
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ARMOUR Then
        Call SendData(ToIndex, UserIndex, 0, "Y103")
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.objIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MARTILLOHERRERO)
    
End If

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

If CarpinteroTieneMateriales(UserIndex, ItemIndex) And _
   UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(UserIndex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then

    Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
    Call SendData(ToIndex, UserIndex, 0, "Y104")
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.objIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]

End Sub

Public Sub DoLingotes(ByVal UserIndex As Integer)
'    Call LogTarea("Sub DoLingotes")

    If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount < 5 Or ObjData(UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).objIndex).ObjType <> OBJTYPE_MINERALES Then
        Call SendData(ToIndex, UserIndex, 0, "Y105")
        Exit Sub
    End If
    
    If RandomNumber(1, ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill) < 10 Then
                UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount - 5
                If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount < 1 Then
                    UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0
                    UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).objIndex = 0
                End If
                Call SendData(ToIndex, UserIndex, 0, "Y106")
                Dim nPos As WorldPos
                Dim MiObj As Obj
                MiObj.Amount = 1
                MiObj.objIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
                End If
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).flags.TargetObjInvSlot)
                Call SendData(ToIndex, UserIndex, 0, "||�Has obtenido un lingote!" & FONTTYPE_INFO)
    Else
        
        UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount - 5
        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount < 1 Then
                UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0
                UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).objIndex = 0
        End If
        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).flags.TargetObjInvSlot)
       '[CDT 17-02-2004]
       If Not UserList(UserIndex).flags.UltimoMensaje = 7 Then
         Call SendData(ToIndex, UserIndex, 0, "Y107")
         UserList(UserIndex).flags.UltimoMensaje = 7
       End If
       '[/CDT]
    End If

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]

End Sub

Function ModNavegacion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "PIRATA"
        ModNavegacion = 1
    Case "PESCADOR"
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "MINERO"
        ModFundicion = 1
    Case "HERRERO"
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "CARPINTERO"
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "HERRERO"
        ModHerreriA = 1
    Case "MINERO"
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal Clase As String) As Integer
Select Case UCase$(Clase)
    Case "DRUIDA"
        ModDomar = 6
    Case "CAZADOR"
        ModDomar = 6
    Case "CLERIGO"
        ModDomar = 7
    Case Else
        ModDomar = 10
End Select
End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long
CalcularPoderDomador = _
UserList(UserIndex).Stats.UserAtributos(Carisma) * _
(UserList(UserIndex).Stats.UserSkills(Domar) / ModDomar(UserList(UserIndex).Clase)) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3)
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
Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "Y108")
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "Y109")
        Exit Sub
    End If
    
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(UserIndex) Then
        Dim Index As Integer
        UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
        Index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(Index) = NpcIndex
        UserList(UserIndex).MascotasType(Index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(ToIndex, UserIndex, 0, "Y110")
        Call SubirSkill(UserIndex, Domar)
        
    Else
          '[CDT 17-02-2004]
          If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call SendData(ToIndex, UserIndex, 0, "Y111")
            UserList(UserIndex).flags.UltimoMensaje = 5
          End If
          '[/CDT]
        
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "Y112")
End If
End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
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
    
    
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    
End Sub
Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj

If Not LegalPos(Map, X, Y) Then Exit Sub

If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
    Call SendData(ToIndex, UserIndex, 0, "Y113")
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
    Obj.objIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount / 3
    
    If Obj.Amount > 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "Y114")
    End If
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
    
    Dim Fogatita As New cGarbage
    Fogatita.Map = Map
    Fogatita.X = X
    Fogatita.Y = Y
    Call TrashCollector.Add(Fogatita)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        Call SendData(ToIndex, UserIndex, 0, "Y115")
        UserList(UserIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UserList(UserIndex).Clase = "Pescador" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Pesca) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    MiObj.Amount = 1
    MiObj.objIndex = Pescado
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "Y116")
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
      Call SendData(ToIndex, UserIndex, 0, "Y117")
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Pesca)

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]

Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UCase(UserList(UserIndex).Clase) = "PESCADOR" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(UserIndex).Stats.UserSkills(Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Select Case iSkill
Case 0:         Suerte = 0
Case 1 To 10:   Suerte = 60
Case 11 To 20:  Suerte = 54
Case 21 To 30:  Suerte = 49
Case 31 To 40:  Suerte = 43
Case 41 To 50:  Suerte = 38
Case 51 To 60:  Suerte = 32
Case 61 To 70:  Suerte = 27
Case 71 To 80:  Suerte = 21
Case 81 To 90:  Suerte = 16
Case 91 To 100: Suerte = 11
Case Else:      Suerte = 0
End Select

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.Amount = RandomNumber(1, 5)
        Else
            MiObj.Amount = 1
        End If
        MiObj.objIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call SendData(ToIndex, UserIndex, 0, "Y118")
        
    Else
        Call SendData(ToIndex, UserIndex, 0, "Y119")
    End If
    
    Call SubirSkill(UserIndex, Pesca)
End If

Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

If MapInfo(UserList(VictimaIndex).Pos.Map).Pk = 1 Then Exit Sub
If UserList(LadrOnIndex).flags.Seguro Then
    Call SendData(ToIndex, LadrOnIndex, 0, "Y120")
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).flags.Privilegios = 0 Then
    Dim Suerte As Integer
    Dim res As Integer
    
       
    If UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 91 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
    
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UCase$(UserList(LadrOnIndex).Clase) = "LADRON") Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene objetos." & FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim N As Integer
                
                N = RandomNumber(1, 100)
                
                If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                
                Call SendData(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
            End If
        End If
    Else
        Call SendData(ToIndex, LadrOnIndex, 0, "Y121")
        Call SendData(ToIndex, VictimaIndex, 0, "||�" & UserList(LadrOnIndex).Name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call SendData(ToIndex, VictimaIndex, 0, "||�" & UserList(LadrOnIndex).Name & " es un criminal!" & FONTTYPE_INFO)
    End If

    If Not Criminal(LadrOnIndex) Then
            Call VolverCriminal(LadrOnIndex)
    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    Call AddtoVar(UserList(LadrOnIndex).Reputacion.LadronesRep, vlLadron, MAXREP)
    Call SubirSkill(LadrOnIndex, Robar)

End If


End Sub


Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregu� los barcos
' Esta funcion determina qu� objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).objIndex

ObjEsRobable = _
ObjData(OI).ObjType <> OBJTYPE_LLAVES And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).ObjType <> OBJTYPE_BARCOS

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).objIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).objIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.objIndex = UserList(VictimaIndex).Invent.Object(i).objIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    Call SendData(ToIndex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.objIndex).Name & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, LadrOnIndex, 0, "Y122")
End If

End Sub
Public Sub DoApu�alar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal da�o As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Apu�alar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Apu�alar) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 3 Then
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Int(da�o * 1.5)
        Call SendData(ToIndex, UserIndex, 0, "||Has apu�alado a " & UserList(VictimUserIndex).Name & " por " & Int(da�o * 1.5) & FONTTYPE_FIGHT)
        Call SendData(ToIndex, VictimUserIndex, 0, "||Te ha apu�alado " & UserList(UserIndex).Name & " por " & Int(da�o * 1.5) & FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(da�o * 2)
        Call SendData(ToIndex, UserIndex, 0, "||Has apu�alado la criatura por " & Int(da�o * 2) & FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Apu�alar)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, Int(da�o * 2))
    End If
    
Else
    Call SendData(ToIndex, UserIndex, 0, "Y123")
End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UserList(UserIndex).Clase = "Le�ador" Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLe�ador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Talar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(UserIndex).Clase = "Le�ador" Then
        MiObj.Amount = RandomNumber(1, 5)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.objIndex = Le�a
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "Y124")
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call SendData(ToIndex, UserIndex, 0, "Y125")
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Talar)

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

If UserList(UserIndex).flags.Privilegios < 2 Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlASALTO, MAXREP)

End Sub


Public Sub DoPlayInstrumento(ByVal UserIndex As Integer)

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UserList(UserIndex).Clase = "Minero" Then
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Mineria) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.objIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    
    If UserList(UserIndex).Clase = "Minero" Then
        MiObj.Amount = RandomNumber(1, 6)
    Else
        MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call SendData(ToIndex, UserIndex, 0, "Y126")
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
        Call SendData(ToIndex, UserIndex, 0, "Y127")
        UserList(UserIndex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Mineria)

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]

Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub



Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(UserIndex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(UserIndex).Counters.bPuedeMeditar = False Then
    UserList(UserIndex).Counters.bPuedeMeditar = True
End If
If UserList(UserIndex).Counters.bPuedeMeditar = False Then Exit Sub

If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
    Call SendData(ToIndex, UserIndex, 0, "Y128")
    Call SendData(ToIndex, UserIndex, 0, "MEDOK")
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & 0 & "," & 0)
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Meditar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    Cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3)
    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Cant, UserList(UserIndex).Stats.MaxMAN)
    Call SendData(ToIndex, UserIndex, 0, "||�Has recuperado " & Cant & " puntos de mana!" & FONTTYPE_INFO)
    Call SendUserStatsBox(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If

End Sub




