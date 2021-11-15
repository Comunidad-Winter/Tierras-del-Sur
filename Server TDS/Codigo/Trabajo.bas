Attribute VB_Name = "Trabajo"
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
'[Misery_Ezequiel 05/06/05]
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 99 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Suerte = 5
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) = 100 Then
                    Exit Sub
End If
'[\]Misery_Ezequiel 05/06/05]

If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then Suerte = Suerte + 50
res = RandomNumber(1, Suerte)

If res > 9 Then
   UserList(UserIndex).flags.Oculto = 0
   UserList(UserIndex).flags.Invisible = 0
   'no hace falta encriptar este (se jode el gil que bypassea esto)
   Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
   Call Senddata(ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
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
'[Misery_Ezequiel 05/06/05]
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 99 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Suerte = 9
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) = 100 Then
                    Suerte = 7
End If
'[\]Misery_Ezequiel 05/06/05]

If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then Suerte = Suerte + 50
res = RandomNumber(1, Suerte)

If UCase$(UserList(UserIndex).Clase) = "CAZADOR" And (UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 612) Then Suerte = 0

If res <= 5 Then
   UserList(UserIndex).flags.Oculto = 1
   UserList(UserIndex).flags.Invisible = 1
  ' If EncriptarProtocolosCriticos Then
   '     Call SendCryptedData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",1")
  ' Else
        Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",1")
   'End If
   Call Senddata(ToIndex, UserIndex, 0, "Y93")
   Call SubirSkill(UserIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
      Call Senddata(ToIndex, UserIndex, 0, "Y1")
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
'aca essssssssssssss
' marche
Dim ModNave As Long
ModNave = ModNavegacion(UserList(UserIndex).Clase)
'If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
 'Call SendData(ToIndex, UserIndex, 0, "||No puedes dejar de navegar en mitad del mar." & FONTTYPE_INFO)
 'Exit Sub
'End If
'[Wizard 03/09/05]=> Repite el mismo codigo que ya se emplea en el UseInvItem(Por esto a pesar de que habian puesto que le pescador navegue sin restriccion) no lo hacia.
'If UserList(UserIndex).Stats.ELV < 25 And UCase$(UserList(UserIndex).Clase) <> "PIRATA" And UserList(UserIndex).Clase <> "Pescador" Then
'   Call Senddata(ToIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
'    Exit Sub
'End If
'[/Wizard]

If UserList(UserIndex).Stats.UserSkills(Navegacion) / ModNave < Barco.MinSkill Then
    Call Senddata(ToIndex, UserIndex, 0, "Y94")
    Call Senddata(ToIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
    Exit Sub
End If
UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
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
     If Not ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, False) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, False) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, False) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)) And _
            UserList(UserIndex).flags.Navegando = 1) Then
              Call Senddata(ToIndex, UserIndex, 0, "||¡¡No puedes dejar de navegar en mitad del mar!!." & FONTTYPE_INFO)
    Exit Sub
    End If
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
Call Senddata(ToIndex, UserIndex, 0, "NAVEG")
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer, Optional Cantidad As Integer)
If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
    If ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).ObjType = OBJTYPE_MINERALES And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(Mineria) / ModFundicion(UserList(UserIndex).Clase) Then
        Call DoLingotes(UserIndex, Cantidad)
        UserList(UserIndex).flags.StartWalk = UserList(UserIndex).flags.StartWalk + 1
    Else
    Call DejarDeTrabajar(UserIndex)
    Call Senddata(ToIndex, UserIndex, 0, "Y95")
    Exit Sub
    End If
End If
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")
Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).Amount
    End If
Next i
If cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")
Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Call Desequipar(UserIndex, i)
        UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - cant
        If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
            cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
            UserList(UserIndex).Invent.Object(i).Amount = 0
            UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        Else
            cant = 0
        End If
        Call UpdateUserInv(False, UserIndex, i)
        If (cant = 0) Then
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
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)
    '[eLwE 19/05/05]Leña de tejo
    If ObjData(ItemIndex).MaderaT > 0 Then Call QuitarObjetos(Leña_tejo, ObjData(ItemIndex).MaderaT, UserIndex)
    '[\]eLwE 19/05/05]
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex) Then
                Call Senddata(ToIndex, UserIndex, 0, "Y96")
                CarpinteroTieneMateriales = False
                Else
                CarpinteroTieneMateriales = True
                End If
    End If
    'Se fija si tiene madera de tejo.
If ObjData(ItemIndex).MaderaT > 0 Then
          ' se fija si puede laburar con eso cagada
        If UserList(UserIndex).Stats.UserAtributos(Inteligencia) * UserList(UserIndex).Stats.UserSkills(Magia) >= 525 Then
       
            If Not TieneObjetos(Leña_tejo, ObjData(ItemIndex).MaderaT, UserIndex) Then
            Call Senddata(ToIndex, UserIndex, 0, "Y96")
            CarpinteroTieneMateriales = False
            Else
            CarpinteroTieneMateriales = True
            End If
        
        Else
            Call Senddata(ToIndex, UserIndex, 0, "Y354")
            CarpinteroTieneMateriales = False
        End If
End If
End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y97")
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y98")
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    '[Misery_Ezequiel 11/06/05]
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y99")
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    '[\]Misery_Ezequiel 11/06/05]
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
        Call Senddata(ToIndex, UserIndex, 0, "Y100")
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ESCUDO Then
        Call Senddata(ToIndex, UserIndex, 0, "Y101")
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_CASCO Then
        Call Senddata(ToIndex, UserIndex, 0, "Y102")
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ARMOUR Then
        Call Senddata(ToIndex, UserIndex, 0, "Y103")
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MARTILLOHERRERO)
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
    'Call Senddata(ToIndex, UserIndex, 0, "Y104")
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If
'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = True
'[/Barrin 30-11-03]
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
            MineralesParaLingote = 10000
    End Select
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer, Optional Cantidad As Integer)
  Dim Slot As Integer
Dim obji As Integer

    Slot = UserList(UserIndex).Invent.HerramientaEqpSlot
    obji = UserList(UserIndex).Invent.HerramientaEqpObjIndex


    If (UserList(UserIndex).Invent.Object(Slot).Amount) < MineralesParaLingote(obji) * Cantidad Or _
        ObjData(obji).ObjType <> OBJTYPE_MINERALES Then
            Call Senddata(ToIndex, UserIndex, 0, "Y105")
             UserList(UserIndex).Invent.HerramientaEqpSlot = 0
             UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
            Call DejarDeTrabajar(UserIndex)
            Exit Sub
    End If
    
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - (MineralesParaLingote(obji) * Cantidad)
    If UserList(UserIndex).Invent.Object(Slot).Amount < 1 Then
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    End If
   
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.Amount = Cantidad
    If Cantidad = 1 Then
    Call Senddata(ToIndex, UserIndex, 0, "Y106")
    Else
    'EnviarPaquete Paquetes.MensajeCompuesto, Chr$(20) & Cantidad, UserIndex
    End If
    MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)

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
    Case Else
        ModDomar = 9
End Select
End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long
CalcularPoderDomador = (UserList(UserIndex).Stats.UserSkills(Domar) * UserList(UserIndex).Stats.UserAtributos(Carisma)) / ModDomar(UserList(UserIndex).Clase)
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
        Call Senddata(ToIndex, UserIndex, 0, "Y108")
        Exit Sub
    End If
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y109")
        Exit Sub
    End If
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(UserIndex) Then
        Dim Index As Integer
        
        If Int(RandomNumber(0, 3)) = 2 Then
        UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
        Index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(Index) = NpcIndex
        UserList(UserIndex).MascotasType(Index) = Npclist(NpcIndex).Numero
        Npclist(NpcIndex).MaestroUser = UserIndex
        '[Wizard]
        Npclist(NpcIndex).PrevMap = Npclist(NpcIndex).Pos.Map
        '[Misery_Ezequiel 11/07/05]
        'MapByEze = Npclist(NpcIndex).Pos.Map
        'XByEze = Npclist(NpcIndex).Pos.X
        'YByEze = Npclist(NpcIndex).Pos.Y
        'Call Senddata(ToAll, UserIndex, 0, "||Mapa: " & MapByEze & ".  x: " & XByEze & ".  y: " & YByEze & FONTTYPE_PARTY)
        '[\]Misery_Ezequiel 11/07/05]
        Call FollowAmo(NpcIndex)
        Call Senddata(ToIndex, UserIndex, 0, "Y110")
        Call SubirSkill(UserIndex, Domar)
        Else
        Call Senddata(ToIndex, UserIndex, 0, "Y111")
        End If
    Else
            '[Misery_Ezequiel 10/07/05]
            'Aunque le diga que no ha logrado domar a la craitura,
            'que suba igual el skill domar...
            Call SubirSkill(UserIndex, Domar)
            '[\]Misery_Ezequiel 10/07/05]
          '[CDT 17-02-2004]
          If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y111")
            UserList(UserIndex).flags.UltimoMensaje = 5
          End If
          '[/CDT]
    End If
Else
    Call Senddata(ToIndex, UserIndex, 0, "Y112")
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

If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
    Call Senddata(ToIndex, UserIndex, 0, "||¡¡No puedes hacer fagatas en zonas seguras!!" + FONTTYPE_INFO)
    Exit Sub
End If
If Not LegalPos(Map, X, Y) Then Exit Sub
If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
    Call Senddata(ToIndex, UserIndex, 0, "Y113")
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
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount / 3
    If Obj.Amount > 1 Then
        Call Senddata(ToIndex, UserIndex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    Else
        Call Senddata(ToIndex, UserIndex, 0, "Y114")
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
        Call Senddata(ToIndex, UserIndex, 0, "Y115")
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

Call TieneEnergia(UserIndex)

res = RandomNumber(1, UserList(UserIndex).Suerte)
If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.Amount = 1
    '[Wizard] El pescador, obtiene 1 de cada 10 peces raro,
    'y si navega 1 de cada 5.
    If UserList(UserIndex).Clase = "Pescador" Then
        If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).flags.Pescaditos = 4 Then
            MiObj.ObjIndex = RandomNumber(544, 546)
            UserList(UserIndex).flags.Pescaditos = 0
        ElseIf UserList(UserIndex).flags.Pescaditos = 9 Then
            MiObj.ObjIndex = RandomNumber(544, 546)
            UserList(UserIndex).flags.Pescaditos = 0
        Else
            MiObj.ObjIndex = Pescado
            UserList(UserIndex).flags.Pescaditos = UserList(UserIndex).flags.Pescaditos + 1
        End If
    Else 'No es pescador
        MiObj.ObjIndex = Pescado
    End If
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
    Call Senddata(ToIndex, UserIndex, 0, "Y116")
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
       Call Senddata(ToIndex, UserIndex, 0, "Y117")
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If
Call SubirSkill(UserIndex, Pesca)
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
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        Call Senddata(ToIndex, UserIndex, 0, "Y118")
    Else
        Call Senddata(ToIndex, UserIndex, 0, "Y119")
    End If
    Call SubirSkill(UserIndex, Pesca)
End If
Exit Sub
errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
If MapInfo(UserList(VictimaIndex).Pos.Map).Pk = 1 Then Exit Sub
If UserList(VictimaIndex).Stats.MinAGU = 0 Or UserList(VictimaIndex).Stats.MinHam = 0 Then Exit Sub
If UserList(LadrOnIndex).flags.Seguro Then
    Call Senddata(ToIndex, LadrOnIndex, 0, "Y120")
    Exit Sub
End If
If UserList(LadrOnIndex).Stats.MinSta < 6 Then
    Call Senddata(ToIndex, LadrOnIndex, 0, "Y167")
    Exit Sub
Else
    Call QuitarSta(LadrOnIndex, 6)
End If



If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
If UserList(VictimaIndex).flags.Privilegios > 0 Then Exit Sub
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
'[Misery_Ezequiel 05/06/05]
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 99 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 91 Then
                        Suerte = 7
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) = 100 Then
                        Suerte = 5
    End If

    res = RandomNumber(1, Suerte)
    If res < 3 Then 'Exito robo
    Dim Mayor As Boolean
    Dim N As Integer
    
If UCase$(UserList(LadrOnIndex).Clase) = "LADRON" And (RandomNumber(1, 50) < 20) Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene objetos." & FONTTYPE_INFO)
            End If
     Else
     
    If UserList(LadrOnIndex).Stats.UserSkills(Robar) < 10 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                      N = RandomNumber(50, 150)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                      Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                      Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 10 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 20 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                      N = RandomNumber(300, 450)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                      Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                      Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 20 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 30 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                      N = RandomNumber(600, 750)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                      Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                      Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 30 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 40 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                      N = RandomNumber(900, 1050)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                      Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                      Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 40 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 50 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                      N = RandomNumber(1200, 1350)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                      Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                      Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 50 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 60 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                         N = RandomNumber(1500, 1650)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                         Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                         Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 60 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 70 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                         N = RandomNumber(1800, 1950)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                         Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                         Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 70 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 80 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                         N = RandomNumber(2100, 2250)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                         Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                         Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 80 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 90 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                          N = RandomNumber(2400, 2550)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                          Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                          Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 90 And UserList(LadrOnIndex).Stats.UserSkills(Robar) < 100 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                          N = RandomNumber(2700, 2850)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                          Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                          Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) = 100 And UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                          N = RandomNumber(3000, 3500)
            If UserList(VictimaIndex).Stats.GLD = 0 Then
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
                Exit Sub
            End If
    If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                          Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                          Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
        End If
End If

If UCase$(UserList(LadrOnIndex).Clase) <> "LADRON" Then
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                N = RandomNumber(1, 100)
                If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, N, MAXORO)
                Call Senddata(ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
            Else
                Call Senddata(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
            End If
        End If
    Else
        Call Senddata(ToIndex, LadrOnIndex, 0, "Y121")
        Call Senddata(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call Senddata(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " es un criminal!" & FONTTYPE_INFO)
    End If
    If Not Criminal(LadrOnIndex) Then
            Call VolverCriminal(LadrOnIndex)
    End If
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)
    Call AddtoVar(UserList(LadrOnIndex).Reputacion.LadronesRep, vlLadron, MAXREP)
    Call SubirSkill(LadrOnIndex, Robar)
End If

Call SendUserStatsBox(VictimaIndex)
Call Senddata(ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "TW" & SOUND_SWING)
Call SendUserStatsBox(LadrOnIndex)
Call Senddata(ToPCArea, LadrOnIndex, UserList(LadrOnIndex).Pos.Map, "TW" & SOUND_SWING)
End Sub
'[\]Misery_Ezequiel 05/07/05]

Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.
Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex
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
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
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
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
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
   '[eLwE 19/05/05]
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    Select Case UserList(LadrOnIndex).Stats.UserSkills(Robar)
        Case Is < 30
            If EsMineral(MiObj.ObjIndex) Then
                num = 100
            Else
                num = RandomNumber(5, 10)
            End If
        Case Is < 50
            If EsMineral(MiObj.ObjIndex) Then
                num = 200
            Else
                num = RandomNumber(20, 25)
            End If
        Case Is < 60
            If EsMineral(MiObj.ObjIndex) Then
                num = 300
            Else
                num = RandomNumber(50, 55)
            End If
        Case Is < 90
            If EsMineral(MiObj.ObjIndex) Then
                num = 400
            Else
                num = 60
            End If
        Case 100
            If EsMineral(MiObj.ObjIndex) Then
                num = 500
            Else
                num = 70
            End If
        Case Else
            If EsMineral(MiObj.ObjIndex) Then
                num = 100
            Else
                num = RandomNumber(5, 10)
            End If
    End Select
    '[\]eLwE 19/05/05]
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
    MiObj.Amount = num
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    Call Senddata(ToIndex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & FONTTYPE_INFO)
Else
    Call Senddata(ToIndex, LadrOnIndex, 0, "Y122")
End If
End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= -1 Then
                    Suerte = 200
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 11 Then
                    Suerte = 190
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 21 Then
                    Suerte = 180
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 31 Then
                    Suerte = 170
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 41 Then
                    Suerte = 160
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 51 Then
                    Suerte = 150
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 61 Then
                    Suerte = 140
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 71 Then
                    Suerte = 130
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 81 Then
                    Suerte = 120
'[Misery_Ezequiel 05/06/05]
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 99 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 91 Then
                    Suerte = 110
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) = 100 Then
                    Suerte = 105
End If

If UCase$(UserList(UserIndex).Clase) = "ASESINO" Then
    res = RandomNumber(0, Suerte * 1.1)
ElseIf UCase$(UserList(UserIndex).Clase) = "CAZADOR" Then
    res = RandomNumber(0, Suerte * 1.7)
ElseIf UCase$(UserList(UserIndex).Clase) = "GUERRERO" Then
    res = RandomNumber(0, Suerte * 1.8)
Else
    res = RandomNumber(0, Suerte * 1.25)
End If

If res <= 16 Then
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Int(daño * 1.5)
        Call Senddata(ToIndex, UserIndex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & Int(daño * 1.5) & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, UserIndex, 0, "||Tu golpe total es de " & Int(daño * 2.5) & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(UserIndex).Name & " por " & Int(daño * 1.5) & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, VictimUserIndex, 0, "||Su golpe total ha sido " & Int(daño * 2.5) & FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call Senddata(ToIndex, UserIndex, 0, "||Has apuñalado la criatura por " & Int(daño * 2) & FONTTYPE_FIGHT)
        Call Senddata(ToIndex, UserIndex, 0, "||Tu golpe total es de " & Int(daño * 3) & FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, Int(daño * 2))
    End If
'[\]Misery_Ezequiel 05/07/05]
Else
    Call Senddata(ToIndex, UserIndex, 0, "Y123")
End If
End Sub
Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal Arbol As Integer)
On Error GoTo errhandler
'Dim Suerte As Integer
Dim res As Integer


If UserList(UserIndex).Clase = "Leñador" Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

Call TieneEnergia(UserIndex)

res = RandomNumber(1, UserList(UserIndex).Suerte)
If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(UserIndex).Clase = "Leñador" Then
        MiObj.Amount = RandomNumber(1, 5)
    Else
        MiObj.Amount = 1
    End If

    If Arbol = 634 Then 'Arbol de tejo
        MiObj.ObjIndex = Leña_tejo
    Else
        MiObj.ObjIndex = Leña
    End If

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call Senddata(ToIndex, UserIndex, 0, "Y124")
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y125")
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Talar)
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
    Call Senddata(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "CH8," & UserList(UserIndex).Char.charindex)
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

Call TieneEnergia(UserIndex)
res = RandomNumber(1, UserList(UserIndex).Suerte)
If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    If UCase$(ObjData(MiObj.ObjIndex).Name) = UCase$("Oro") And UCase$(ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Name) <> UCase$("Piquete de oro") Then
    Call Senddata(ToIndex, UserIndex, 0, "||Para minar oro necesitas un piquete de oro.~65~190~156~0~0")
    Call DejarDeTrabajar(UserIndex)
    Exit Sub
    End If
    If UserList(UserIndex).Clase = "Minero" Then
        MiObj.Amount = RandomNumber(1, 6)
    Else
        MiObj.Amount = 1
    End If
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
       ' EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_MINERO), UserIndex, ToPCArea
         Call Senddata(ToIndex, UserIndex, 0, "Y127")
Else
    If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y127")
        UserList(UserIndex).flags.UltimoMensaje = 9
    End If
End If
Call SubirSkill(UserIndex, Mineria)
Exit Sub
errhandler:
    Call LogError("Error en Sub DoMineria")
End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
UserList(UserIndex).Counters.IdleCount = 0
Dim Suerte As Integer
Dim res As Integer
Dim cant As Integer

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
    Call Senddata(ToIndex, UserIndex, 0, "Y128")
    Call Senddata(ToIndex, UserIndex, 0, "MEDOK")
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & 0 & "," & 0)
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
'[Misery_Ezequiel 05/06/05]
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 99 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 91 Then
                    Suerte = 8
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) = 100 Then
                    Suerte = 5
End If
'[\]Misery_Ezequiel 05/06/05]
res = RandomNumber(1, Suerte)
If res = 1 Then
    cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3)
    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, cant, UserList(UserIndex).Stats.MaxMAN)
    Call Senddata(ToIndex, UserIndex, 0, "||¡Has recuperado " & cant & " puntos de mana!" & FONTTYPE_INFO)
    Call SendUserMana(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If
End Sub

Sub VolverCriminal2(ByVal UserIndex As Integer)
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
If UserList(UserIndex).flags.Privilegios < 2 Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    UserList(UserIndex).Faccion.ArmadaReal = 1
    UserList(UserIndex).Faccion.CiudadanosMatados = 1
    Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If
End Sub

Sub VolverCiudadano2(ByVal UserIndex As Integer)
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Faccion.CiudadanosMatados = 0
Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlASALTO, MAXREP)
End Sub

'[eLwE 23/05/05]Elwe mira si lo de arriba va o no va ( no sabia si borrarlo o no)Nacho xD
Private Function EsMineral(ObjIndex As Integer) As Boolean
    If ObjIndex = OBJTYPE_MINERALES Then
        EsMineral = True
        Exit Function
    End If
    EsMineral = False
End Function
'[\]eLwE 23/05/05]

'[Misery_Ezequiel 15/06/05]
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
        Call Senddata(ToIndex, UserIndex, 0, "Y346")
        Call Senddata(ToIndex, VictimIndex, 0, "||Tu oponente te ha desarmado!" & FONTTYPE_FIGHT)
End If
End Sub
Public Sub TieneEnergia(UserIndex As Integer)
If UserList(UserIndex).Stats.MinSta <= 0 Then
Call DejarDeTrabajar(UserIndex)
End If
End Sub

Public Sub DejarDeTrabajar(UserIndex As Integer)
UserList(UserIndex).TyTrabajo = 0
UserList(UserIndex).TyTrabajoMod = 0
UserList(UserIndex).Suerte = 0
Call Senddata(ToIndex, UserIndex, 0, "DMPT")
UserList(UserIndex).flags.Trabajando = False
End Sub
