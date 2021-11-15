Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
'17/09/02
'Agregue que la función se asegure que el objeto no es un barco
On Error Resume Next
Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).ObjType <> OBJTYPE_LLAVES And _
                ObjData(ObjIndex).ObjType <> OBJTYPE_BARCOS) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    End If
Next i
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador
'Call LogTarea("ClasePuedeUsarItem")
Dim flag As Boolean

If ObjData(ObjIndex).ClaseProhibida(1) <> "" Then
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(UserIndex).Clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
Else
End If
ClasePuedeUsarItem = True
Exit Function
manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next
'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
'[Marche 19-4-04]
If UserList(UserIndex).Pos.Map = 37 Or _
UserList(UserIndex).Pos.Map = 167 Or _
UserList(UserIndex).Pos.Map = 168 Then
'[Marche 19-4-04]
    Dim DeDonde As WorldPos
    Select Case UCase$(UserList(UserIndex).Hogar)
        Case "LINDOS" 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case "ULLATHORPE"
            DeDonde = Ullathorpe
        Case "BANDERBILL"
            DeDonde = Banderbill
        Case "NIX"
        '[Misery_Ezequiel 10/07/05]
            DeDonde = Nix
        Case Else
            DeDonde = Arghâl
        '[\]Misery_Ezequiel 10/07/05]
    End Select
    Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
End If
'[/Barrin]
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
Next

UserList(UserIndex).Invent.NroItems = 0
UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0
UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0
UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0
UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0
UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
UserList(UserIndex).Invent.HerramientaEqpSlot = 0
UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0
UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
On Error GoTo errhandler
If Cantidad > 100000 Then Exit Sub

'SI EL NPC TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        Do While (Cantidad > 0) And (UserList(UserIndex).Stats.GLD > 0)
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If
            MiObj.ObjIndex = iORO
            If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
        Loop
End If
Exit Sub
errhandler:
End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)
'Quita un objeto
UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad
'¿Quedan mas?
If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    UserList(UserIndex).Invent.Object(Slot).Amount = 0
End If
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then
    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If
Else
'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If
    Next LoopC
End If
End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
Dim Obj As Obj
If num > 0 Then
   If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Then
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)
        Obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        '[Misery_Ezequiel 07/06/05]
        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
           Call Senddata(ToIndex, UserIndex, 0, "Y378")
           Exit Sub
        End If
        '[\]Misery_Ezequiel 07/06/05]
        
        Obj.Amount = num
        Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        If ObjData(Obj.ObjIndex).ObjType = OBJTYPE_BARCOS Then
            Call Senddata(ToIndex, UserIndex, 0, "Y150")
        End If
        If ObjData(Obj.ObjIndex).Caos = 1 Or ObjData(Obj.ObjIndex).Real = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y151")
        End If
        If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).Name, False)
  Else
    Call Senddata(ToIndex, UserIndex, 0, "Y152")
  End If
End If
End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)
MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num
If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.ObjIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    Call Senddata(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
End If
End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)
If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then
'Crea un Objeto
    MapData(Map, X, Y).OBJInfo = Obj
    Call Senddata(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)
End If
End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler
'Call LogTarea("MeterItemEnInventario")
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte
'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call Senddata(ToIndex, UserIndex, 0, "Y153")
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
MeterItemEnInventario = True
Call UpdateUserInv(False, UserIndex, Slot)
Exit Function
errhandler:
End Function

Sub GetObj(ByVal UserIndex As Integer)
Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        Obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
        
        If UCase(Obj.Name) = UCase("Pocima sagrada") Then
        Call Senddata(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " tiene la pocima sagrada. Se encuentra en el mapa " & UserList(UserIndex).Pos.Map & "~200~200~200~1~0~" & ENDC)
        End If
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call Senddata(ToIndex, UserIndex, 0, "Y154")
        Else
            'Quitamos el objeto
            Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
        End If
    End If
Else
    Call Senddata(ToIndex, UserIndex, 0, "Y155")
End If
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData

If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If
Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)
Select Case Obj.ObjType
    Case OBJTYPE_WEAPON
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
          If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    '[Misery_Ezequiel 26/06/05]
    Case OBJTYPE_ANILLOS
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0
        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    '[\]Misery_Ezequiel 26/06/05]
    Case OBJTYPE_FLECHAS
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    Case OBJTYPE_HERRAMIENTAS, OBJTYPE_MINERALES
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Case OBJTYPE_ARMOUR
        Select Case Obj.SubTipo
            Case OBJTYPE_ARMADURA
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
                UserList(UserIndex).Invent.ArmourEqpSlot = 0
                Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            Case OBJTYPE_CASCO
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.CascoEqpObjIndex = 0
                UserList(UserIndex).Invent.CascoEqpSlot = 0
                UserList(UserIndex).Char.CascoAnim = NingunCasco
                  If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            Case OBJTYPE_ESCUDO
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
                UserList(UserIndex).Invent.EscudoEqpSlot = 0
                UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                     If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End Select
End Select
Call SendUserStatsBox(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)
End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler
If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "HOMBRE"
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If
Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(UserIndex) Then
        FaccionPuedeUsarItem = UserList(UserIndex).Faccion.ArmadaReal = 1
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(UserIndex) Then
        FaccionPuedeUsarItem = UserList(UserIndex).Faccion.FuerzasCaos = 1
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If
End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler
'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)
If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call Senddata(ToIndex, UserIndex, 0, "Y156")
     Exit Sub
End If

'[Wizard] como no se que objetos puede neceitar va para todos:D
If Obj.SkillM > UserList(UserIndex).Stats.UserSkills(Magia) Then
    Call Senddata(ToIndex, UserIndex, 0, "||Para usar este objeto nececitas " & Obj.SkillM & " skills en magia." & FONTTYPE_INFO)
    Exit Sub
End If

    
'[Misery_Ezequiel 26/06/05]
Select Case Obj.ObjType
    Case OBJTYPE_WEAPON
         'marche
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
        '[Misery_Ezequiel 29/05/05]
        
       If ObjData(ObjIndex).proyectil = 1 Then
            If ObjData(ObjIndex).SkillCombate > UserList(UserIndex).Stats.UserSkills(Proyectiles) Then
            Call Senddata(ToIndex, UserIndex, 0, "||Para usar este arma necesitas " & ObjData(ObjIndex).SkillCombate & " skills en armas de proyectiles." & FONTTYPE_INFO)
            Exit Sub
            End If
        ElseIf ObjData(ObjIndex).Apuñala = 1 Then
            If ObjData(ObjIndex).SkillCombate > UserList(UserIndex).Stats.UserSkills(Apuñalar) Then
            Call Senddata(ToIndex, UserIndex, 0, "||Para usar este arma necesitas " & ObjData(ObjIndex).SkillCombate & " skills en Apuñalar." & FONTTYPE_INFO)
            Exit Sub
            End If
        Else
            If ObjData(ObjIndex).SkillCombate > UserList(UserIndex).Stats.UserSkills(Armas) Then
            Call Senddata(ToIndex, UserIndex, 0, "||Para usar este arma necesitas " & ObjData(ObjIndex).SkillCombate & " skills en Combate con armas." & FONTTYPE_INFO)
            Exit Sub
            End If
        End If
         
         '[\]Misery_Ezequiel 29/05/05]
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                       If UserList(UserIndex).flags.Mimetizado = 1 Then
                        UserList(UserIndex).CharMimetizado.WeaponAnim = NingunArma
                    Else
                        UserList(UserIndex).Char.WeaponAnim = NingunArma
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Or UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                '[Misery_Ezequiel 05/06/05]
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                '[\]Misery_Ezequiel 05/06/05]
                End If
                
                
                'Marche
                Select Case UserList(UserIndex).Clase
                        
                    Case "Asesino"
                    If UCase(Obj.Name) = "KATANA" Or UCase(Obj.Name) = "SABLE" Then
                        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                        End If
                    End If
                        
                    Case "Cazador"
                    If UCase(Obj.Name) = "ESPADA DOS MANOS" Then
                        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                        End If
                    End If
                    
                    Case "Guerrero"
                        If UCase(Obj.Name) = "ESPADA DOS MANOS" Or UCase(Obj.Name) = "ESPADA DE PLATA" Then
                            If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                            End If
                    End If
                    
                  
                        
                    End Select
                
                
                
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.WeaponEqpSlot = Slot
 
                'Sonido
                Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SACARARMA)
               If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
                Else
                    UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
       Else
            Call Senddata(ToIndex, UserIndex, 0, "Y157")
       End If
   Case OBJTYPE_ANILLOS
      If ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    'Animacion por defecto
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
                End If
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.HerramientaEqpSlot = Slot
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call Senddata(ToIndex, UserIndex, 0, "Y286")
       End If
    Case OBJTYPE_HERRAMIENTAS, OBJTYPE_MINERALES
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    'Animacion por defecto
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Or UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                '[Misery_Ezequiel 05/06/05]
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                '[\]Misery_Ezequiel 05/06/05]
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
                End If
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.HerramientaEqpSlot = Slot
                '[Misery_Ezequiel 05/06/05]
                UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                '[\]Misery_Ezequiel 05/06/05]
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call Senddata(ToIndex, UserIndex, 0, "Y286")
       End If
    Case OBJTYPE_FLECHAS
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
       Else
            Call Senddata(ToIndex, UserIndex, 0, "Y286")
       End If
    Case OBJTYPE_ARMOUR
            'marche
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        

         Select Case Obj.SubTipo
            Case OBJTYPE_ARMADURA
            'marche
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                   'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
                        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
            '[Misery_Ezequiel 29/05/05]
            If ObjData(ObjIndex).SkillTacticass > UserList(UserIndex).Stats.UserSkills(Tacticas) Then
                Call Senddata(ToIndex, UserIndex, 0, "||Necesitas " & ObjData(ObjIndex).SkillTacticass & " skills en Tacticas de combate para usar esta Armadura." & FONTTYPE_INFO)
            Exit Sub
            Else
            End If
            '[\]Misery_Ezequiel 29/05/05]
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        Call DarCuerpoDesnudo(UserIndex)
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        Exit Sub
                    End If
                    'Quita el anterior
                    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
                    End If
                   'Lo equipa
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                        
                    If UserList(UserIndex).flags.Mimetizado = 1 Then
                        UserList(UserIndex).CharMimetizado.Body = Obj.Ropaje
                    Else
                        UserList(UserIndex).Char.Body = Obj.Ropaje
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                    UserList(UserIndex).flags.Desnudo = 0
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y158")
                End If
            Case OBJTYPE_CASCO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                 
        '[Misery_Ezequiel 29/05/05]
            If ObjData(ObjIndex).SkillTacticassT > UserList(UserIndex).Stats.UserSkills(Tacticas) Then
                Call Senddata(ToIndex, UserIndex, 0, "||Necesitas " & ObjData(ObjIndex).SkillTacticassT & " skills en Tacticas de combate para usar este casco." & FONTTYPE_INFO)
            Exit Sub
            Else
            End If
            
            
               'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        If UserList(UserIndex).flags.Mimetizado = 1 Then
                            UserList(UserIndex).CharMimetizado.CascoAnim = NingunCasco
                        Else
                            UserList(UserIndex).Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
        '[\]Misery_Ezequiel 29/05/05]
                   ' If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    '    Call Desequipar(UserIndex, Slot)
                     '   UserList(UserIndex).Char.CascoAnim = NingunCasco
                      '  Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                       ' Exit Sub
                    'End If
                    'Quita el anterior
                    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
                    End If
                    'Lo equipa
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.CascoEqpSlot = Slot
                    UserList(UserIndex).Char.CascoAnim = Obj.CascoAnim
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y286")
                End If
            Case OBJTYPE_ESCUDO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                    'Si esta equipado lo quita
            '[Misery_Ezequiel 29/05/05]
                If ObjData(ObjIndex).SkillDefe > UserList(UserIndex).Stats.UserSkills(Defensa) Then
                    Call Senddata(ToIndex, UserIndex, 0, "||Necesitas " & ObjData(ObjIndex).SkillDefe & " skills en Defensa con escudos para usar este escudo." & FONTTYPE_INFO)
                Exit Sub
                Else
                End If
                
            '[\]Misery_Ezequiel 29/05/05]
                     'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        If UserList(UserIndex).flags.Mimetizado = 1 Then
                            UserList(UserIndex).CharMimetizado.ShieldAnim = NingunEscudo
                        Else
                            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    'If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                     '   Call Desequipar(UserIndex, Slot)
                      '  UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                       ' Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        'Exit Sub
                    'End If
                    'Quita el anterior
                    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                    End If
                    
                    
    
                   'marche
                Select Case UserList(UserIndex).Clase
                    Case "Asesino"
                    If UCase(Obj.Name) = UCase("Escudo de Tortuga") Then
                         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                            ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
                            If ObjData(ObjIndex).Name = "Katana" Or UCase(ObjData(ObjIndex).Name) = "SABLE" Then
                            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                            End If
                        End If
                    End If
                        

                    Case "Cazador"
                         If UCase(Obj.Name) = UCase("Escudo de Tortuga") Then
                         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                            ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
                            If UCase(ObjData(ObjIndex).Name) = UCase("Espada dos Manos") Then
                            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                            End If
                        End If
                        End If
                        
                    Case "Guerrero"
                        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                            ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
                            If UCase(ObjData(ObjIndex).Name) = "ESPADA DOS MANOS" Or UCase(ObjData(ObjIndex).Name) = "ESPADA DE PLATA" Then
                            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                            End If
                        End If
    
                        
                End Select
                    
                    'Lo equipa
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.EscudoEqpSlot = Slot
                    If UserList(UserIndex).flags.Mimetizado = 1 Then
                        UserList(UserIndex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
                    Else
                        UserList(UserIndex).Char.ShieldAnim = Obj.ShieldAnim
                        
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y286")
                End If
        End Select
End Select
'[\]Misery_Ezequiel 26/06/05]
'Actualiza
Call UpdateUserInv(True, UserIndex, 0)
Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler
'Verifica si la raza puede usar la ropa
If UserList(UserIndex).Raza = "Humano" Or _
   UserList(UserIndex).Raza = "Elfo" Or _
   UserList(UserIndex).Raza = "Elfo Oscuro" Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If
Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)
End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Usa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(UserIndex).Invent.Object(Slot).Amount = 0 Then Exit Sub
Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)
If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call Senddata(ToIndex, UserIndex, 0, "Y287")
    Exit Sub
End If
'[Misery_Ezequiel 26/06/05]
If Obj.ObjType = 24 Then
    If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
    Else
    Call Senddata(ToIndex, UserIndex, 0, "Y350")
    Exit Sub
    End If
End If

'[\]Misery_Ezequiel 26/06/05]
If Not IntervaloPermiteUsar(UserIndex) Then
    Exit Sub
End If
ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
UserList(UserIndex).flags.TargetObjInvSlot = Slot

Select Case Obj.ObjType
    Case OBJTYPE_USEONCE
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y26")
            Exit Sub
        End If
        'Usa el item
        Call AddtoVar(UserList(UserIndex).Stats.MinHam, Obj.MinHam, UserList(UserIndex).Stats.MaxHam)
        UserList(UserIndex).flags.Hambre = 0
        Call EnviarHambreYsed(UserIndex)
        'Sonido
        Senddata ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_COMIDA
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call UpdateUserInv(False, UserIndex, Slot)

    Case OBJTYPE_GUITA
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y26")
            Exit Sub
        End If
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        Call UpdateUserInv(False, UserIndex, Slot)
        Call SendUserStatsBox(UserIndex)
        
    Case OBJTYPE_WEAPON
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
        End If
        If ObjData(ObjIndex).proyectil = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "T01" & Proyectiles)
        Else
            If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
            TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
            '¿El target-objeto es leña?
            If TargObj.ObjType = OBJTYPE_LEÑA Then
                    If UserList(UserIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                        Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap _
                             , UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
                    Else
                    End If
            End If
        End If
'[Misery_Ezequiel 26/06/05]
    Case OBJTYPE_ANILLOS
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
        End If
        Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Obj.Snd1)
'[\]Misery_Ezequiel 26/06/05]
    Case OBJTYPE_POCIONES
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y26")
            Exit Sub
        End If
'        If UserList(UserIndex).flags.PuedeAtacar = 0 Then
        If Not IntervaloPermiteAtacar(UserIndex, False) Then
            Call Senddata(ToIndex, UserIndex, 0, "Y159")
            Exit Sub
        End If
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(UserIndex).flags.TipoPocion
            Case 1 'Modif la agilidad
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                If UserList(UserIndex).Stats.UserAtributos(Agilidad) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) Then UserList(UserIndex).Stats.UserAtributos(Agilidad) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                'Call SendUserStatsBox(UserIndex)
            Case 2 'Modif la fuerza
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                
                If UserList(UserIndex).Stats.UserAtributos(Fuerza) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) Then UserList(UserIndex).Stats.UserAtributos(Fuerza) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
 
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                AddtoVar UserList(UserIndex).Stats.MinHP, RandomNumber(Obj.MinModificador, Obj.MaxModificador), UserList(UserIndex).Stats.MaxHP
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                Call SendUserVida(UserIndex)
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5), UserList(UserIndex).Stats.MaxMAN)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                Call SendUserMana(UserIndex)
            Case 5 ' Pocion violeta
                If UserList(UserIndex).flags.Envenenado = 1 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call Senddata(ToIndex, UserIndex, 0, "Y160")
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            Case 6  ' Pocion Negra
                If UserList(UserIndex).flags.Privilegios = 0 Then
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UserDiePocionNegra(UserIndex)
                    Call Senddata(ToIndex, UserIndex, 0, "Y161")
                End If
            Case 7 'Pocion de energia
                    Call AddtoVar(UserList(UserIndex).Stats.MinSta, UserList(UserIndex).Stats.MaxSta * 0.1, UserList(UserIndex).Stats.MaxSta)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                    Call SendUserEsta(UserIndex)
       End Select
       Call UpdateUserInv(False, UserIndex, Slot)
     Case OBJTYPE_BEBIDA
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y26")
            Exit Sub
        End If
        AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call EnviarHambreYsed(UserIndex)
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case OBJTYPE_LLAVES
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y26")
            Exit Sub
        End If
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.ObjType = OBJTYPE_PUERTAS Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call Senddata(ToIndex, UserIndex, 0, "Y162")
                        Exit Sub
                     Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y163")
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call Senddata(ToIndex, UserIndex, 0, "Y164")
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y163")
                        Exit Sub
                     End If
                  End If
            Else
                  Call Senddata(ToIndex, UserIndex, 0, "Y165")
                  Exit Sub
            End If
        End If
    
        Case OBJTYPE_BOTELLAVACIA
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
            End If
            If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
                Call Senddata(ToIndex, UserIndex, 0, "Y166")
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            Call UpdateUserInv(False, UserIndex, Slot)
    
        Case OBJTYPE_BOTELLALLENA
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
            End If
            AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
            UserList(UserIndex).flags.Sed = 0
            Call EnviarHambreYsed(UserIndex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            
        Case OBJTYPE_HERRAMIENTAS
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
            End If
            If Not UserList(UserIndex).Stats.MinSta > 0 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y167")
                Exit Sub
            End If
            If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y168")
                Exit Sub
            End If
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlProleta, MAXREP)
            
            Select Case ObjIndex
              '  Case OBJTYPE_CAÑA, RED_PESCA
             '       Call Senddata(ToIndex, UserIndex, 0, "T01" & Pesca)
              '  Case HACHA_LEÑADOR
               '     Call Senddata(ToIndex, UserIndex, 0, "T01" & Talar)
                '[Misery_Ezequiel 27/05/05]
               ' Case HACHA_DORADA
               '     Call Senddata(ToIndex, UserIndex, 0, "T01" & Talar)
                '[\]Misery_Ezequiel 27/05/05]
               ' Case PIQUETE_MINERO
               '     Call Senddata(ToIndex, UserIndex, 0, "T01" & Mineria)
                Case MARTILLO_HERRERO
                    Call Senddata(ToIndex, UserIndex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(UserIndex)
                    Call Senddata(ToIndex, UserIndex, 0, "SFC")
            End Select
        
        Case OBJTYPE_PERGAMINOS
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
            End If
            If UserList(UserIndex).flags.Hambre = 0 And _
               UserList(UserIndex).flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                Call UpdateUserInv(False, UserIndex, Slot)
            Else
               Call Senddata(ToIndex, UserIndex, 0, "Y169")
            End If
       
       Case OBJTYPE_MINERALES
           If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
           End If
           Call Senddata(ToIndex, UserIndex, 0, "T01" & FundirMetal)
       
       Case OBJTYPE_INSTRUMENTOS
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
            End If
            Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Obj.Snd1)
       
       Case OBJTYPE_BARCOS
    'Verifica si esta aproximado al agua antes de permitirle navegar
    '[Misery_Ezequiel 12/06/05]
      If UserList(UserIndex).Stats.ELV < 25 Then
          If UCase$(UserList(UserIndex).Clase) <> "PESCADOR" And UCase$(UserList(UserIndex).Clase) <> "PIRATA" Then
              Call Senddata(ToIndex, UserIndex, 0, "Y33")
            Exit Sub
         End If
      End If
     '[\]Misery_Ezequiel 12/06/05]
        If ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)) And _
            UserList(UserIndex).flags.Navegando = 0) _
            Or UserList(UserIndex).flags.Navegando = 1 Then
           Call DoNavega(UserIndex, Obj, Slot)
        Else
            Call Senddata(ToIndex, UserIndex, 0, "Y170")
        End If
End Select
'Actualiza
'Call SendUserStatsBox(UserIndex)
'Call UpdateUserInv(False, UserIndex, Slot)
End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) \ ModHerreriA(UserList(UserIndex).Clase) Then
        If ObjData(ArmasHerrero(i)).ObjType = OBJTYPE_WEAPON Then
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & " (" & ObjData(ArmasHerrero(i)).MinHIT & "/" & ObjData(ArmasHerrero(i)).MaxHIT & ")" & "," & ArmasHerrero(i) & ","
        Else
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i
Call Senddata(ToIndex, UserIndex, 0, "LAH" & cad$)
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
Dim i As Integer, cad$

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(Carpinteria) / ModCarpinteria(UserList(UserIndex).Clase) Then _
        
        If ObjData(ObjCarpintero(i)).MaderaT > 0 Then
        cad$ = cad$ & ObjData(ObjCarpintero(i)).Name & " (" & ObjData(ObjCarpintero(i)).MaderaT & ")" & "," & ObjCarpintero(i) & ","
        Else
        cad$ = cad$ & ObjData(ObjCarpintero(i)).Name & " (" & ObjData(ObjCarpintero(i)).Madera & ")" & "," & ObjCarpintero(i) & ","
        End If
    End If
Next i
Call Senddata(ToIndex, UserIndex, 0, "OBR" & cad$)
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
Dim i As Integer, cad$

For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).Clase) Then _
        cad$ = cad$ & ObjData(ArmadurasHerrero(i)).Name & " (" & ObjData(ArmadurasHerrero(i)).MinDef & "/" & ObjData(ArmadurasHerrero(i)).MaxDef & ")" & "," & ArmadurasHerrero(i) & ","
Next i
Call Senddata(ToIndex, UserIndex, 0, "LAR" & cad$)
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
Call TirarTodosLosItems(UserIndex)
Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)
End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
ItemSeCae = ObjData(Index).Real <> 1 And _
            ObjData(Index).Caos <> 1 And _
            ObjData(Index).ObjType <> OBJTYPE_LLAVES And _
            ObjData(Index).ObjType <> OBJTYPE_BARCOS And _
            ObjData(Index).NoSeCae = 0
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
'Call LogTarea("Sub TirarTodosLosItems")
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
  ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
  If ItemIndex > 0 Then
         If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
  End If
Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
ItemNewbie = ObjData(ItemIndex).Newbie = 1
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer
'[Wizard 03/09/05] Asi a los ladrones se le cae el oro aunque sean newbies.
'If Criminal(UserIndex) And UserList(UserIndex).Clase = "Ladron" Then
'    Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)
'End If
'[/Wizard]=> Desechado por balance; se decidio lo siguiente:
Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)



If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
For i = 1 To MAX_INVENTORY_SLOTS
  ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
  If ItemIndex > 0 Then
         If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
  End If
Next i
End Sub
'********************Misery_Ezequiel 28/05/05********************'
'Function ClasePuedeUsarHECHIZO(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
'On Error GoTo manejador
'Dim flag As Boolean

'If Hechizo(ObjIndex).ClaseProhibida(1) <> "" Then
 '   Dim i As Integer
  '  For i = 1 To NUMCLASES
   '     If Hechizo(ObjIndex).ClaseProhibida(i) = UCase$(UserList(UserIndex).Clase) Then
    '            ClasePuedeUsarItem = False
  '              Exit Function
     ''   End If
   ' Next i
'Else
'End If
'ClasePuedeUsarItem = True
'Exit Function
'M 'anejador:
  '  LogError ("Error en ClasePuedeUsarItem")
'End Function
