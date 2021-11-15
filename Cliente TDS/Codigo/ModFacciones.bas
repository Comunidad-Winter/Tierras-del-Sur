Attribute VB_Name = "ModFacciones"
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

Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarquía
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos


Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer

Public Const ExpAlUnirse = 100000
Public Const ExpX100 = 5000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No se permiten criminales en el ejercito imperial!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < 10 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 10 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 18 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 18!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.CriminalesMatados \ 100

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al Ejercito Imperial!!!, aqui tienes tu armadura. Por cada centena de criminales que acabes te dare un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
           If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
              UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.ObjIndex = TunicaMagoImperialEnanos
           Else
                  MiObj.ObjIndex = TunicaMagoImperial
           End If
    ElseIf UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Or _
           UCase$(UserList(UserIndex).Clase) = "PALADIN" Or _
           UCase$(UserList(UserIndex).Clase) = "BANDIDO" Or _
           UCase$(UserList(UserIndex).Clase) = "ASESINO" Then
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.ObjIndex = ArmaduraImperial3
              Else
                  MiObj.ObjIndex = ArmaduraImperial1
              End If
    Else
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.ObjIndex = ArmaduraImperial3
              Else
                  MiObj.ObjIndex = ArmaduraImperial2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoReal(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CriminalesMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 crinales mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
UserList(UserIndex).Faccion.ArmadaReal = 0
'Call PerderItemsFaccionarios(UserIndex)
Call SendData(ToIndex, UserIndex, 0, "Y182")
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)
UserList(UserIndex).Faccion.FuerzasCaos = 0
'Call PerderItemsFaccionarios(UserIndex)
Call SendData(ToIndex, UserIndex, 0, "Y183")
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Aprendiz "
    Case 1
        TituloReal = "Escudero"
    Case 2
        TituloReal = "Caballero"
    Case 3
        TituloReal = "Capitan"
    Case 4
        TituloReal = "Teniente"
    Case 5
        TituloReal = "Comandante"
    Case 6
        TituloReal = "Mariscal"
    Case 7
        TituloReal = "Senescal"
    Case 8
        TituloReal = "Protector"
    Case 9
        TituloReal = "Guardian del Bien"
    Case Else
        TituloReal = "Campeón de la Luz"
End Select

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)

If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Largate de aqui, bufon!!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas del caos!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

'[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No permitiré que ningún insecto real ingrese ¡Traidor del Rey!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If
'[/Barrin]

If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ja ja ja tu no eres bienvenido aqui!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados < 150 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 150 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Exit Sub
End If

UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.CiudadanosMatados \ 100

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te dare un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
              '[Barrin 30-11-03] BUG de la túnica de GNOMO del caos!
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.ObjIndex = TunicaMagoCaosEnanos
              Else
                  MiObj.ObjIndex = TunicaMagoCaos
              End If
              '[/Barrin 30-11-03]
    ElseIf UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Or _
           UCase$(UserList(UserIndex).Clase) = "PALADIN" Or _
           UCase$(UserList(UserIndex).Clase) = "BANDIDO" Or _
           UCase$(UserList(UserIndex).Clase) = "ASESINO" Then
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.ObjIndex = ArmaduraCaos3
              Else
                  MiObj.ObjIndex = ArmaduraCaos1
              End If
    Else
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.ObjIndex = ArmaduraCaos3
              Else
                  MiObj.ObjIndex = ArmaduraCaos2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoCaos(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CiudadanosMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 ciudadanos mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
End If


End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Esbirro"
    Case 1
        TituloCaos = "Servidor de las Sombras"
    Case 2
        TituloCaos = "Acólito"
    Case 3
        TituloCaos = "Guerrero Sombrío"
    Case 4
        TituloCaos = "Sanguinario"
    Case 5
        TituloCaos = "Caballero de la Oscuridad"
    Case 6
        TituloCaos = "Condenado"
    Case 7
        TituloCaos = "Heraldo Impío"
    Case 8
        TituloCaos = "Corruptor"
    Case Else
        TituloCaos = "Devorador de Almas"
End Select


End Function

'[Barrin 17-12-03]
'Sub PerderItemsFaccionarios(ByVal UserIndex As Integer)
'Dim i As Byte
'Dim MiObj As Obj
'Dim ItemIndex As Integer
'
'For i = 1 To MAX_INVENTORY_SLOTS
'  ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
'  If ItemIndex > 0 Then
'         If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then
'            Call QuitarUserInvItem(UserIndex, i, UserList(UserIndex).Invent.Object(i).Amount)
'            Call UpdateUserInv(False, UserIndex, i)
'            If ObjData(ItemIndex).ObjType = OBJTYPE_ARMOUR Then
'                If ObjData(ItemIndex).Real = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
'                If ObjData(ItemIndex).Caos = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
'            Else
'                UserList(UserIndex).Faccion.RecibioItemFaccionario = 0
'            End If
'         End If
'
'  End If
'Next i
'
'End Sub
'[/Barrin]
