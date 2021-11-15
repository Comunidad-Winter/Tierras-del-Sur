Attribute VB_Name = "ModFacciones"
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

'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

'[Misery_Ezequiel 11/06/05]
Public ArmaduraImperial1_Hombre As Integer 'Primer jerarquia
Public ArmaduraImperial1_Mujer As Integer 'Primer jerarquia
Public ArmaduraImperial2_Hombre As Integer 'Segunda jerarqu�a
Public ArmaduraImperial2_Mujer As Integer 'Segunda jerarqu�a
Public ArmaduraImperial3_Hombre As Integer 'Enanos
Public ArmaduraImperial3_Mujer As Integer 'Enanos
Public TunicaMagoImperial_Hombre As Integer 'Magos
Public TunicaMagoImperial_Mujer As Integer 'Magos
Public TunicaMagoImperialEnanos_Hombre As Integer 'Magos
Public TunicaMagoImperialEnanos_Mujer As Integer 'Magos
'**************************************************************
Public ArmaduraCaos1_Hombre As Integer
Public ArmaduraCaos1_Mujer As Integer
Public TunicaMagoCaos_Hombre As Integer
Public TunicaMagoCaos_Mujer As Integer
Public TunicaMagoCaosEnanos_Hombre As Integer
Public TunicaMagoCaosEnanos_Mujer As Integer
Public ArmaduraCaos2_Hombre As Integer
Public ArmaduraCaos2_Mujer As Integer
Public ArmaduraCaos3_Hombre As Integer
Public ArmaduraCaos3_Mujer As Integer
'[\]Misery_Ezequiel 11/06/05]

Public Const ExpAlUnirse = 100000
Public Const ExpX100 = 5000

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If Criminal(UserIndex) Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No se permiten criminales en el ejercito imperial!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.CriminalesMatados < 70 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes matar al menos 70 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Stats.ELV < 25 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.CriminalesMatados \ 100

Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Bienvenido a al Ejercito Imperial!!!, aqui tienes tu armadura. Por cada centena de criminales que acabes te dare un recompensa, buena suerte soldado!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    
'[Wizard 03/09/05] no se quien hizo lo que estaba aca, pero por dios mandenlo a un curso de redaccion
'Habia 3 cases diciendo lo mismo, 1 If clause que nunca se accedia por suerte porque si se accedia daba armadura del caos
'ademas usan los Ucase$ para esto, que son cosas que los escribe el codigo y no pueden cambiar, gastan memoria ram al pedo.
Select Case UserList(UserIndex).Raza
    Case "Elfo Oscuro", "Elfo", "Humano"
        If UserList(UserIndex).Clase = "Clerigo" Or UserList(UserIndex).Clase = "Druida" Or UserList(UserIndex).Clase = "Bardo" Then
            MiObj.ObjIndex = 372
        ElseIf UserList(UserIndex).Genero = "Hombre" And UserList(UserIndex).Clase = "Mago" Then
            MiObj.ObjIndex = 517
        ElseIf UserList(UserIndex).Genero = "Mujer" And UserList(UserIndex).Clase = "Mago" Then
            MiObj.ObjIndex = 516
        ElseIf (UserList(UserIndex).Genero = "Mujer") And (UserList(UserIndex).Clase = "Paladin" Or UserList(UserIndex).Clase = "Guerrero" Or UserList(UserIndex).Clase = "Asesino" Or UserList(UserIndex).Clase = "Cazador") Then
            MiObj.ObjIndex = 520
        ElseIf (UserList(UserIndex).Genero = "Hombre") And (UserList(UserIndex).Clase = "Paladin" Or UserList(UserIndex).Clase = "Guerrero" Or UserList(UserIndex).Clase = "Asesino" Or UserList(UserIndex).Clase = "Cazador") Then
            MiObj.ObjIndex = 370
        End If
    
    Case "Gnomo", "Enano"
        If UserList(UserIndex).Clase = "Guerrero" Or UserList(UserIndex).Clase = "Paladin" Or UserList(UserIndex).Clase = "Cazador" Or UserList(UserIndex).Clase = "Asesino" Then
            MiObj.ObjIndex = 492
        ElseIf UserList(UserIndex).Clase = "Mago" Or UserList(UserIndex).Clase = "Bardo" Or UserList(UserIndex).Clase = "Druida" Or UserList(UserIndex).Clase = "Clerigo" Then
            MiObj.ObjIndex = 549
        Else 'Trabajadoras
            MiObj.ObjIndex = 678
        End If
End Select
'[/Wizard]
If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
End If
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call Senddata(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(UserIndex)
End If
Call LogEjercitoReal(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
If UserList(UserIndex).Faccion.CriminalesMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya has recibido tu recompensa, mata 100 crinales mas para recibir la proxima!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
Else
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call Senddata(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
End If
End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
If UserList(UserIndex).GuildInfo.GuildName <> "" And UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
   Dim oGuild As cGuild
   Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
   Call oGuild.RemoveMember(UserList(UserIndex).Name)
   Call AddtoVar(UserList(UserIndex).GuildInfo.Echadas, 1, 1000)
   UserList(UserIndex).GuildInfo.GuildPoints = 0
   UserList(UserIndex).GuildInfo.GuildName = ""
   '[Wizard 03/09/05] Forma burda de actualizar el nick ahorrar lineas, anchodebanda y clonacion de pjs jajajaja pero = es feo.
   Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, False)
End If

UserList(UserIndex).Faccion.ArmadaReal = 0
'Call PerderItemsFaccionarios(UserIndex)
Call Senddata(ToIndex, UserIndex, 0, "Y182")
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)
If UserList(UserIndex).GuildInfo.GuildName <> "" And UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
   Dim oGuild As cGuild
   Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
   Call oGuild.RemoveMember(UserList(UserIndex).Name)
   Call AddtoVar(UserList(UserIndex).GuildInfo.Echadas, 1, 1000)
   UserList(UserIndex).GuildInfo.GuildPoints = 0
   UserList(UserIndex).GuildInfo.GuildName = ""
   '[Wizard 03/09/05] Forma burda de actualizar el nick ahorrar lineas, anchodebanda y clonacion de pjs jajajaja pero = es feo.
   Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, False)
End If

UserList(UserIndex).Faccion.FuerzasCaos = 0
'Call PerderItemsFaccionarios(UserIndex)
Call Senddata(ToIndex, UserIndex, 0, "Y183")
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
        TituloReal = "Campe�n de la Luz"
End Select
End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
If Not Criminal(UserIndex) Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Largate de aqui, bufon!!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya perteneces a las tropas del caos!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
'[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Or UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1 Then 'Tomamos el valor de ah�: �Recibio la experiencia para entrar?
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No permitir� que ning�n insecto real ingrese �Traidor del Rey!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
'[/Barrin]
If Not Criminal(UserIndex) Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ja ja ja tu no eres bienvenido aqui!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.CiudadanosMatados < 150 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes matar al menos 150 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
If UserList(UserIndex).Stats.ELV < 25 Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Exit Sub
End If
UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.CiudadanosMatados \ 100

Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te dare un recompensa, buena suerte soldado!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    

Select Case UserList(UserIndex).Raza
'[Wizard 03/09/05] Arregle la redaccion de este fragmento porque era horrible, 3 case que hacian lo mismo es codigo al pedo.
'Tambien arregla el error de redaccion que producia la mal entrega de las armaduras faccionarias
Case "Humano", "Elfo", "Elfo Oscuro"
    If UserList(UserIndex).Clase = "Clerigo" Or UserList(UserIndex).Clase = "Druida" Or UserList(UserIndex).Clase = "Bardo" Then
        MiObj.ObjIndex = 523
    ElseIf UserList(UserIndex).Clase = "Mago" Then
        MiObj.ObjIndex = 518
    ElseIf (UserList(UserIndex).Genero = "Hombre") And (UserList(UserIndex).Clase = "Paladin" Or UserList(UserIndex).Clase = "Guerrero" Or UserList(UserIndex).Clase = "Asesino" Or UserList(UserIndex).Clase = "Cazador") Then
        MiObj.ObjIndex = 379
    ElseIf (UserList(UserIndex).Genero = "Mujer") And (UserList(UserIndex).Clase = "Paladin" Or UserList(UserIndex).Clase = "Guerrero" Or UserList(UserIndex).Clase = "Asesino" Or UserList(UserIndex).Clase = "Cazador") Then
        MiObj.ObjIndex = 498
    End If
Case "Gnomo", "Enano"
    If UserList(UserIndex).Clase = "Guerrero" Or UserList(UserIndex).Clase = "Paladin" Or UserList(UserIndex).Clase = "Cazador" Or UserList(UserIndex).Clase = "Asesino" Then
        MiObj.ObjIndex = 383
    ElseIf UserList(UserIndex).Clase = "Mago" Or UserList(UserIndex).Clase = "Bardo" Or UserList(UserIndex).Clase = "Druida" Or UserList(UserIndex).Clase = "Clerigo" Then
        MiObj.ObjIndex = 558
    End If
'[/Wizard]
End Select
 


    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
End If
If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call Senddata(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(UserIndex)
End If
Call LogEjercitoCaos(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
If UserList(UserIndex).Faccion.CiudadanosMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya has recibido tu recompensa, mata 100 ciudadanos mas para recibir la proxima!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
Else
    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call Senddata(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
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
        TituloCaos = "Ac�lito"
    Case 3
        TituloCaos = "Guerrero Sombr�o"
    Case 4
        TituloCaos = "Sanguinario"
    Case 5
        TituloCaos = "Caballero de la Oscuridad"
    Case 6
        TituloCaos = "Condenado"
    Case 7
        TituloCaos = "Heraldo Imp�o"
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
'********************Misery_Ezequiel 28/05/05********************'
