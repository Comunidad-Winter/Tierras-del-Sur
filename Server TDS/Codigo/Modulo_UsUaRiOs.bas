Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
Dim DaExp As Integer
'
'Call AddtoVar(UserList(AttackerIndex).Stats.Exp, DaExp, MAXEXP)
'Lo mata
Call Senddata(ToIndex, AttackerIndex, 0, "||Has matado " & UserList(VictimIndex).Name & "!" & FONTTYPE_FIGHT)
'Call Senddata(ToIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & FONTTYPE_FIGHT)
Call Senddata(ToIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).Name & " te ha matado!" & FONTTYPE_FIGHT)
If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
    If (Not Criminal(VictimIndex)) Then
         Call AddtoVar(UserList(AttackerIndex).Reputacion.AsesinoRep, vlASESINO * 2, MAXREP)
         UserList(AttackerIndex).Reputacion.BurguesRep = 0
         UserList(AttackerIndex).Reputacion.NobleRep = 0
         UserList(AttackerIndex).Reputacion.PlebeRep = 0
    Else
         Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
    End If
End If
Call UserDie(VictimIndex)
Call AddtoVar(UserList(AttackerIndex).Stats.UsuariosMatados, 1, 31000)
'Log
Call LogAsesinato(UserList(AttackerIndex).Name & " asesino a " & UserList(VictimIndex).Name)
End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)
UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = 1
UserList(UserIndex).Stats.MinSta = 0
If UserList(UserIndex).flags.Navegando Then
UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
UserList(UserIndex).Char.Head = 0
Else
Call DarCuerpoDesnudo(UserIndex)
End If

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendUserStatsBox(UserIndex)
Call EnviarHambreYsed(UserIndex)
End Sub

Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
UserList(UserIndex).Char.Body = Body
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.Heading = Heading
UserList(UserIndex).Char.WeaponAnim = Arma
UserList(UserIndex).Char.ShieldAnim = Escudo
UserList(UserIndex).Char.CascoAnim = Casco
Call Senddata(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.charindex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco)
End Sub

Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
Call Senddata(ToIndex, UserIndex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad$
For i = 1 To NUMSKILLS
   cad$ = cad$ & UserList(UserIndex).Stats.UserSkills(i) & ","
Next
Senddata ToIndex, UserIndex, 0, "SKILLS" & cad$
End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
Dim cad$
cad$ = cad$ & UserList(UserIndex).Reputacion.AsesinoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BandidoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BurguesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.LadronesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.NobleRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.PlebeRep & ","
Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
UserList(UserIndex).Reputacion.Promedio = L
cad$ = cad$ & UserList(UserIndex).Reputacion.Promedio
Senddata ToIndex, UserIndex, 0, "FAMA" & cad$
End Sub

Sub EnviarAtrib(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad$
For i = 1 To NUMATRIBUTOS
  cad$ = cad$ & UserList(UserIndex).Stats.UserAtributos(i) & ","
Next
Call Senddata(ToIndex, UserIndex, 0, "ATR" & cad$)
End Sub

Public Sub EnviarMiniEstadisticas(ByVal UserIndex As Integer)
With UserList(UserIndex)
    Call Senddata(ToIndex, UserIndex, 0, "MEST" & .Faccion.CiudadanosMatados & "," & _
                .Faccion.CriminalesMatados & "," & .Stats.UsuariosMatados & "," & _
                .Stats.NPCsMuertos & "," & .Clase & "," & .Counters.Pena)
End With
End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)
On Error GoTo ErrorHandler
Dim Linea As Integer
Linea = 0
    CharList(UserList(UserIndex).Char.charindex) = 0
    If UserList(UserIndex).Char.charindex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    Linea = 1
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    'Le mandamos el mensaje para que borre el personaje a los clientes que este en el mismo mapa
    Call Senddata(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "BP" & UserList(UserIndex).Char.charindex)
    Linea = 2
    UserList(UserIndex).Char.charindex = 0
    NumChars = NumChars - 1
    Linea = 3
    Exit Sub
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description & "Usuario: " & UserList(UserIndex).Name & " Probocado en: " & Linea)
End Sub

Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Local Error GoTo hayerror
Dim charindex As Integer

If InMapBounds(Map, X, Y) Then
       'If needed make a new character in list
       If UserList(UserIndex).Char.charindex = 0 Then
           charindex = NextOpenCharIndex
           UserList(UserIndex).Char.charindex = charindex
           CharList(charindex) = UserIndex
       End If
       'Place character on map
       MapData(Map, X, Y).UserIndex = UserIndex
       'Send make character command to clients
       Dim klan$
       klan$ = UserList(UserIndex).GuildInfo.GuildName
       Dim bCr As Byte
       bCr = Criminal(UserIndex)
       If klan$ <> "" Then
'            If EncriptarProtocolosCriticos And sndRoute = ToMap Then
 '               If UserList(UserIndex).flags.Privilegios > 0 Then
  '                  Call SendCryptedData(ToMapButIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
   '                 Call Senddata(ToIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)    'porque no le di todavia el charindeX!!!
    '            Else
     '               Call SendCryptedData(ToMapButIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 0, 0))
      '              Call Senddata(ToIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 0, 0))    'porque no le di todavia el charindeX!!!
       '         End If
        '    Else
                If UserList(UserIndex).flags.Privilegios > 0 Then
                    Call Senddata(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                Else
                    Call Senddata(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1 Or UserList(UserIndex).flags.PertAlConsCaos = 1, 4, 0))
                End If
         '   End If
       Else
           ' If EncriptarProtocolosCriticos And sndRoute = ToMap Then
            '    If UserList(UserIndex).flags.Privilegios > 0 Then
             '       Call SendCryptedData(ToMapButIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
              ''      Call Senddata(ToIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                'Else
                 '   Call SendCryptedData(ToMapButIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 0, 0))
                  '  Call Senddata(ToIndex, UserIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 0, 0))
                'End If
            'Else
                If UserList(UserIndex).flags.Privilegios > 0 Then
                    Call Senddata(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                Else
                    Call Senddata(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.charindex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1 Or UserList(UserIndex).flags.PertAlConsCaos = 1, 4, 0))
                End If
            'End If
       End If
End If
Exit Sub
hayerror:
LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
Call CloseSocket(UserIndex)
End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoST As Integer
Dim AumentoMANA As Integer
Dim WasNewbie As Boolean

'¿Alcanzo el maximo nivel?
If UserList(UserIndex).Stats.ELV = STAT_MAXELV Then
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 0
    Exit Sub
End If
WasNewbie = EsNewbie(UserIndex)
'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
Do While UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_NIVEL)
    Call Senddata(ToIndex, UserIndex, 0, "Y47")
    If UserList(UserIndex).Stats.ELV = 1 Then
      Pts = 10
    Else
      Pts = 5
    End If
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
    Call Senddata(ToIndex, UserIndex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
    UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    
'[Misery_Ezequiel 30/06/05]
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp - UserList(UserIndex).Stats.ELU
    If Not EsNewbie(UserIndex) And WasNewbie Then Call QuitarNewbieObj(UserIndex)
    If UserList(UserIndex).Stats.ELV < 11 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    ElseIf UserList(UserIndex).Stats.ELV <= 24 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    ElseIf UserList(UserIndex).Stats.ELV = 25 Then
        UserList(UserIndex).Stats.ELU = 544727
    ElseIf UserList(UserIndex).Stats.ELV = 26 Then
        UserList(UserIndex).Stats.ELU = 663672
    ElseIf UserList(UserIndex).Stats.ELV = 27 Then
        UserList(UserIndex).Stats.ELU = 784406
    ElseIf UserList(UserIndex).Stats.ELV = 28 Then
        UserList(UserIndex).Stats.ELU = 941287
    ElseIf UserList(UserIndex).Stats.ELV = 29 Then
        UserList(UserIndex).Stats.ELU = 1129544
    ElseIf UserList(UserIndex).Stats.ELV = 30 Then
        UserList(UserIndex).Stats.ELU = 1355453
    ElseIf UserList(UserIndex).Stats.ELV = 31 Then
        UserList(UserIndex).Stats.ELU = 1626544
    ElseIf UserList(UserIndex).Stats.ELV = 32 Then
        UserList(UserIndex).Stats.ELU = 1951853
    ElseIf UserList(UserIndex).Stats.ELV = 33 Then
        UserList(UserIndex).Stats.ELU = 2342224
    ElseIf UserList(UserIndex).Stats.ELV = 34 Then
        UserList(UserIndex).Stats.ELU = 3372803
    ElseIf UserList(UserIndex).Stats.ELV = 35 Then
        UserList(UserIndex).Stats.ELU = 4047364
    ElseIf UserList(UserIndex).Stats.ELV = 36 Then
        UserList(UserIndex).Stats.ELU = 5828204
    ElseIf UserList(UserIndex).Stats.ELV = 37 Then
        UserList(UserIndex).Stats.ELU = 6993845
    ElseIf UserList(UserIndex).Stats.ELV = 38 Then
        UserList(UserIndex).Stats.ELU = 8392614
    ElseIf UserList(UserIndex).Stats.ELV = 39 Then
        UserList(UserIndex).Stats.ELU = 10071137
    ElseIf UserList(UserIndex).Stats.ELV = 40 Then
        UserList(UserIndex).Stats.ELU = 120853640
    ElseIf UserList(UserIndex).Stats.ELV = 41 Then
        UserList(UserIndex).Stats.ELU = 145024370
    ElseIf UserList(UserIndex).Stats.ELV = 42 Then
        UserList(UserIndex).Stats.ELU = 174029240
    ElseIf UserList(UserIndex).Stats.ELV = 43 Then
        UserList(UserIndex).Stats.ELU = 208835090
    ElseIf UserList(UserIndex).Stats.ELV = 44 Then
        UserList(UserIndex).Stats.ELU = 417670180
    ElseIf UserList(UserIndex).Stats.ELV = 45 Then
        UserList(UserIndex).Stats.ELU = 835340360
    ElseIf UserList(UserIndex).Stats.ELV = 46 Then
        UserList(UserIndex).Stats.ELU = 1670680720
    ElseIf UserList(UserIndex).Stats.ELV = 47 Then
        UserList(UserIndex).Stats.ELU = 3341361440#
    Else
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
    End If
'[\]Misery_Ezequiel 30/06/05]
    Dim AumentoHP As Integer
    Select Case UserList(UserIndex).Clase
'[MISERY_EZEQUIEL 26/06/05]********************************************
        Case "Guerrero"
            'marche
        Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 12)
                Case 19
                    AumentoHP = RandomNumber(8, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPCazador
            End Select
           'marche
            AumentoST = 15
            '[Misery_Ezequiel 17/06/05]
            If UserList(UserIndex).Stats.MaxHIT < 99 Then
                AumentoHIT = 3
            Else
                If UserList(UserIndex).Stats.MaxHIT >= 99 Then
                AumentoHIT = 2
                End If
            End If
            '[\]Misery_Ezequiel 17/06/05]
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
        Case "Cazador"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPCazador
            End Select
            AumentoST = 15
            '[Misery_Ezequiel 17/06/05]
            If UserList(UserIndex).Stats.MaxHIT < 99 Then
                AumentoHIT = 3
            Else
                If UserList(UserIndex).Stats.MaxHIT >= 99 Then
                AumentoHIT = 1
                End If
            End If
            '[\]Misery_Ezequiel 17/06/05]
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
        Case "Pirata"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPCazador
            End Select
            AumentoST = 15
            AumentoHIT = 3
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
        Case "Paladin"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPCazador
            End Select
            AumentoST = 15
            '[Misery_Ezequiel 17/06/05]
            If UserList(UserIndex).Stats.MaxHIT < 99 Then
                AumentoHIT = 3
            Else
                If UserList(UserIndex).Stats.MaxHIT >= 99 Then
                AumentoHIT = 1
                End If
            End If
            '[\]Misery_Ezequiel 17/06/05]
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            'Mana
            Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN)
            'STA
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            'Golpe
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
        Case "Ladron"
            '[eLwE 19/05/05]
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 9)
                Case 17
                    AumentoHP = RandomNumber(4, 8)
                Case 16
                    AumentoHP = RandomNumber(3, 7)
                Case 16
                    AumentoHP = RandomNumber(3, 6)
                Case 14
                    AumentoHP = RandomNumber(2, 6)
                Case 13
                    AumentoHP = RandomNumber(2, 5)
                Case 12
                    AumentoHP = RandomNumber(1, 5)
                Case 11
                    AumentoHP = RandomNumber(1, 4)
                Case 10
                    AumentoHP = RandomNumber(0, 4)
                Case Else
                    AumentoHP = RandomNumber(3, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero
            End Select
            '[\]eLwE 19/05/05]
            'HP
            AumentoST = 15 + AdicionalSTLadron
            AumentoHIT = 1
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Mago"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 8)
                Case 20
                    AumentoHP = RandomNumber(5, 8)
                Case 19
                    AumentoHP = RandomNumber(4, 8)
                Case 18
                    AumentoHP = RandomNumber(3, 8)
                Case Else
                    AumentoHP = RandomNumber(3, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            AumentoST = 15 - AdicionalSTLadron / 2
            If AumentoST < 1 Then AumentoST = 5
            AumentoHIT = 1
            If UserList(UserIndex).Stats.MaxMAN < 2000 Then
            AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                Else
                If UserList(UserIndex).Stats.MaxMAN >= 2000 Then
                AumentoMANA = (3 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 2
                End If
            End If
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Leñador"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 8)
                Case 20
                    AumentoHP = RandomNumber(5, 8)
                Case 19
                    AumentoHP = RandomNumber(4, 8)
                Case 18
                    AumentoHP = RandomNumber(3, 8)
                Case Else
                    AumentoHP = RandomNumber(2, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero
            End Select
            AumentoST = 14
            AumentoHIT = 2
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Minero"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 8)
                Case 20
                    AumentoHP = RandomNumber(5, 8)
                Case 19
                    AumentoHP = RandomNumber(4, 8)
                Case 18
                    AumentoHP = RandomNumber(3, 8)
                Case Else
                    AumentoHP = RandomNumber(3, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero
            End Select
            AumentoST = 14
            AumentoHIT = 2
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Pescador"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 8)
                Case 20
                    AumentoHP = RandomNumber(5, 8)
                Case 19
                    AumentoHP = RandomNumber(4, 8)
                Case 18
                    AumentoHP = RandomNumber(3, 8)
                Case Else
                    AumentoHP = RandomNumber(3, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero
            End Select
            AumentoST = 14
            AumentoHIT = 1
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Clerigo"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPCazador
            End Select
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Druida"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPCazador
            End Select
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Asesino"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPCazador
            End Select
            AumentoST = 15
            '[Misery_Ezequiel 17/06/05]
            If UserList(UserIndex).Stats.MaxHIT < 99 Then
                AumentoHIT = 3
            Else
                If UserList(UserIndex).Stats.MaxHIT >= 99 Then
                AumentoHIT = 1
                End If
            End If
            '[\]Misery_Ezequiel 17/06/05]
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Bardo"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPCazador
            End Select
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Herrero"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPCazador
            End Select
            AumentoST = 14
            AumentoHIT = 1
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case Else
             Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPCazador
            End Select
            AumentoST = 15
            AumentoHIT = 2
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
    End Select
'[\]MISERY_EZEQUIEL 17/06/05]*******************************************
    'AddtoVar UserList(UserIndex).Stats.MaxHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.MinHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.Def, 2, STAT_MAXDEF
    If AumentoHP > 0 Then Senddata ToIndex, UserIndex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO
    If AumentoST > 0 Then Senddata ToIndex, UserIndex, 0, "||Has ganado " & AumentoST & " puntos de vitalidad." & FONTTYPE_INFO
    If AumentoMANA > 0 Then Senddata ToIndex, UserIndex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO
    If AumentoHIT > 0 Then
        Senddata ToIndex, UserIndex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        Senddata ToIndex, UserIndex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    Call LogDesarrollo(Date & " " & UserList(UserIndex).Name & " paso a nivel " & UserList(UserIndex).Stats.ELV & " gano HP: " & AumentoHP)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, Pts)
    SendUserStatsBox UserIndex
Loop
Exit Sub
errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
    PuedeAtravesarAgua = _
        UserList(UserIndex).flags.Navegando = 1 Or _
        UserList(UserIndex).flags.Vuela = 1
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)
On Error Resume Next
Dim nPos As WorldPos
Dim vpos As WorldPos
Dim Apos As WorldPos
vpos = UserList(UserIndex).Pos
nPos = UserList(UserIndex).Pos
Call HeadtoPos(nHeading, nPos)
If LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(UserIndex)) Then
    Call Senddata(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, "M" & nHeading & UserList(UserIndex).Char.charindex)
    'Update map and user pos
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).Pos = nPos
    UserList(UserIndex).Char.Heading = nHeading
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
Else
    Call Senddata(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
End If

'[Barrin 30-11-03]
UserList(UserIndex).flags.Trabajando = False
'[/Barrin 30-11-03]
End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)
UserList(UserIndex).Invent.Object(Slot) = Object

If Object.ObjIndex > 0 Then
    Call Senddata(ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).ObjType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef & "," _
    & ObjData(Object.ObjIndex).Valor \ 3)
Else
    Call Senddata(ToIndex, UserIndex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")
End If
End Sub

Function NextOpenCharIndex() As Integer
On Local Error GoTo hayerror
Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next LoopC
Exit Function
hayerror:
LogError ("NextOpenCharIndex: num: " & Err.Number & " desc: " & Err.Description)
End Function

Function NextOpenUser() As Integer
Dim LoopC As Integer

For LoopC = 1 To MaxUsers + 1
  If LoopC > MaxUsers Then Exit For
  If (UserList(LoopC).ConnID = -1) Then Exit For
Next LoopC
NextOpenUser = LoopC
End Function

Sub SendUserStatsBox(ByVal UserIndex As Integer)
Call Senddata(ToIndex, UserIndex, 0, "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.Exp)
End Sub

Sub EnviarHambreYsed(ByVal UserIndex As Integer)
Call Senddata(ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MaxAGU & "," & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MaxHam & "," & UserList(UserIndex).Stats.MinHam)
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
Call Senddata(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & FONTTYPE_INFO)
Call Senddata(ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
Call Senddata(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(UserIndex).Stats.FIT & FONTTYPE_INFO)
Call Senddata(ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Call Senddata(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
Else
    Call Senddata(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & FONTTYPE_INFO)
End If
If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    Call Senddata(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call Senddata(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If
If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    Call Senddata(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call Senddata(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If
If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    Call Senddata(ToIndex, sendIndex, 0, "||Clan: " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_INFO)
    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
       If UserList(UserIndex).GuildInfo.ClanFundado = UserList(UserIndex).GuildInfo.GuildName Then
            Call Senddata(ToIndex, sendIndex, 0, "||Status:" & "Fundador/Lider" & FONTTYPE_INFO)
       Else
            Call Senddata(ToIndex, sendIndex, 0, "||Status:" & "Lider" & FONTTYPE_INFO)
       End If
    Else
        Call Senddata(ToIndex, sendIndex, 0, "||Status:" & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call Senddata(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
End If
Call Senddata(ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)
Call Senddata(ToIndex, sendIndex, 0, "||Dados: " & UserList(UserIndex).Stats.UserAtributos(1) & ", " & UserList(UserIndex).Stats.UserAtributos(2) & ", " & UserList(UserIndex).Stats.UserAtributos(3) & ", " & UserList(UserIndex).Stats.UserAtributos(4) & ", " & UserList(UserIndex).Stats.UserAtributos(5) & FONTTYPE_INFO)
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
With UserList(UserIndex)
    Call Senddata(ToIndex, sendIndex, 0, "||Pj: " & .Name & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
'    Call SendData(ToIndex, sendIndex, 0, "||CriminalesMatados: " & .Faccion.CriminalesMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
'    Call SendData(ToIndex, sendIndex, 0, "||UsuariosMatados: " & .Stats.UsuariosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Clase: " & .Clase & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"

If CharExist(CharName) Then
    Call Senddata(ToIndex, sendIndex, 0, "||Pj: " & CharName & FONTTYPE_INFO)
    ' 3 en uno :p
    Call Senddata(ToIndex, sendIndex, 0, "||CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes") & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
    Ban = GetVar(CharFile, "FLAGS", "Ban")
    Call Senddata(ToIndex, sendIndex, 0, "||Ban: " & Ban & FONTTYPE_INFO)
    If Ban = "1" Then
        Call Senddata(ToIndex, sendIndex, 0, "||Ban por: " & GetVar(BanDetailPath, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)
    End If
Else
    Call Senddata(ToIndex, sendIndex, 0, "||El pj no existe: " & CharName & FONTTYPE_INFO)
End If
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call Senddata(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
Call Senddata(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
For j = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
        Call Senddata(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"
If FileExist(CharFile, vbNormal) Then
    Call Senddata(ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
    For j = 1 To MAX_INVENTORY_SLOTS
        Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
        ObjInd = ReadField(1, Tmp, Asc("-"))
        ObjCant = ReadField(2, Tmp, Asc("-"))
        If ObjInd > 0 Then
            Call Senddata(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
        End If
    Next
Else
    Call Senddata(ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call Senddata(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call Senddata(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
End Sub

Sub UpdateUserMap(ByVal UserIndex As Integer)
Dim Map As Integer
Dim X As Integer
Dim Y As Integer

On Error GoTo 0
Map = UserList(UserIndex).Pos.Map
Dim cadena As String
Dim Cantidad As Integer
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
            Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
            If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Then Call Senddata(ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.charindex & ",1")
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            Call MakeNPCChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                Dim Obj As Obj
                If Obj.ObjIndex <= UBound(ObjData) Then
                  cadena = cadena & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y & ";"
                Cantidad = Cantidad + 1
                End If
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
                      Call Bloquear(ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                      Call Bloquear(ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
            End If
        End If
        
    Next X
Next Y
Call Senddata(ToIndex, UserIndex, 0, "CI" & Cantidad & ";" & cadena)
End Sub
Function DameUserindex(SocketId As Integer) As Integer
Dim LoopC As Integer
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId
    LoopC = LoopC + 1
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
Loop
DameUserindex = LoopC
End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer
Dim LoopC As Integer
LoopC = 1
  
Nombre = UCase$(Nombre)
Do Until UCase$(UserList(LoopC).Name) = Nombre
    LoopC = LoopC + 1
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
Loop
DameUserIndexConNombre = LoopC
End Function

Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call Senddata(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!" & FONTTYPE_FIGHT)
End If
End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            Npclist(NpcIndex).Movement = NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                Call VolverCriminal(UserIndex)
       Else
            If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
            End If
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR / 2, MAXREP)
    End If
    
    'hacemos que el npc se defienda
    Npclist(NpcIndex).Movement = NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
End If
'Marche
' Call CheckPets(NpcIndex, UserIndex)
End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(UserIndex).Clase = "Asesino") And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
If UserList(UserIndex).flags.Hambre = 0 And _
   UserList(UserIndex).flags.Sed = 0 Then
    Dim Aumenta As Integer
    Dim Prob As Integer
    
    
    If Skill = 1 Then
    If UserList(UserIndex).Stats.MaxMAN = 0 And UserList(UserIndex).Stats.UserSkills(1) < 5 Then Exit Sub
    End If
    
    If UserList(UserIndex).Stats.ELV <= 3 Then
        Prob = 20
    ElseIf UserList(UserIndex).Stats.ELV > 3 _
        And UserList(UserIndex).Stats.ELV < 6 Then
        Prob = 30
    ElseIf UserList(UserIndex).Stats.ELV >= 6 _
        And UserList(UserIndex).Stats.ELV < 10 Then
        Prob = 35
    ElseIf UserList(UserIndex).Stats.ELV >= 10 _
        And UserList(UserIndex).Stats.ELV < 20 Then
        Prob = 40
    Else
        Prob = 45
    End If
    Aumenta = Int(RandomNumber(1, Prob))
    
    Dim lvl As Integer
    lvl = UserList(UserIndex).Stats.ELV
    
    If lvl >= UBound(LevelSkill) Then Exit Sub
    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    If Aumenta = 7 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
            Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
            Call Senddata(ToIndex, UserIndex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
            Call AddtoVar(UserList(UserIndex).Stats.Exp, 50, MAXEXP)
            Call Senddata(ToIndex, UserIndex, 0, "Y48")
            Call CheckUserLevel(UserIndex)
            Call SendUserStatsBox(UserIndex)
    End If
End If
End Sub

Sub UserDie(ByVal UserIndex As Integer)
'Call LogTarea("Sub UserDie")
On Error GoTo ErrorHandler
'Sonido
Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_USERMUERTE)
'Quitar el dialogo del user muerto
Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.charindex)
UserList(UserIndex).Stats.MinHP = 0
UserList(UserIndex).Stats.MinSta = 0
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.AtacadoPorUser = 0
UserList(UserIndex).flags.Envenenado = 0

'[Wizard 03/09/05]
UserList(UserIndex).Counters.Veneno = 0

'[Wizard]
UserList(UserIndex).flags.Muerto = 1
Dim aN As Integer
aN = UserList(UserIndex).flags.AtacadoPorNpc

If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If
'<<<< Paralisis >>>>
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).flags.Paralizado = 0
    Call Senddata(ToIndex, UserIndex, 0, "PARADOK")
End If
'<<<< Descansando >>>>
If UserList(UserIndex).flags.Descansar Then
    UserList(UserIndex).flags.Descansar = False
    Call Senddata(ToIndex, UserIndex, 0, "DOK")
End If
'<<<< Meditando >>>>
If UserList(UserIndex).flags.Meditando Then
    UserList(UserIndex).flags.Meditando = False
    Call Senddata(ToIndex, UserIndex, 0, "MEDOK")
End If
'<<<< Invisible >>>>
If UserList(UserIndex).flags.Invisible = 1 Then
    UserList(UserIndex).flags.Oculto = 0
    UserList(UserIndex).flags.Invisible = 0
    'no hace falta encriptar este NOVER
    Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
End If
If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
'[Misery_Ezequiel 10/06/05]
If EsNewbie(UserIndex) And MapInfo(UserList(UserIndex).Pos.Map).Restringir = "Si" Then
    'NO PIÑATEO NADA
ElseIf EsNewbie(UserIndex) And MapInfo(UserList(UserIndex).Pos.Map).Restringir <> "Si" Then
     Call TirarTodosLosItemsNoNewbies(UserIndex)
ElseIf Not EsNewbie(UserIndex) Then
    Call TirarTodo(UserIndex)
End If
End If
'[\]Misery_Ezequiel 10/06/05]
' DESEQUIPA TODOS LOS OBJETOS
'desequipar armadura
If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End If
'desequipar arma
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
End If
'[Misery_Ezequiel 06/06/05]
'desequipar escudo
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
End If
'[\]Misery_Ezequiel 06/06/05]
'desequipar casco
If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
End If
'[Misery_Ezequiel 06/06/05]
'desequipar herramienta
If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
End If
'[\]Misery_Ezequiel 06/06/05]
'desequipar municiones
If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
End If
' << Reseteamos los posibles FX sobre el personaje >>
If UserList(UserIndex).Char.loops = LoopAdEternum Then
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
End If

'[Wizard 07/09/05] Actualiza el inventario.
    Call UpdateUserInv(True, UserIndex, 0)
'/Wizard


' << Restauramos el mimetismo
If UserList(UserIndex).flags.Mimetizado = 1 Then
    UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
    UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
End If

'<< Cambiamos la apariencia del char >>
If UserList(UserIndex).flags.Navegando = 0 Then
'[Misery_Ezequiel 12/06/05]
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    UserList(UserIndex).Char.Body = iCuerpoMuertoCrimi
    UserList(UserIndex).Char.Head = iCabezaMuertoCrimi
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
Else
'[\]Misery_Ezequiel 12/06/05]
    UserList(UserIndex).Char.Body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
End If
ElseIf UserList(UserIndex).flags.Navegando = 1 Then
    UserList(UserIndex).Char.Body = iFragataFantasmal ';)
End If

If UserList(UserIndex).flags.Estupidez = 1 Then
Call Senddata(ToIndex, UserIndex, 0, "NESTUP")
UserList(UserIndex).flags.Estupidez = 0
End If

Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
            If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
           Else
               Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
               Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
                UserList(UserIndex).MascotasIndex(i) = 0
               UserList(UserIndex).MascotasType(i) = 0
           End If
    End If
Next i

UserList(UserIndex).NroMacotas = 0
If UserList(UserIndex).flags.Estupidez = True Then
UserList(UserIndex).flags.Estupidez = 0
Call Senddata(ToIndex, UserIndex, 0, "NESTUP")
Call InfoHechizo(UserIndex)
End If
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
Call SendUserStatsBox(UserIndex)
Call Senddata(ToIndex, UserIndex, 0, "CSI")



UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
Call UpdateUserInv(True, UserIndex, 0)
Exit Sub
ErrorHandler:
    Call LogError("Error en SUB USERDIE")
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
If EsNewbie(Muerto) Then Exit Sub
If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(Atacante).Faccion.CriminalesMatados, 1, 65000)
        End If
        If UserList(Atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CriminalesMatados = 0
            UserList(Atacante).Faccion.RecompensasReal = 0
        End If
Else
        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(Atacante).Faccion.CiudadanosMatados, 1, 65000)
        End If
        If UserList(Atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CiudadanosMatados = 0
            UserList(Atacante).Faccion.RecompensasCaos = 0
        End If
End If
End Sub

Sub Tilelibre(Pos As WorldPos, nPos As WorldPos)
'Call LogTarea("Sub Tilelibre")
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
hayobj = False
nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
    
    If LoopC > 15 Then
        Notfound = True
        Exit Do
    End If
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            If LegalPos(nPos.Map, tX, tY) = True Then
               hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0)
               If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                     nPos.X = tX
                     nPos.Y = tY
                     tX = Pos.X + LoopC
                     tY = Pos.Y + LoopC
                End If
            End If
        Next tX
    Next tY
    LoopC = LoopC + 1
Loop
If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
'Quitar el dialogo
Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.charindex)
Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "QTDL")
Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

OldMap = UserList(UserIndex).Pos.Map
OldX = UserList(UserIndex).Pos.X
OldY = UserList(UserIndex).Pos.Y
Errorestaen = "1"
Call EraseUserChar(ToMap, 0, OldMap, UserIndex)
UserList(UserIndex).Pos.X = X
UserList(UserIndex).Pos.Y = Y
UserList(UserIndex).Pos.Map = Map

If OldMap <> Map Then
    Call Senddata(ToIndex, UserIndex, 0, "CM" & Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion & "," & MapInfo(UserList(UserIndex).Pos.Map).Terreno & "," & MapInfo(UserList(UserIndex).Pos.Map).Zona)
    Call Senddata(ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)
'    Call EnviarNoche(UserIndex)
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Call Senddata(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.charindex)
    'Update new Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    'Update old Map Users
    MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
    If MapInfo(OldMap).NumUsers < 0 Then
        MapInfo(OldMap).NumUsers = 0
    End If
    Call UpdateUserMap(UserIndex)
Else
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Call Senddata(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.charindex)
End If


        'Seguis invisible al pasar de mapa
        If (UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
           ' If EncriptarProtocolosCriticos Then
            '    Call SendCryptedData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",1")
           ' Else
                Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",1")
            'End If
        End If


If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
If MapInfo(Map).Name = "Dungeon Magma" Then
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & 19 & "," & 0)
Else
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXWARP & "," & 0)
End If
End If
Call WarpMascotas(UserIndex)
End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer
Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer
Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(UserIndex).NroMacotas
InvocadosMatados = 0

'Matamos los invocados
'[Alejo 18-03-2004]
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            UserList(UserIndex).MascotasIndex(i) = 0
            InvocadosMatados = InvocadosMatados + 1
            NroPets = NroPets - 1
        End If
    End If
Next i
If InvocadosMatados > 0 Then
    Call Senddata(ToIndex, UserIndex, 0, "Y49")
End If
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
        PetTypes(i) = UserList(UserIndex).MascotasType(i)
        PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
        Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i
For i = 1 To MAXMASCOTAS
    If PetTypes(i) > 0 Then
        UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
        UserList(UserIndex).MascotasType(i) = PetTypes(i)
        'Controlamos que se sumoneo OK
        If UserList(UserIndex).MascotasIndex(i) = MAXNPCS Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                Exit Sub
        End If
        Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
        Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = SIGUE_AMO
        Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
        Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
    End If
Next i
UserList(UserIndex).NroMacotas = NroPets
End Sub

Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
Next i
If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0
End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal Tiempo As Integer = -1)
 
    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True
        'UserList(UserIndex).Counters.Salir = IIf(UserList(UserIndex).flags.Privilegios > 0 Or MapInfo(UserList(UserIndex).Pos.Map).Pk = False, 0, Tiempo)
        '[Misery_Ezequiel 11/06/05]
        If UserList(UserIndex).flags.Privilegios > 0 Or MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
           UserList(UserIndex).Counters.Salir = 0
        ElseIf UserList(UserIndex).flags.Privilegios = 0 Or MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
           UserList(UserIndex).Counters.Salir = Tiempo
        Call Senddata(ToIndex, UserIndex, 0, "||Cerrando...Se cerrará el juego en " & UserList(UserIndex).Counters.Salir & " segundos..." & FONTTYPE_INFO)
        'Call CloseUser(UserIndex)
        End If
        '[\]Misery_Ezequiel 11/06/05]
    'ElseIf Not UserList(UserIndex).Counters.Saliendo Then
    '    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    '    Call SendData(ToIndex, UserIndex, 0, "||Gracias por jugar Argentum Online" & FONTTYPE_INFO)
    '    Call SendData(ToIndex, UserIndex, 0, "FINOK")
    '
    '    Call CloseUser(UserIndex)
    '    UserList(UserIndex).ConnID = -1: UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    '    frmMain.Socket2(UserIndex).Cleanup
    '    Unload frmMain.Socket2(UserIndex)
    '    Call ResetUserSlot(UserIndex)
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
Dim ViejoNick As String
Dim ViejoCharBackup As String

If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
ViejoNick = UserList(UserIndexDestino).Name
If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
    'hace un backup del char
    ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
    Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
End If
End Sub

Public Sub Empollando(ByVal UserIndex As Integer)
'On Error Resume Next
'Dim nPos As WorldPos
'nPos = UserList(UserIndex).Pos
'If MapData(nPos.Map, nPos.X, nPos.Y).OBJInfo.Amount > 0 And _
'(LegalPos(UserList(UserIndex).Pos.Map, nPos.X + 1, nPos.Y, PuedeAtravesarAgua(UserIndex)) Or _
'LegalPos(UserList(UserIndex).Pos.Map, nPos.X - 1, nPos.Y, PuedeAtravesarAgua(UserIndex)) Or _
'LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y + 1, PuedeAtravesarAgua(UserIndex)) Or _
'LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y - 1, PuedeAtravesarAgua(UserIndex))) Then
'    UserList(UserIndex).flags.EstaEmpo = 1
'Else
'    UserList(UserIndex).flags.EstaEmpo = 0
'End If
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    UserList(UserIndex).flags.EstaEmpo = 1
Else
    UserList(UserIndex).flags.EstaEmpo = 0
    UserList(UserIndex).EmpoCont = 0
End If
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)
Dim FileNamE As String
FileNamE = Nombre
If CharExist(FileNamE) = False Then
    Call Senddata(ToIndex, sendIndex, 0, "Y50")
Else
    Call Senddata(ToIndex, sendIndex, 0, "||Estadisticas de: " & Nombre & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Nivel: " & rs!elvb & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Vitalidad: " & rs!MaxStaB & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Salud: " & rs!MinHPB & "/" & rs!MaxHPB & "  Mana: " & rs!MinMANB & "/" & rs!MaxMANb & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & rs!MinHITB & "/" & rs!MaxHITB & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "||Oro: " & rs!gldb & FONTTYPE_INFO)
End If
Exit Sub
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"
If FileExist(CharFile, vbNormal) Then
    Call Senddata(ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call Senddata(ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
    Call Senddata(ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If
End Sub

Sub RevivirUsuarioEnREeto(ByVal UserIndex As Integer)
UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendUserStatsBox(UserIndex)
Call EnviarHambreYsed(UserIndex)
End Sub

'[Misery_Ezequiel 05/06/05]
Sub UserDiePocionNegra(ByVal UserIndex As Integer)
'Call LogTarea("Sub UserDie")
On Error GoTo ErrorHandler
'Sonido
Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_USERMUERTE)
'Quitar el dialogo del user muerto
Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.charindex)
UserList(UserIndex).Stats.MinHP = 0
UserList(UserIndex).Stats.MinSta = 0
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.AtacadoPorUser = 0
UserList(UserIndex).flags.Envenenado = 0
UserList(UserIndex).flags.Muerto = 1
Dim aN As Integer
aN = UserList(UserIndex).flags.AtacadoPorNpc

If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If
'<<<< Paralisis >>>>
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).flags.Paralizado = 0
    Call Senddata(ToIndex, UserIndex, 0, "PARADOK")
End If
'<<<< Descansando >>>>
If UserList(UserIndex).flags.Descansar Then
    UserList(UserIndex).flags.Descansar = False
    Call Senddata(ToIndex, UserIndex, 0, "DOK")
End If
'<<<< Meditando >>>>
If UserList(UserIndex).flags.Meditando Then
    UserList(UserIndex).flags.Meditando = False
    Call Senddata(ToIndex, UserIndex, 0, "MEDOK")
End If
'<<<< Invisible >>>>
If UserList(UserIndex).flags.Invisible = 1 Then
    UserList(UserIndex).flags.Oculto = 0
    UserList(UserIndex).flags.Invisible = 0
    'no hace falta encriptar este NOVER
    Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
End If
If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
    ' << Si es newbie no pierde el inventario >>
    If EsNewbie(UserIndex) Then
    'NO PIÑATEO NADA
    ElseIf EsNewbie(UserIndex) And Criminal(UserIndex) Then
    'NO PIÑATEO NADA
    ElseIf Not EsNewbie(UserIndex) Then
      Call TirarTodo(UserIndex)
    End If
End If
' DESEQUIPA TODOS LOS OBJETOS
'desequipar armadura
If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End If
'desequipar arma
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
End If
'desequipar escudo
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
End If
'desequipar casco
If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
End If
'desequipar herramienta
If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
End If
'desequipar municiones
If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
End If
' << Reseteamos los posibles FX sobre el personaje >>
If UserList(UserIndex).Char.loops = LoopAdEternum Then
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
End If
'<< Cambiamos la apariencia del char >>
If UserList(UserIndex).flags.Navegando = 0 Then
'[Misery_Ezequiel 12/06/05]
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    UserList(UserIndex).Char.Body = iCuerpoMuertoCrimi
    UserList(UserIndex).Char.Head = iCabezaMuertoCrimi
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
Else
'[\]Misery_Ezequiel 12/06/05]
    UserList(UserIndex).Char.Body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
End If
ElseIf UserList(UserIndex).flags.Navegando = 1 Then
    UserList(UserIndex).Char.Body = iFragataFantasmal ';)
End If

Dim i As Integer
For i = 1 To MAXMASCOTAS
    'If UserList(UserIndex).MascotasIndex(i) > 0 Then
    '       If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
    '       Else
    '            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
    '            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
    '            Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
    '            UserList(UserIndex).MascotasIndex(i) = 0
    '            UserList(UserIndex).MascotasType(i) = 0
    '       End If
    'End If
Next i

UserList(UserIndex).NroMacotas = 0

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
Call SendUserStatsBox(UserIndex)
Call Senddata(ToIndex, UserIndex, 0, "CSI")

 If UserList(UserIndex).flags.Retando = True Then
 UserList(UserIndex).flags.Retando = False
 UserList(UserIndex).flags.Perdio = 1
 Call HayGanador(UserIndex, UserList(UserIndex).RetYA.RetUsu)
 End If
'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
'        Dim MiObj As Obj
'        Dim nPos As WorldPos
'        MiObj.ObjIndex = RandomNumber(554, 555)
'        MiObj.Amount = 1
'        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
'        Dim ManchaSangre As New cGarbage
'        ManchaSangre.Map = nPos.Map
'        ManchaSangre.X = nPos.X
'        ManchaSangre.Y = nPos.Y
'        Call TrashCollector.Add(ManchaSangre)
'End If
'<< Actualizamos clientes >>
'[Misery_Ezequiel 26/06/05]
UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
'[\]Misery_Ezequiel 26/06/05]
Exit Sub
ErrorHandler:
    Call LogError("Error en SUB USERDIE")
End Sub

'[\]MISERY_EZEQUIEL 11/07/05]
Sub MuereNpcByEze(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo errhandler
'   Call LogTarea("Sub MuereNpc")
Dim MiNPC As npc
   MiNPC = Npclist(NpcIndex)
    If (esPretoriano(NpcIndex) = 4) Then
        'seteamos todos estos 'flags' acorde para que cambien solos de alcoba
        Dim i As Integer
        Dim j As Integer
        Dim NPCI As Integer
        For i = 8 To 90
            For j = 8 To 90
                NPCI = MapData(Npclist(NpcIndex).Pos.Map, i, j).NpcIndex
                If NPCI > 0 Then
                    If esPretoriano(NPCI) > 0 Then
                        Npclist(NPCI).Invent.ArmourEqpSlot = IIf(Npclist(NpcIndex).Pos.X > 50, 1, 5)
                    End If
                End If
            Next j
        Next i
        Call CrearClanPretoriano(MAPA_PRETORIANO, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
    End If
   'Quitamos el npc
   Call QuitarNPCByEze(NpcIndex)
   If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MiNPC.flags.Snd3)
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        'El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMacotas > 0 Then
                Dim t As Integer
                For t = 1 To MAXMASCOTAS
                      If UserList(UserIndex).MascotasIndex(t) > 0 Then
                          If Npclist(UserList(UserIndex).MascotasIndex(t)).TargetNPC = NpcIndex Then
                                  Call FollowAmo(UserList(UserIndex).MascotasIndex(t))
                          End If
                      End If
                Next t
        End If
     '[KEVIN]
        '[Alejo] faltaba este if :P
        If MiNPC.flags.ExpCount > 0 Then
            '[Arkaris] Si está en party...
'            If UserList(UserIndex).PartyData.PIndex > 0 Then   partyexp
'                Call GivePartyXP(UserIndex, MiNPC.flags.ExpCount)
'            Else
            If UserList(UserIndex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(UserIndex, MiNPC.flags.ExpCount, MiNPC.Pos.Map, MiNPC.Pos.X, MiNPC.Pos.Y)
            Else
                Call AddtoVar(UserList(UserIndex).Stats.Exp, MiNPC.flags.ExpCount, MAXEXP)
                Call Senddata(ToIndex, UserIndex, 0, "||Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia." & FONTTYPE_FIGHT)
'            End If
            End If
            MiNPC.flags.ExpCount = 0
            '[/Arkaris]
        Else
            Call Senddata(ToIndex, UserIndex, 0, "||No has ganado experiencia al matar la criatura." & FONTTYPE_FIGHT)
        End If
        Call Senddata(ToIndex, UserIndex, 0, "Y25")
        Call AddtoVar(UserList(UserIndex).Stats.NPCsMuertos, 1, 32000)
        If MiNPC.Stats.Alineacion = 0 Then
              If MiNPC.Numero = Guardias Then
                    Call VolverCriminal(UserIndex)
              End If
              If MiNPC.MaestroUser = 0 Then
                    Call AddtoVar(UserList(UserIndex).Reputacion.AsesinoRep, vlASESINO, MAXREP)
              End If
        ElseIf MiNPC.Stats.Alineacion = 1 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 2 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, vlASESINO / 2, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 4 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        End If
        If Not Criminal(UserIndex) And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)
        'Controla el nivel del usuario
        Call CheckUserLevel(UserIndex)
   End If ' Userindex > 0
   If MiNPC.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCTirarOro(MiNPC, UserIndex)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
   End If
'ReSpawn o no
Call ReSpawnNpcByEze(MiNPC)
Exit Sub
errhandler:
    Call LogError("Error en MuereNpcByEze")
End Sub

Sub QuitarNPCByEze(ByVal NpcIndex As Integer)
On Error GoTo errhandler
Npclist(NpcIndex).flags.NPCActive = False
If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
    Call EraseNPCChar(ToMap, 0, Npclist(NpcIndex).Pos.Map, NpcIndex)
End If
'Nos aseguramos de que el inventario sea removido...
'asi los lobos no volveran a tirar armaduras ;))
Call ResetNpcInv(NpcIndex)
Call ResetNpcFlags(NpcIndex)
Call ResetNpcCounters(NpcIndex)
Call ResetNpcMainInfo(NpcIndex)
If NpcIndex = LastNPC Then
    Do Until Npclist(LastNPC).flags.NPCActive
        LastNPC = LastNPC - 1
        If LastNPC < 1 Then Exit Do
    Loop
End If
If NumNPCs <> 0 Then
    NumNPCs = NumNPCs - 1
End If
Exit Sub
errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPCByEze")
End Sub

Sub ReSpawnNpcByEze(MiNPC As npc)
If (MiNPC.flags.Respawn = 0) Then Call CrearNPCByEze(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)
End Sub

Sub CrearNPCByEze(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC
Dim Pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim Map As Integer
Dim X As Integer
Dim Y As Integer
Dim UserIndex As Integer

nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
If nIndex > MAXNPCS Then Exit Sub
'Necesita ser respawned en un lugar especifico
If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
    'Map = OrigPos.Map
    Map = MapByEze
    X = OrigPos.X
    Y = OrigPos.Y
    Npclist(nIndex).Orig = OrigPos
    Npclist(nIndex).Pos = OrigPos
Else
    Pos.Map = mapa
    altpos.Map = mapa
    Do While Not PosicionValida
        Randomize (Timer)
        Pos.X = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en x
        Pos.Y = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en y
        Call ClosestLegalPos(Pos, newpos) ' 'Nos devuelve la posicion valida mas cercana
        If newpos.X <> 0 Then altpos.X = newpos.X
        If newpos.Y <> 0 Then altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida) And _
           Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
            'Asignamos las nuevas coordenas solo si son validas
            'Npclist(nIndex).Pos.Map = newpos.Map
            Npclist(nIndex).Pos.Map = MapByEze
            Npclist(nIndex).Pos.X = newpos.X
            Npclist(nIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0
        End If
        'for debug
        Iteraciones = Iteraciones + 1
        If Iteraciones > MAXSPAWNATTEMPS Then
            If altpos.X <> 0 And altpos.Y <> 0 Then
                'Map = altpos.Map
                MapByEze = altpos.Map
                X = altpos.X
                Y = altpos.Y
                'Npclist(nIndex).Pos.Map = Map
                Npclist(nIndex).Pos.Map = MapByEze
                Npclist(nIndex).Pos.X = X
                Npclist(nIndex).Pos.Y = Y
                Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)
                Exit Sub
            Else
                altpos.X = 50
                altpos.Y = 50
                Call ClosestLegalPos(altpos, newpos)
                If newpos.X <> 0 And newpos.Y <> 0 Then
                    'Npclist(nIndex).Pos.Map = newpos.Map
                    Npclist(nIndex).Pos.Map = MapByEze
                    Npclist(nIndex).Pos.X = newpos.X
                    Npclist(nIndex).Pos.Y = newpos.Y
                    Call MakeNPCChar(ToMap, 0, newpos.Map, nIndex, newpos.Map, newpos.X, newpos.Y)
                    Exit Sub
                Else
                    Call QuitarNPC(nIndex)
                    Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                    Exit Sub
                End If
            End If
        End If
    Loop
    'asignamos las nuevas coordenas
    'Map = newpos.Map
    MapByEze = newpos.Map
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y
End If
'Crea el NPC
Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)
End Sub
'[\]MISERY_EZEQUIEL 11/07/05]

Sub SendUserMana(ByVal UserIndex As Integer)
Call Senddata(ToIndex, UserIndex, 0, "MAN" & UserList(UserIndex).Stats.MinMAN)
End Sub

Sub SendUserVida(ByVal UserIndex As Integer)
Call Senddata(ToIndex, UserIndex, 0, "VID" & UserList(UserIndex).Stats.MinHP)
End Sub

Sub SendUserEsta(ByVal UserIndex As Integer)
Call Senddata(ToIndex, UserIndex, 0, "ENE" & UserList(UserIndex).Stats.MinSta)
End Sub

Public Function ITS(ByVal Var As Integer) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    'No aceptamos valores que utilicen los últimos bits, pues los usamos como flag para evitar Chr(0)s
    If Var > &H3FFF Then GoTo errhandler
    
    Dim Temp As String
    
    'Si el primer Byte es cero
    If (Var And &HFF00) = 0 Then _
        Var = Var Or &H4000
    
    'Si el segundo Byte es cero
    If (Var And &HFF) = 0 Then _
        Var = Var Or &H8001
    
    'Convertimos a hexa
    Temp = hex$(Var)
    
    'Nos aseguramos tenga 4 Bytes de largo
    While Len(Temp) < 4
        Temp = "0" & Temp
    Wend
    
    'Convertimos a string
    ITS = Chr$(val("&H" & Left$(Temp, 2))) & Chr$(val("&H" & Right$(Temp, 2)))
Exit Function
errhandler:
End Function
