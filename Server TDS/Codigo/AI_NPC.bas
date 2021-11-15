Attribute VB_Name = "AI"
'Argentum Online 0.11.20
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

Public Const ESTATICO = 1
Public Const MUEVE_AL_AZAR = 2
Public Const NPC_MALO_ATACA_USUARIOS_BUENOS = 3
Public Const NPCDEFENSA = 4
Public Const GUARDIAS_ATACAN_CRIMINALES = 5
Public Const SIGUE_AMO = 8
Public Const NPC_ATACA_NPC = 9
Public Const NPC_PATHFINDING = 10
Public Const ELEMENTALFUEGO = 93
Public Const ELEMENTALTIERRA = 94
Public Const ELEMENTALAGUA = 92

'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo AI_NPC
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'AI de los NPC
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
Private Sub GuardiasAI(ByVal NpcIndex As Integer, Optional ByVal DelCaos As Boolean = False)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer

For headingloop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or headingloop = Npclist(NpcIndex).Char.Heading Then
        Call HeadtoPos(headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
            If UI > 0 Then
                  If UserList(UI).flags.Muerto = 0 Then
                         '�ES CRIMINAL?
                         If Not DelCaos Then
                            If Criminal(UI) Then
                                   Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, headingloop)
                                   Call NpcAtacaUser(NpcIndex, UI)
                                   Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name _
                                      And Not Npclist(NpcIndex).flags.Follow Then
                                  Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, headingloop)
                                  Call NpcAtacaUser(NpcIndex, UI)
                                  Exit Sub
                            End If
                        Else
                            If Not Criminal(UI) Then
                                   Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, headingloop)
                                   Call NpcAtacaUser(NpcIndex, UI)
                                   Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name _
                                      And Not Npclist(NpcIndex).flags.Follow Then
                                  Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, headingloop)
                                  Call NpcAtacaUser(NpcIndex, UI)
                                  Exit Sub
                            End If
                        End If
                  End If
            End If
        End If
    End If  'not inmovil
Next headingloop
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
Dim NPCI As Integer
Dim atacoPJ As Boolean

atacoPJ = False
For headingloop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).Char.Heading = headingloop Then
        Call HeadtoPos(headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
            NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
            If UI > 0 And Not atacoPJ Then
            'fruta
                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Mimetizado = 0 Then
                    atacoPJ = True
                    If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                        Dim k As Integer
                        k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                        Call NpcLanzaUnSpell(NpcIndex, UI)
                    End If
                    Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, headingloop)
                    Call NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex)
                    Exit Sub
                End If
            ElseIf NPCI > 0 Then
                    If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                        Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, headingloop)
                        Call NpcAtacaNpc(NpcIndex, NPCI, False)
                        Exit Sub
                    End If
            End If
        End If
    End If  'inmo
Next headingloop
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For headingloop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).Char.Heading = headingloop Then
        Call HeadtoPos(headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
            If UI > 0 Then
                If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                    If UserList(UI).flags.Muerto = 0 Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, headingloop)
                            Call NpcAtacaUser(NpcIndex, UI)
                            Exit Sub
                    End If
                End If
            End If
        End If
    End If
Next headingloop
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case NORTH
            SignoNS = -1
            SignoEO = 0
        Case EAST
            SignoNS = 0
            SignoEO = 1
        Case SOUTH
            SignoNS = 1
            SignoEO = 0
        Case WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * 10 Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * 10 Step IIf(SignoEO = 0, 1, SignoEO)
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                   UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                   If UI > 0 Then
                      If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                            If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            Exit Sub
                      End If
                   End If
            End If
        Next X
    Next Y
Else
    For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
        For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                   UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                   If UI > 0 Then
                      If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                            If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                      End If
                   End If
            End If
        Next X
    Next Y
End If
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Or Npclist(NpcIndex).flags.Paralizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case NORTH
            SignoNS = -1
            SignoEO = 0
        Case EAST
            SignoNS = 0
            SignoEO = 1
        Case SOUTH
            SignoNS = 1
            SignoEO = 0
        Case WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * 10 Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * 10 Step IIf(SignoEO = 0, 1, SignoEO)
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                If UI > 0 Then
                    If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                        If Npclist(NpcIndex).MaestroUser > 0 Then
                            If Not Criminal(Npclist(NpcIndex).MaestroUser) And Not Criminal(UI) And (UserList(Npclist(NpcIndex).MaestroUser).flags.Seguro Or UserList(Npclist(NpcIndex).MaestroUser).Faccion.ArmadaReal = 1) Then
                                Call Senddata(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||La mascota no atacar� a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado" & FONTTYPE_INFO)
                                Npclist(NpcIndex).flags.AttackedBy = ""
                                Exit Sub
                            End If
                        End If
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                             If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                             End If
                             Exit Sub
                        End If
                    End If
                End If
            End If
        Next X
    Next Y
Else

 'Aca creo que mande mucha pero muuucha fruta.
 If UserList(NameIndex(Npclist(NpcIndex).flags.AttackedBy)).flags.Mimetizado = 1 Then
 Call RestoreOldMovement(NpcIndex)
 Exit Sub
 End If
 'aca termino la fruta

    For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
        For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                If UI > 0 Then
                    If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                        If Npclist(NpcIndex).MaestroUser > 0 Then
                               If Not Criminal(Npclist(NpcIndex).MaestroUser) And Not Criminal(UI) And (UserList(Npclist(NpcIndex).MaestroUser).flags.Seguro Or UserList(Npclist(NpcIndex).MaestroUser).Faccion.ArmadaReal = 1) Then
                                Call Senddata(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||La mascota no atacar� a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado" & FONTTYPE_INFO)
                                Npclist(NpcIndex).flags.AttackedBy = ""
                                Call FollowAmo(NpcIndex)
                                Exit Sub
                            End If
                        End If
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                             If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                             End If
                             tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                             Call MoveNPCChar(NpcIndex, tHeading)
                             Exit Sub
                        End If
                    End If
                End If
            End If
        Next X
    Next Y
End If
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
If Npclist(NpcIndex).MaestroUser = 0 Then
    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    Npclist(NpcIndex).flags.AttackedBy = ""
End If
End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
           If UI > 0 Then
                If Not Criminal(UI) Then
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next X
Next Y
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case NORTH
            SignoNS = -1
            SignoEO = 0
        Case EAST
            SignoNS = 0
            SignoEO = 1
        Case SOUTH
            SignoNS = 1
            SignoEO = 0
        Case WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * 10 Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * 10 Step IIf(SignoEO = 0, 1, SignoEO)
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
               If UI > 0 Then
                    If Criminal(UI) Then
                       If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
'                                  Dim k As Integer
'                                  k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            Exit Sub
                       End If
                    End If
               End If
            End If
        Next X
    Next Y
Else
    For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
        For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
               If UI > 0 Then
                    If Criminal(UI) Then
                       If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  'Dim k As Integer
                                  'k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Sub
                            tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                       End If
                    End If
               End If
            End If
        Next X
    Next Y
End If
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
            If Npclist(NpcIndex).Target = 0 And Npclist(NpcIndex).TargetNPC = 0 Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                If UI > 0 Then
                   If UserList(UI).flags.Muerto = 0 _
                   And UserList(UI).flags.Invisible = 0 _
                   And UI = Npclist(NpcIndex).MaestroUser _
                   And Distancia(Npclist(NpcIndex).Pos, UserList(UI).Pos) > 3 Then
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
            End If
        End If
    Next X
Next Y
Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim NI As Integer
Dim bNoEsta As Boolean

Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case NORTH
            SignoNS = -1
            SignoEO = 0
        Case EAST
            SignoNS = 0
            SignoEO = 1
        Case SOUTH
            SignoNS = 1
            SignoEO = 0
        Case WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * 10 Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * 10 Step IIf(SignoEO = 0, 1, SignoEO)
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               NI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex
               If NI > 0 Then
                    If Npclist(NpcIndex).TargetNPC = NI Then
                         bNoEsta = True
                         If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Then
                             Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                             If Npclist(NI).NPCtype = DRAGON Then
                                Npclist(NI).CanAttack = 1
                                Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                             End If
                         Else
                            'aca verificamosss la distancia de ataque
                            If Distancia(Npclist(NpcIndex).Pos, Npclist(NI).Pos) <= 3 Then
                                Call NpcAtacaNpc(NpcIndex, NI)
                            End If
                         End If
                         Exit Sub
                    End If
               End If
            End If
        Next X
    Next Y
Else
    For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
        For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               NI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex
               If NI > 0 Then
                    If Npclist(NpcIndex).TargetNPC = NI Then
                         bNoEsta = True
                         If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Or Npclist(NpcIndex).Numero = ELEMENTALTIERRA Or Npclist(NpcIndex).Numero = 110 Or Npclist(NpcIndex).Numero = 111 Then
                             'Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                             'If Npclist(NI).NPCtype = DRAGON Then
                              '  Npclist(NI).CanAttack = 1
                               ' Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                             'End If
                         Else
                            'aca verificamosss la distancia de ataque
                            If Distancia(Npclist(NpcIndex).Pos, Npclist(NI).Pos) <= 3 Then
                                Call NpcAtacaNpc(NpcIndex, NI)
                            End If
                         End If
                         
                         If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Sub
                         
                           tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex).Pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         Exit Sub
                    End If
               End If
            End If
        Next X
    Next Y
End If

If Not bNoEsta Then
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call FollowAmo(NpcIndex)
    Else
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    End If
End If
    
End Sub

Function NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If Npclist(NpcIndex).MaestroUser = 0 Then
            'Busca a alguien para atacar
            '�Es un guardia?
  
            If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                    Call GuardiasAI(NpcIndex)
            ElseIf Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIASCAOS Then
                    Call GuardiasAI(NpcIndex, True)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion <> 0 Then
                    Call HostilMalvadoAI(NpcIndex)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Call HostilBuenoAI(NpcIndex)
            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case Npclist(NpcIndex).Movement
            Case MUEVE_AL_AZAR
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                    Call PersigueCriminal(NpcIndex)
                ElseIf Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIASCAOS Then
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                    Call PersigueCiudadano(NpcIndex)
                Else
                        If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                End If
            'Va hacia el usuario cercano
            Case NPC_MALO_ATACA_USUARIOS_BUENOS
                Call IrUsuarioCercano(NpcIndex)
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            'Persigue criminales
            Case GUARDIAS_ATACAN_CRIMINALES
                Call PersigueCriminal(NpcIndex)
            Case SIGUE_AMO
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                Call SeguirAmo(NpcIndex)
                If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                End If
            Case NPC_ATACA_NPC
                Call AiNpcAtacaNpc(NpcIndex)
            Case NPC_PATHFINDING
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    'Existe el camino?
                    If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))
                    End If
                Else
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        Npclist(NpcIndex).PFINFO.PathLenght = 0
                    End If
                End If
        End Select
Exit Function
ErrorHandler:

End Function

Function UserNear(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
UserNear = Not Int(Distance(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.X, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.Y)) > 1
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
    ReCalculatePath = True
ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
    ReCalculatePath = True
End If
End Function

Function SimpleAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Old Ore4 AI function
'#################################################################
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

For Y = Npclist(NpcIndex).Pos.Y - 5 To Npclist(NpcIndex).Pos.Y + 5    'Makes a loop that looks at
    For X = Npclist(NpcIndex).Pos.X - 5 To Npclist(NpcIndex).Pos.X + 5   '5 tiles in every direction
           'Make sure tile is legal
            If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                'look for a user
                If MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex > 0 Then
                    'Move towards user
                    tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                    MoveNPCChar NpcIndex, tHeading
                    'Leave
                    Exit Function
                End If
            End If
    Next X
Next Y
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################
Dim tmpPos As WorldPos
Dim tHeading As Byte

tmpPos.Map = Npclist(NpcIndex).Pos.Map
tmpPos.X = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).Y ' invert� las coordenadas
tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).X
'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
tHeading = FindDirection(Npclist(NpcIndex).Pos, tmpPos)
MoveNPCChar NpcIndex, tHeading
Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1
End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock / 11-07-02
'www.geocities.com/gmorgolock
'morgolock@speedy.com.ar
'This function seeks the shortest path from the Npc
'to the user's location.
'#################################################################
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer


For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
     For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10   '5 tiles in every direction
         'Make sure tile is legal
         If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
             'look for a user
             If MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex > 0 Then
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                  If UserList(tmpUserIndex).flags.Muerto = 0 And UserList(tmpUserIndex).flags.Invisible = 0 And UserList(tmpUserIndex).flags.Mimetizado = 0 Then
                    'We have to invert the coordinates, this is because
                    'ORE refers to maps in converse way of my pathfinding
                    'routines.
                    Npclist(NpcIndex).PFINFO.Target.X = UserList(tmpUserIndex).Pos.Y
                    Npclist(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).Pos.X 'ops!
                    Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                    Call SeekPath(NpcIndex)
                    Exit Function
                  End If
             End If
         End If
     Next X
 Next Y
End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
If UserList(UserIndex).flags.Invisible = 1 Then Exit Sub
Dim k As Integer
k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))
End Sub
'********************Misery_Ezequiel 28/05/05********************'
