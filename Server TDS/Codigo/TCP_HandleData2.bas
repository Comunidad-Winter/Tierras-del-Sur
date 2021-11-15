Attribute VB_Name = "TCP_HandleData2"
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

Public Sub HandleData_2(ByVal UserIndex As Integer, rdata As String, ByRef Procesado As Boolean)
Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim Name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer

Procesado = True 'ver al final del sub

    Select Case UCase$(rdata)
         
      ' Para los gms ' Marche
        Case "/CONSOL"
        If UserList(UserIndex).flags.Privilegios > 0 Then
        Call Senddata(ToIndex, UserIndex, 0, "FF")
        End If
        '[Misery_Ezequiel And Marche 09/07/05]
       ' Case "/ABANDONAR"
        '     'Se asegura que el target es un npc
            ' If UserList(UserIndex).flags.TargetNPC = 0 Then
         '       Call Senddata(ToIndex, UserIndex, 0, "Y351")
          '      Exit Sub
           '  End If
             
             'If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              '  UserIndex Then Exit Sub
       
             
            'For i = 1 To UserList(UserIndex).NroMacotas
               ' If UserList(UserIndex).MascotasIndex(i) = UserList(UserIndex).flags.TargetNPC Then
             '         UserList(UserIndex).MascotasType(i) = 0
              '      UserList(UserIndex).MascotasIndex(i) = 0
                'End If
            'Next i
          
         '   UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
            
          
          '  Call MuereNpc(UserList(UserIndex).flags.TargetNPC, 0)
           ' Call Senddata(ToIndex, UserIndex, 0, "Y355")
            'Exit Sub
        '[\]Misery_Ezequiel And Marche 09/07/05]
        Case "/ONLINE"
            N = 0
            tStr = ""
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios < 1 Then
                    N = N + 1
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            Next LoopC
            If Len(tStr) > 2 Then
                tStr = Left(tStr, Len(tStr) - 2)
            End If
            Call Senddata(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
            Call Senddata(ToIndex, UserIndex, 0, "||Número de usuarios: " & N & FONTTYPE_INFO)
            Exit Sub
        Case "/SALIR"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y255")
                Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call Senddata(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "Y129")
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call Senddata(ToIndex, UserIndex, 0, "Y256")
                Call FinComerciarUsu(UserIndex)
            End If
  
            Cerrar_Usuario (UserIndex)
            Exit Sub
    ''    Case "/SALIRCLAN"
    ''        If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    ''            Call EacharMember(UserIndex, UserList(UserIndex).Name)
    ''            UserList(UserIndex).GuildInfo.GuildName = ""
    ''            UserList(UserIndex).GuildInfo.EsGuildLeader = 0
    ''        End If
    ''        Exit Sub
        Case "/ACTIVAR"
       If UserList(UserIndex).Stats.GlobAl = 2 Then
       UserList(UserIndex).Stats.GlobAl = 0
       Call Senddata(ToIndex, UserIndex, 0, "||No podras ni recibiras mensajes globales.." & FONTTYPE_INFO)
       Else
       UserList(UserIndex).Stats.GlobAl = 2
       Call Senddata(ToIndex, UserIndex, 0, "||Podras escribir y recibir mensajes globales." & FONTTYPE_INFO)
       End If
       Exit Sub
        
        Case "/FUNDARCLAN"
            If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y257")
                Exit Sub
            End If
            If CanCreateGuild(UserIndex) Then
                Call Senddata(ToIndex, UserIndex, 0, "SHOWFUN" & FONTTYPE_INFO)
            End If
            Exit Sub
            
        '[Barrin 1-12-03]
        Case "/SALIRCLAN"
            If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y258")
                      Exit Sub
            ElseIf UserList(UserIndex).GuildInfo.GuildName = "" Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y223")
                      Exit Sub
            Else
                Call Senddata(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & " decidió dejar al clan." & FONTTYPE_GUILD)
                Dim oGuild As cGuild
                Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
                Call oGuild.RemoveMember(UserList(UserIndex).Name)
                Call AddtoVar(UserList(UserIndex).GuildInfo.Echadas, 1, 1000)
                UserList(UserIndex).GuildInfo.GuildPoints = 0
                UserList(UserIndex).GuildInfo.GuildName = ""
                Errorestaen = "salirclan"
                '[Wizard 03/09/05] Forma burda de actualizar el nick ahorrar lineas, anchodebanda y clonacion de pjs jajajaja pero = es feo.
                Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, False)
            '''''''''''''''''
            End If
            Exit Sub
        '[/Barrin 1-12-03]
            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y4")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y7")
                      Exit Sub
            End If
            Select Case Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case NPCTYPE_BANQUERO
                'If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                '      Call Senddata(ToIndex, UserIndex, 0, "Y295")
                '      CloseSocket (UserIndex)
                '      Exit Sub
                'End If
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
            Case NPCTYPE_TIMBERO
                If UserList(UserIndex).flags.Privilegios > 0 Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    N = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call Senddata(ToIndex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & N & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call Senddata(ToIndex, UserIndex, 0, "Y3")
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y4")
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                          Call Senddata(ToIndex, UserIndex, 0, "Y5")
                          Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Sub
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y4")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y5")
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y4")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y5")
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                    Call Senddata(ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y259")
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y260")
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y260")
                        UserList(UserIndex).flags.Descansar = False
                        Call Senddata(ToIndex, UserIndex, 0, "DOK")
                        Exit Sub
                    End If
                    Call Senddata(ToIndex, UserIndex, 0, "Y261")
            End If
            Exit Sub
        Case "/MEDITAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
            End If
            '[Misery_Ezequiel 12/06/05]
            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y45")
                Exit Sub
            End If
            '[\]Misery_Ezequiel 12/06/05]
            Call Senddata(ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
               Call Senddata(ToIndex, UserIndex, 0, "Y262")
            Else
               Call Senddata(ToIndex, UserIndex, 0, "Y216")
            End If
            UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call Senddata(ToIndex, UserIndex, 0, "||Te estás concentrando. En " & TIEMPO_INICIOMEDITAR & " segundos comenzarás a meditar." & FONTTYPE_INFO)
                UserList(UserIndex).Char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 15 Then
                    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
                'Nacho 09/04/05
                ElseIf UserList(UserIndex).Stats.ELV < 45 Then
                    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGRANDE
              ElseIf UserList(UserIndex).Stats.ELV >= 45 Then
                    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARGIGANTE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGIGANTE
                '/Nacho 09/04/05
                End If
            Else
                UserList(UserIndex).Counters.bPuedeMeditar = False
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call Senddata(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & 0 & "," & 0)
            End If
            Exit Sub
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y4")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 1 _
           Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y9")
               Exit Sub
           End If
           'If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
           '    Call Senddata(ToIndex, UserIndex, 0, "Y295")
           '    CloseSocket (UserIndex)
           '    Exit Sub
           'End If
           Call RevivirUsuario(UserIndex)
           Call Senddata(ToIndex, UserIndex, 0, "Y296")
           Exit Sub
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y4")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 1 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y8")
               Exit Sub
           End If
           UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
           Call SendUserStatsBox(val(UserIndex))
           Call Senddata(ToIndex, UserIndex, 0, "Y17")
           Exit Sub
        Case "/AYUDA"
           Call SendHelp(UserIndex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        '[Misery_Ezequiel 12/06/05]
        Case "/SEG"
            If UserList(UserIndex).flags.Seguro Then
                Call Senddata(ToIndex, UserIndex, 0, "SEGOFF")
            Else
                Call Senddata(ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
        '[\]Misery_Ezequiel 12/06/05]
    
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
            End If
            If UserList(UserIndex).flags.Comerciando Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y27")
                    Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios = 1 Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                  '¿El NPC puede comerciar?
                  If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                     If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                     Exit Sub
                  End If
                  If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y7")
                      Exit Sub
                  End If
                  'Iniciamos la rutina pa' comerciar.
                  Call IniciarCOmercioNPC(UserIndex)
             '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Call SendData(ToIndex, UserIndex, 0, "||COMERCIO SEGURO ENTRE USUARIOS TEMPORALMENTE DESHABILITADO" & FONTTYPE_INFO)
                'Exit Sub
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y264")
                    Exit Sub
                End If
                 If UserList(UserIndex).flags.Navegando = 1 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y297")
                    Exit Sub
                 End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y265")
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y298")
                    Exit Sub
                End If
                '[Wizard 03/09/05]Es Consejero????o yo Soy consejero?
                If UserList(UserIndex).flags.Privilegios = 1 Or UserList(UserList(UserIndex).flags.TargetUser).flags.Privilegios = 1 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y373")
                    Exit Sub
                End If
                '[/Wizard]
                'Ya ta comerciando ? es con migo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y266")
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).Name
                UserList(UserIndex).ComUsu.cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y299")
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                  If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y7")
                      Exit Sub
                  End If
                  If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 4 Then
                    Call IniciarDeposito(UserIndex)
                  Else
                    Exit Sub
                  End If
            Else
              Call Senddata(ToIndex, UserIndex, 0, "Y299")
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y4")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           '[Wizard]
           'Si tiene clan no puede alterar su faccion
           If UserList(UserIndex).GuildInfo.GuildName <> "" Then
            Call Senddata(ToIndex, UserIndex, 0, "Y380")
            Exit Sub
           End If
           
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y8")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y4")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y5")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y4")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y8")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(UserIndex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call Senddata(ToIndex, UserIndex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call Senddata(ToIndex, UserIndex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
          '[Marche 9-4-05]
            Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Exit Sub
            
            Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub
            
            Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
            
            Case "/ONLINEPARTY"
            Call OnlineParty(UserIndex)
            Exit Sub
            
            Case "/CPARTY"
            Call CParty(UserIndex)
            Exit Sub
    End Select
    '[Barrin 1-12-03]
    If UCase$(Left$(rdata, 6)) = "/CMSG " Then
        'If UserList(UserIndex).flags.Muerto = 1 Then
        '    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
        '    Exit Sub
        'End If
        If Len(UserList(UserIndex).GuildInfo.GuildName) = 0 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y267")
                Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 6)
        If rdata <> "" And UserList(UserIndex).GuildInfo.GuildName <> "" Then
            'Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_GUILDMSG)
            Call Senddata(ToDiosesYclan, UserIndex, 0, "PT" & UserList(UserIndex).Name & "> " & rdata)
        End If
        Exit Sub
    End If
    
     If UCase$(Left$(rdata, 6)) = "/PMSG " Then
        Call mdParty.BroadCastParty(UserIndex, UserList(UserIndex).Name & "> " & Mid$(rdata, 7) & FONTTYPE_TALK)
        Exit Sub
    End If

    If UCase$(rdata) = "/ONLINECLAN" Then
    
        If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        tStr = Left$(tStr, Len(tStr) - 2)
        Call Senddata(ToIndex, UserIndex, 0, "||Usuarios de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    '[/Barrin 1-12-03]
    
    '[yb]
     If UCase$(Left$(rdata, 6)) = "/BMSG " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call Senddata(ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call Senddata(ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rdata, 6)) = "/GMSG " And UserList(UserIndex).flags.Privilegios > 0 Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Call LogGM(UserList(UserIndex).Name, "Mensaje a Gms:" & rdata, (UserList(UserIndex).flags.Privilegios = 1))
        If rdata <> "" Then
            Call Senddata(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~150~150~150~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rdata, 3))
        Case "/GM"
            If Not Ayuda.Existe(UserList(UserIndex).Name) Then
                Call Senddata(ToIndex, UserIndex, 0, "Y268")
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
            Else
                Call Ayuda.Quitar(UserList(UserIndex).Name)
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
                Call Senddata(ToIndex, UserIndex, 0, "Y269")
            End If
            Exit Sub
    End Select

    Select Case UCase(Left(rdata, 5))
        Case "/BUG "
            N = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
            Print #N,
            Print #N,
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N, "Usuario:" & UserList(UserIndex).Name & "  Fecha:" & Date & "    Hora:" & Time
            Print #N, "########################################################################"
            Print #N, "BUG:"
            Print #N, Right$(rdata, Len(rdata) - 5)
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N,
            Print #N,
            Close #N
            Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 6))
        'Case "/YUCH "
         '   rdata = Right$(rdata, Len(rdata) - 7)
          '  Call Senddata(ToAdmins, 0, 0, "||Servidor> " & ReadField(1, rdata, 44) & ". Fue echado del server por " & ReadField(2, rdata, 44) & "." & FONTTYPE_SERVER)
           ' LogCheats (ReadField(1, rdata, 44) & ". Fue echado del server por " & ReadField(2, rdata, 44) & ".")
            'UserList(UserIndex).Stats.Veceshechado = UserList(UserIndex).Stats.Veceshechado + 1
            
           ' If Len(ReadField(2, rdata, 44)) > 3 Then
            '
             '   If UserList(UserIndex).Stats.Veceshechado = 1 Then
              '  Call Senddata(ToIndex, UserIndex, 0, "ERREs la " & UserList(UserIndex).Stats.Veceshechado & "º que eres expulsado por uso de cheats. La proxima vez seras baneado.")
               ' End If
                
               ' If UserList(UserIndex).Stats.Veceshechado > 1 Then
                'UserList(UserIndex).flags.Ban = 1
                'UserList(UserIndex).flags.Banrazon = UserList(UserIndex).flags.Banrazon & vbCrLf & "Sistema anticheat." & " " & Date
                'LogCheats (ReadField(1, rdata, 44) & ". Fue echado del server por " & ReadField(2, rdata, 44) & ".¡¡BANEADO!!")
                'Call SaveUser(UserIndex, UserList(UserIndex).Name)
                'End If
            'End If
            
           ' Call CloseSocket(UserIndex)
           ' Exit Sub
        Case "/MUY1 "
            rdata = Right$(rdata, Len(rdata) - 6)
            Call Senddata(ToAdmins, 0, 0, "|| Servidor> El sistema anti-cheats sospecha de " & rdata & " por uso de macro para tomar pociones." & FONTTYPE_SERVER)
       Case "/DESC "
            rdata = Right$(rdata, Len(rdata) - 6)
          If Not AsciiValidos(rdata) Then
               Call Senddata(ToIndex, UserIndex, 0, "Y270")
              Exit Sub
          End If
            If UserList(UserIndex).flags.Muerto = 1 Then
               Call Senddata(ToIndex, UserIndex, 0, "Y300")
               Exit Sub
           End If
           If Len(rdata) >= 255 Then
           Exit Sub
           End If
           UserList(UserIndex).Desc = rdata
            Call Senddata(ToIndex, UserIndex, 0, "Y271")
            Exit Sub
        Case "/VOTO "
        
        
                rdata = Right$(rdata, Len(rdata) - 6)
                Call ComputeVote(UserIndex, rdata)
                Exit Sub
              
        Case "/VOTC "
        
        If votacionabierta Then
        
        rdata = Right$(rdata, Len(rdata) - 6)
        If UserList(UserIndex).Name = UCase(rdata) Then
        Call Senddata(ToIndex, UserIndex, 0, "||No te puedes votar a ti mismo. " & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(UserIndex).Faccion.FuerzasCaos <> 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "||No puedes votar si no eres del Caos " & FONTTYPE_INFO)
            Exit Sub
        Else
        
            If rs.State = 1 Then
            rs.Close
            End If
                    
            sql = "select * from usuarios where NickB = '" & rdata & "'"
            rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False Then
            Else
            Call Senddata(ToIndex, UserIndex, 0, "||El Personaje no existe. " & FONTTYPE_INFO)
            rs.Close
            Exit Sub
            End If
        
            If UserList(UserIndex).Stats.VotC <> 1 Then
                If rs!ejercitocaosb <> 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "||El Personaje no es del caos. " & FONTTYPE_INFO)
                rs.Close
                Exit Sub
                Else
                rs!VotC = rs!VotC + 1
                UserList(UserIndex).Stats.VotC = 1
                rs.Update
                rs.Close
                Exit Sub
                End If
            Else
            Call Senddata(ToIndex, UserIndex, 0, "||Ya has votado. " & FONTTYPE_INFO)
            rs.Close
            Exit Sub
            End If
        End If
    Else
    Call Senddata(ToIndex, UserIndex, 0, "||La votación se encuentra cerrada. " & FONTTYPE_INFO)
    End If
    Exit Sub
    End Select
    
        Select Case UCase$(Left$(rdata, 7))
        Case "/CHEQUE" 'RETIRA ORO DE UN CHEQUE
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
               '       Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y4")
                  Exit Sub
             End If
             rdata = UCase$(Right$(rdata, Len(rdata) - 8))
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO Then
                Call Senddata(ToIndex, UserIndex, 0, "||Debes clickear a un banquero." & FONTTYPE_INFO)
                Exit Sub
            End If
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y5")
                  Exit Sub
             End If
             sql = "SELECT * FROM Cheques WHERE Codigo = '" & rdata & "'"
             rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
             If rs.EOF = False Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + rs!Dinero
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has cobrado el cheque por " & rs!Dinero & " monedas de oro." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
                conn.Execute "INSERT INTO Logs values('" & "Se retira el cheque " & rdata & " por " & rs!Dinero & " con el pj " & UserList(UserIndex).Name & "','" & Date & "')", , adExecuteNoRecords
                rs.Delete
             Else
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " El cheque que intenta retirar no existe." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
             End If
          
             rs.Close
             Call SendUserStatsBox(val(UserIndex))
             Exit Sub
     End Select
    Select Case UCase$(Left$(rdata, 7))
        Case "/R0TAR "
        Dim regla1 As Integer
        Dim regla2 As Integer
        Dim regla3 As Integer
        
        regla1 = Mid(rdata, 8, 1)
        regla2 = Mid(rdata, 9, 1)
        regla3 = Mid(rdata, 10, 1)
        rdata = Mid(rdata, 11, Len(rdata) - 10)
                
        If rdata = "" Then
        Exit Sub
        End If
        
        'UserList(UserIndex).Reglas.Invisib = regla1
        'UserList(UserIndex).Reglas.Element = regla2
        'UserList(UserIndex).Reglas.Estupidez = regla3
        'UserList(UserIndex).RetYA.RetOro = CLng(val(rdata))
                
        Call PuedeCrear(UserIndex, CLng(val(rdata)))
                
    
        Case "/R3TAR"
       '     Call Senddata(ToIndex, UserIndex, 0, "FZ")
        
        Case "/RETAR"
             'Dim ororeto As Single
        '     If UserList(UserIndex).flags.TargetUser > 0 Then
        '
        '        If UserList(UserList(UserIndex).flags.TargetUser).RetYA.RettNick = UserList(UserIndex).Name Then
         '           ororeto = UserList(UserList(UserIndex).flags.TargetUser).RetYA.RetOro
        '            UserList(UserIndex).Reglas.Invisib = UserList(UserList(UserIndex).flags.TargetUser).Reglas.Invisib
        ''            UserList(UserIndex).Reglas.Element = UserList(UserList(UserIndex).flags.TargetUser).Reglas.Element
          '          UserList(UserIndex).Reglas.Estupidez = UserList(UserList(UserIndex).flags.TargetUser).Reglas.Estupidez
          '          UserList(UserIndex).RetYA.RetOro = val(ororeto)
          '          Call Puedejugar(UserIndex, CLng(val(ororeto)))
          '      Else
          '           Call Senddata(ToIndex, UserIndex, 0, "||Esta persona no quiere retarte." & FONTTYPE_INFO)
          '      End If
          '  End If
            
           ' Case "/AETAR"
           ' rdata = Right$(rdata, Len(rdata) - 8)
           ' PuedeCrearDuplas
            
      
            
     End Select
    
    Select Case UCase$(Left$(rdata, 8))
        Case "/PASSWD "
            rdata = Right$(rdata, Len(rdata) - 8)
            If Len(rdata) < 6 Then
                 Call Senddata(ToIndex, UserIndex, 0, "Y272")
            Else
                 Call Senddata(ToIndex, UserIndex, 0, "Y273")
                 UserList(UserIndex).Password = rdata
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rdata = Right(rdata, Len(rdata) - 9)
            tLong = CLng(val(rdata))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y3")
            ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call Senddata(ToIndex, UserIndex, 0, "Y4")
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y5")
            ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_TIMBERO Then
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            ElseIf N < 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            ElseIf N > 5000 Then
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call SendUserStatsBox(UserIndex)
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y4")
                  Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 5 Then
                'Se quiere retirar de la armada
                'Wizard 07/09/05 Cambie algunos mensajes..
                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                    '[Wizard]
                     If UserList(UserIndex).GuildInfo.GuildName <> "" Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y380")
                        Exit Sub
                    End If
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(UserIndex)
                        Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Vete, pero recuerda, no podrás volver a unirte a la armada real" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex)
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    End If
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    If UserList(UserIndex).GuildInfo.GuildName <> "" Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y380")
                        Exit Sub
                    End If
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Lárgate de aquí escoria, mas nunca vuelvas, porque no te aceptare de nuevo en mis filas." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                End If
                '/Wizard
            Exit Sub
        End If
             If Len(rdata) = 8 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y274")
                Exit Sub
             End If
             rdata = Right$(rdata, Len(rdata) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y5")
                  Exit Sub
             End If
             'If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
             '     Call Senddata(ToIndex, UserIndex, 0, "Y295")
             '     CloseSocket (UserIndex)
             '     Exit Sub
             'End If
             If val(rdata) > 0 And val(rdata) <= UserList(UserIndex).Stats.Banco Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
                  Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
             Else
                  Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
             End If
             Call SendUserStatsBox(val(UserIndex))
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y3")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y4")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call Senddata(ToIndex, UserIndex, 0, "Y5")
                      Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call Senddata(ToIndex, UserIndex, 0, "Y5")
                  Exit Sub
            End If
            If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                  Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
            Else
                  Call Senddata(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
            End If
            Call SendUserStatsBox(val(UserIndex))
            Exit Sub
        Case "/DENUNCIAR "
            rdata = Right$(rdata, Len(rdata) - 11)
            Call Senddata(ToAdmins, 0, 0, "|| " & LCase$(UserList(UserIndex).Name) & " DENUNCIA: " & rdata & FONTTYPE_GUILDMSG)
            Call Senddata(ToIndex, UserIndex, 0, "Y301")
            Exit Sub
    End Select

    Debug.Print (UCase$(Left$(rdata, 4)))
    'marche 4-9
    ' aca ponemos todos los comandos del party
    Select Case UCase$(Left$(rdata, 4))
    Case "/AP "
            rdata = LTrim(Right$(rdata, Len(rdata) - 3))
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y302")
            End If
            Exit Sub
     Case "/EP "
            rdata = Right$(rdata, Len(rdata) - 4)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y302")
            End If
            Exit Sub
        Case "/PL"
            rdata = Right$(rdata, Len(rdata) - 3)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y302")
            End If
            Exit Sub
    End Select
Procesado = False
End Sub
'********************Misery_Ezequiel 28/05/05********************'
