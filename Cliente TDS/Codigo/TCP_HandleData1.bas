Attribute VB_Name = "TCP_HandleData1"
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

Public Sub HandleData_1(ByVal UserIndex As Integer, rdata As String, ByRef Procesado As Boolean)


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

    Select Case UCase$(Left$(rdata, 1))
        Case ";" 'Hablar
            If UserList(UserIndex).flags.Muerto = 1 Then
                  '  Call SendData(ToDeadArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & rdata & "°" & str(ind))
            End If
            rdata = Right$(rdata, Len(rdata) - 1)
            If InStr(rdata, "°") Then
                Exit Sub
            End If
        
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Then
                Call LogGM(UserList(UserIndex).Name, "Dijo: " & rdata, True)
            End If
            
            ind = UserList(UserIndex).Char.charindex
            '[CDT 17-02-2004]
            rdata = " " & rdata & " "
            rdata = Replace(rdata, " pt", " capo")
            rdata = Replace(rdata, " nw", " pro")
            rdata = Replace(rdata, " PT", " CAPO")
            rdata = Replace(rdata, " NW", " PRO")
            rdata = Replace(rdata, " Pt", " Capo")
            rdata = Replace(rdata, " Nw", " Pro")
            rdata = Replace(rdata, " pT", " CaPo")
            rdata = Replace(rdata, " nW", " PrO")
            rdata = Mid(rdata, 2, Len(rdata) - 2)
            '[/CDT]
            
            'piedra libre para todos los compas!
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                Call SendData(ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
            End If
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToDeadArea, UserIndex, UserList(UserIndex).Pos.Map, "||12632256°" & rdata & "°" & str(ind))
                'Call SendData(ToAdminsAreaButConsejeros, UserIndex, UserList(UserIndex).Pos.Map, "||12632256°" & rdata & "°" & str(ind))
            Else
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & rdata & "°" & str(ind))
            End If
            Exit Sub
        Case "-" 'Gritar
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 1)
            If InStr(rdata, "°") Then
                Exit Sub
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Then
                Call LogGM(UserList(UserIndex).Name, "Grito: " & rdata, True)
            End If
    
            'piedra libre para todos los compas!
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                Call SendData(ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
            End If
    
    
            ind = UserList(UserIndex).Char.charindex
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "°" & rdata & "°" & str(ind))
            Exit Sub
        Case "\" 'Susurrar al oido
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 1)
            tName = ReadField(1, rdata, 32)
            tIndex = NameIndex(tName)
            If tIndex <> 0 Then
                If UserList(tIndex).flags.Privilegios > 0 And UserList(UserIndex).flags.Privilegios = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y213")
                    Exit Sub
                End If
                If Len(rdata) <> Len(tName) Then
                    tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                If Not EstaPCarea(UserIndex, tIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas muy lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                ind = UserList(UserIndex).Char.charindex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                
                '[Consejeros]
                If UserList(UserIndex).flags.Privilegios = 1 Then
                    Call LogGM(UserList(UserIndex).Name, "Le dijo a '" & UserList(tIndex).Name & "' " & tMessage, True)
                End If
    
                Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                Call SendData(ToIndex, tIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                '[CDT 17-02-2004]
                If UserList(UserIndex).flags.Privilegios < 2 Then
                    Call SendData(ToAdminsAreaButConsejeros, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°" & "a " & UserList(tIndex).Name & "> " & tMessage & "°" & str(ind))
                End If
                '[/CDT]
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "Y214")
            Exit Sub
        Case "M" 'Moverse
            Dim dummy As Long
            Dim TempTick As Long
            If UserList(UserIndex).flags.TimesWalk >= 30 Then
                TempTick = GetTickCount And &H7FFFFFFF
                dummy = (TempTick - UserList(UserIndex).flags.StartWalk)
                If dummy < 6050 Then
                    If TempTick - UserList(UserIndex).flags.CountSH > 90000 Then
                        UserList(UserIndex).flags.CountSH = 0
                    End If
                    If Not UserList(UserIndex).flags.CountSH = 0 Then
                        dummy = 126000 / dummy
                        Call LogHackAttemp("Tramposo SH: " & UserList(UserIndex).Name & " , " & dummy)
                        Call SendData(ToAdmins, 0, 0, "||Servidor> " & UserList(UserIndex).Name & " ha sido echado por el servidor por posible uso de SH." & FONTTYPE_SERVER)
                        Call CloseSocket(UserIndex)
                        Exit Sub
                    Else
                        UserList(UserIndex).flags.CountSH = TempTick
                    End If
                End If
                UserList(UserIndex).flags.StartWalk = TempTick
                UserList(UserIndex).flags.TimesWalk = 0
            End If
            
            UserList(UserIndex).flags.TimesWalk = UserList(UserIndex).flags.TimesWalk + 1
            
            rdata = Right$(rdata, Len(rdata) - 1)
            
            If UserList(UserIndex).flags.Paralizado = 0 Then
                If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando Then
                    Call MoveUserChar(UserIndex, val(rdata))
                ElseIf UserList(UserIndex).flags.Descansar Then
                  UserList(UserIndex).flags.Descansar = False
                  Call SendData(ToIndex, UserIndex, 0, "DOK")
                  Call SendData(ToIndex, UserIndex, 0, "Y215")
                  Call MoveUserChar(UserIndex, val(rdata))
                ElseIf UserList(UserIndex).flags.Meditando Then
                  UserList(UserIndex).flags.Meditando = False
                  Call SendData(ToIndex, UserIndex, 0, "MEDOK")
                  Call SendData(ToIndex, UserIndex, 0, "Y216")
                  UserList(UserIndex).Char.FX = 0
                  UserList(UserIndex).Char.loops = 0
                  Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & 0 & "," & 0)
                  Call MoveUserChar(UserIndex, val(rdata))
                End If
            Else    'paralizado
              '[CDT 17-02-2004] (<- emmmmm ?????)
              If Not UserList(UserIndex).flags.UltimoMensaje = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "Y217")
                UserList(UserIndex).flags.UltimoMensaje = 1
              End If
              '[/CDT]
              UserList(UserIndex).flags.CountSH = 0
            End If
            
            If UserList(UserIndex).flags.Oculto = 1 Then
                
                If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then
                    Call SendData(ToIndex, UserIndex, 0, "Y23")
                    UserList(UserIndex).flags.Oculto = 0
                    UserList(UserIndex).flags.Invisible = 0
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                End If
                
            End If
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Empollando(UserIndex)
            Else
                UserList(UserIndex).flags.EstaEmpo = 0
                UserList(UserIndex).EmpoCont = 0
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(rdata)
        Case "RPU" 'Pedido de actualizacion de la posicion
            Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            Exit Sub
        Case "AT"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes atacar a nadie porque estas muerto!!. " & FONTTYPE_INFO)
                Exit Sub
            End If
            '[Consejeros]
'            If UserList(UserIndex).flags.Privilegios = 1 Then
'                Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a nadie. " & FONTTYPE_INFO)
'                Exit Sub
'            End If
            If Not UserList(UserIndex).flags.ModoCombate Then
                Call SendData(ToIndex, UserIndex, 0, "Y218")
            Else
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                                Call SendData(ToIndex, UserIndex, 0, "Y219")
                                Exit Sub
                    End If
                End If
                Call UsuarioAtaca(UserIndex)
                
                'piedra libre para todos los compas!
                If UserList(UserIndex).flags.Oculto > 0 Then
                    UserList(UserIndex).flags.Oculto = 0
                    UserList(UserIndex).flags.Invisible = 0
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                    Call SendData(ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
                End If
                
            End If
            Exit Sub
        Case "AG"
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden tomar objetos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y220")
                    Exit Sub
            End If
            Call GetObj(UserIndex)
            Exit Sub
        Case "TAB" 'Entrar o salir modo combate
            If UserList(UserIndex).flags.ModoCombate Then
                Call SendData(ToIndex, UserIndex, 0, "Y221")
            Else
                Call SendData(ToIndex, UserIndex, 0, "Y222")
            End If
            UserList(UserIndex).flags.ModoCombate = Not UserList(UserIndex).flags.ModoCombate
            Exit Sub
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(UserIndex).flags.Seguro Then
                  Call SendData(ToIndex, UserIndex, 0, "SEGOFF")
            Else
                  Call SendData(ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
        Case "ACTUALIZAR"
            Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            Exit Sub
        Case "GLINFO"
            If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                        Call SendGuildLeaderInfo(UserIndex)
            Else
                        Call SendGuildsList(UserIndex)
            End If
            Exit Sub
        Case "ATRI"
            Call EnviarAtrib(UserIndex)
            Exit Sub
        Case "FAMA"
            Call EnviarFama(UserIndex)
            Exit Sub
        Case "ESKI"
            Call EnviarSkills(UserIndex)
            Exit Sub
        Case "FEST" 'Mini estadisticas :)
            Call EnviarMiniEstadisticas(UserIndex)
            Exit Sub
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(ToIndex, UserIndex, 0, "FINCOMOK")
            Exit Sub
        Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(UserIndex).ComUsu.DestUsu > 0 And _
                UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
            
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(ToIndex, UserIndex, 0, "FINBANOK")
            Exit Sub
        '-------------------------------------------------------
        '[/KEVIN]**************************************
        Case "COMUSUOK"
            'Aceptar el cambio
            Call AceptarComercioUsu(UserIndex)
            Exit Sub
        Case "COMUSUNO"
            'Rechazar el cambio
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Call SendData(ToIndex, UserIndex, 0, "Y226")
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[/Alejo]
    
    
    End Select
    
    
    
    Select Case UCase$(Left$(rdata, 2))
    '    Case "/Z"
    '        Dim Pos As WorldPos, Pos2 As WorldPos
    '        Dim O As Obj
    '
    '        For LoopC = 1 To 100
    '            Pos = UserList(UserIndex).Pos
    '            O.Amount = 1
    '            O.ObjIndex = iORO
    '            'Exit For
    '            Call TirarOro(100000, UserIndex)
    '            'Call Tilelibre(Pos, Pos2)
    '            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
    '
    '            'Call MakeObj(ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
    '        Next LoopC
    '
    '        Exit Sub
        Case "TI" 'Tirar item
                If UserList(UserIndex).flags.Navegando = 1 Or _
                   UserList(UserIndex).flags.Muerto = 1 Or _
                   UserList(UserIndex).flags.Privilegios = 1 Then Exit Sub
                   '[Consejeros]
                
                rdata = Right$(rdata, Len(rdata) - 2)
                Arg1 = ReadField(1, rdata, 44)
                Arg2 = ReadField(2, rdata, 44)
                If val(Arg1) = FLAGORO Then
                    
                    Call TirarOro(val(Arg2), UserIndex)
                    
                    Call SendUserStatsBox(UserIndex)
                    Exit Sub
                Else
                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                    Else
                        Exit Sub
                    End If
                End If
                Exit Sub
        Case "LH" ' Lanzar hechizo
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "Y3")
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 2)
            UserList(UserIndex).flags.Hechizo = val(rdata)
            Exit Sub
        Case "LC" 'Click izquierdo
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Sub
        Case "RC" 'Click derecho
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Sub
        Case "UK"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "Y3")
                Exit Sub
            End If
    
            rdata = Right$(rdata, Len(rdata) - 2)
            Select Case val(rdata)
                Case Robar
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Robar)
                Case Magia
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Magia)
                Case Domar
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Domar)
                Case Ocultarse
                    
                    If UserList(UserIndex).flags.Navegando = 1 Then
                              '[CDT 17-02-2004]
                              If Not UserList(UserIndex).flags.UltimoMensaje = 3 Then
                                Call SendData(ToIndex, UserIndex, 0, "Y229")
                                UserList(UserIndex).flags.UltimoMensaje = 3
                              End If
                              '[/CDT]
                          Exit Sub
                    End If
                    
                    If UserList(UserIndex).flags.Oculto = 1 Then
                              '[CDT 17-02-2004]
                              If Not UserList(UserIndex).flags.UltimoMensaje = 2 Then
                                Call SendData(ToIndex, UserIndex, 0, "Y2")
                                UserList(UserIndex).flags.UltimoMensaje = 2
                              End If
                              '[/CDT]
                          Exit Sub
                    End If
                    
                    Call DoOcultarse(UserIndex)
            End Select
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rdata, 3))
         Case "UMH" ' Usa macro de hechizos
            Call SendData(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " fue expulsado por Anti-macro de hechizos " & FONTTYPE_VENENO)
            Call SendData(ToIndex, UserIndex, 0, "ERR Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros" & FONTTYPE_INFO)
            Call CloseSocket(UserIndex)
            Exit Sub
        Case "USA"
            rdata = Right$(rdata, Len(rdata) - 3)
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            Call UseInvItem(UserIndex, val(rdata))
            Exit Sub
        Case "CNS" ' Construye herreria
            rdata = Right$(rdata, Len(rdata) - 3)
            X = CInt(rdata)
            If X < 1 Then Exit Sub
            If ObjData(X).SkHerreria = 0 Then Exit Sub
            Call HerreroConstruirItem(UserIndex, X)
            Exit Sub
        Case "CNC" ' Construye carpinteria
            rdata = Right$(rdata, Len(rdata) - 3)
            X = CInt(rdata)
            If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
            Call CarpinteroConstruirItem(UserIndex, X)
            Exit Sub
        Case "WLC" 'Click izquierdo en modo trabajo
            rdata = Right$(rdata, Len(rdata) - 3)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            Arg3 = ReadField(3, rdata, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
            
            X = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(UserIndex).flags.Muerto = 1 Or _
               UserList(UserIndex).flags.Descansar Or _
               UserList(UserIndex).flags.Meditando Or _
               Not InMapBounds(UserList(UserIndex).Pos.Map, X, Y) Then Exit Sub
                              
            If Not InRangoVision(UserIndex, X, Y) Then
                Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                'Nos aseguramos que este usando un arma de proyectiles
                DummyInt = 0
                
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.WeaponEqpSlot < 1 Or UserList(UserIndex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.MunicionEqpSlot < 1 Or UserList(UserIndex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                ElseIf ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).ObjType <> OBJTYPE_FLECHAS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).Amount < 1 Then
                    DummyInt = 1
                End If
                
                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call SendData(ToIndex, UserIndex, 0, "Y230")
                    End If
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                    Exit Sub
                End If
                
                DummyInt = 0
                
                'Quitamos stamina
                If UserList(UserIndex).Stats.MinSta >= 10 Then
                     Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                     Call SendData(ToIndex, UserIndex, 0, "Y11")
                     Exit Sub
                End If
                 
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
                
                TU = UserList(UserIndex).flags.TargetUser
                tN = UserList(UserIndex).flags.TargetNpc
                                
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = UserIndex Then
                        Call SendData(ToIndex, UserIndex, 0, "Y231")
                        DummyInt = 1
                        Exit Sub
                    End If
                End If
    
                If DummyInt = 0 Then
                    'Saca 1 flecha
                    DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
                    If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                        UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, UserIndex, DummyInt)
                        UserList(UserIndex).Invent.MunicionEqpSlot = 0
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
                    End If
                    '-----------------------------------
                End If
                
                If tN > 0 Then
                    If Npclist(tN).Attackable <> 0 Then
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                ElseIf TU > 0 Then
                    If UserList(UserIndex).flags.Seguro Then
                        If Not Criminal(TU) Then
                            Call SendData(ToIndex, UserIndex, 0, "Y232")
                            Exit Sub
                        End If
                    End If
                    Call UsuarioAtacaUsuario(UserIndex, TU)
                End If
                
            Case Magia
'                If UserList(UserIndex).flags.PuedeLanzarSpell = 0 Then Exit Sub
                '[Consejeros]
'                If UserList(UserIndex).flags.Privilegios = 1 Then Exit Sub
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(UserIndex).Pos.Map
                wp2.X = X
                wp2.Y = Y
                                
                If UserList(UserIndex).flags.Hechizo > 0 Then
                    If IntervaloPermiteLanzarSpell(UserIndex) Then
                        Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                    '    UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(UserIndex).flags.Hechizo = 0
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "Y233")
                End If
                
                'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
                If (Abs(UserList(UserIndex).Pos.X - wp2.X) > 9 Or Abs(UserList(UserIndex).Pos.Y - wp2.Y) > 8) Then
                    Dim txt As String
                    txt = "Ataque fuera de rango de " & UserList(UserIndex).Name & "(" & UserList(UserIndex).Pos.Map & "/" & UserList(UserIndex).Pos.X & "/" & UserList(UserIndex).Pos.Y & ") ip: " & UserList(UserIndex).ip & " a la posicion (" & wp2.Map & "/" & wp2.X & "/" & wp2.Y & ") "
                    If UserList(UserIndex).flags.Hechizo > 0 Then
                        txt = txt & ". Hechizo: " & Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).Nombre
                    End If
                    If MapData(wp2.Map, wp2.X, wp2.Y).UserIndex > 0 Then
                        txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.X, wp2.Y).UserIndex).Name
                    ElseIf MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex > 0 Then
                        txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex).Name
                    End If
                    
                    Call LogCheating(txt)
                End If
                
            
            
            
            Case Pesca
                        
                AuxInd = UserList(UserIndex).Invent.HerramientaEqpObjIndex
                If AuxInd = 0 Then Exit Sub
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If AuxInd <> OBJTYPE_CAÑA And AuxInd <> RED_PESCA Then
                        Call Cerrar_Usuario(UserIndex)
                        Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y234")
                    Exit Sub
                End If
                
                If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
                    
                    Select Case AuxInd
                    Case OBJTYPE_CAÑA
                        Call DoPescar(UserIndex)
                    Case RED_PESCA
                        With UserList(UserIndex)
                            wpaux.Map = .Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                        End With
                        
                        If Distancia(UserList(UserIndex).Pos, wpaux) > 2 Then
                            Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos para pescar." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoPescarRed(UserIndex)
                    End Select
    
                Else
                    Call SendData(ToIndex, UserIndex, 0, "Y235")
                End If
                
            Case Robar
               If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                    'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                       If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = val(ReadField(1, rdata, 44))
                            wpaux.Y = val(ReadField(2, rdata, 44))
                            If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                                Call SendData(ToIndex, UserIndex, 0, "Y5")
                                Exit Sub
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(UserIndex).flags.TargetUser).Pos.Map, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y).trigger = TRIGGER_ZONASEGURA Then
                                Call SendData(ToIndex, UserIndex, 0, "Y236")
                                Exit Sub
                            End If
                            If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = TRIGGER_ZONASEGURA Then
                                Call SendData(ToIndex, UserIndex, 0, "Y236")
                                Exit Sub
                            End If
                            
                            Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "Y237")
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "Y238")
                End If
            Case Talar
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y239")
                    Exit Sub
                End If
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                        Call CloseSocket(UserIndex)
                        Exit Sub
                End If
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(ToIndex, UserIndex, 0, "Y5")
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(UserIndex).Pos) = 0 Then
                        Call SendData(ToIndex, UserIndex, 0, "Y240")
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).ObjType = OBJTYPE_ARBOLES Then
                        Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                        Call DoTalar(UserIndex)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "241")
                End If
            Case Mineria
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                        Call CloseSocketSL(UserIndex)
                        Call Cerrar_Usuario(UserIndex)
                        Exit Sub
                End If
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(ToIndex, UserIndex, 0, "Y5")
                        Exit Sub
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(AuxInd).ObjType = OBJTYPE_YACIMIENTO Then
                        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
                        Call DoMineria(UserIndex)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "Y242")
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "Y242")
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
              CI = UserList(UserIndex).flags.TargetNpc
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 2 Then
                                  Call SendData(ToIndex, UserIndex, 0, "Y5")
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(ToIndex, UserIndex, 0, "Y243")
                                  Exit Sub
                            End If
                            Call DoDomar(UserIndex, CI)
                        Else
                            Call SendData(ToIndex, UserIndex, 0, "Y244")
                        End If
              Else
                     Call SendData(ToIndex, UserIndex, 0, "Y245")
              End If
              
            Case FundirMetal
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex <> UserList(UserIndex).flags.TargetObjInvIndex Then
                            If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call SendData(ToIndex, UserIndex, 0, "Y246")
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                            'Call LogCheating(UserList(UserIndex).Name & " intento crear minerales a partir de otros: FlagSlot/usaba/usoconclick/cantidad/IP:" & UserList(UserIndex).flags.TargetObjInvSlot & "/" & UserList(UserIndex).flags.TargetObjInvIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount & "/" & UserList(UserIndex).ip)
                            'UserList(UserIndex).flags.Ban = 1
                            'Call SendData(ToAll, 0, 0, "||>>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                            Call SendData(ToIndex, UserIndex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        Call FundirMineral(UserIndex)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "Y247")
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "Y247")
                End If
                
            Case Herreria
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call SendData(ToIndex, UserIndex, 0, "SFH")
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "Y248")
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "Y248")
                End If
                
            End Select
            
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
        Case "CIG"
            rdata = Right$(rdata, Len(rdata) - 3)
            X = Guilds.Count
            
            If CreateGuild(UserList(UserIndex).Name, UserList(UserIndex).Reputacion.Promedio, UserIndex, rdata) Then
                If X = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y249")
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el clan numero " & X + 1 & " de Argentum!!!." & FONTTYPE_INFO)
                End If
                Call SaveGuildsDB
            End If
            
            Exit Sub
    End Select
    
    
    
    
    
    Select Case UCase$(Left$(rdata, 4))
        Case "INFS" 'Informacion del hechizo
                rdata = Right$(rdata, Len(rdata) - 4)
                If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
                    Dim H As Integer
                    H = UserList(UserIndex).Stats.UserHechizos(val(rdata))
                    If H > 0 And H < NumeroHechizos + 1 Then
                        Call SendData(ToIndex, UserIndex, 0, "Y250")
                        Call SendData(ToIndex, UserIndex, 0, "||Nombre:" & Hechizos(H).Nombre & FONTTYPE_INFO)
                        Call SendData(ToIndex, UserIndex, 0, "||Descripcion:" & Hechizos(H).Desc & FONTTYPE_INFO)
                        Call SendData(ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(H).MinSkill & " de magia." & FONTTYPE_INFO)
                        Call SendData(ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(H).ManaRequerido & FONTTYPE_INFO)
                        Call SendData(ToIndex, UserIndex, 0, "||Stamina necesaria: " & Hechizos(H).StaRequerido & FONTTYPE_INFO)
                        Call SendData(ToIndex, UserIndex, 0, "Y251")
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "Y252")
                End If
                Exit Sub
        'el usuario reporta un md5
        Case "RMDC"
            UserList(UserIndex).flags.MD5Reportado = UCase$(Right$(rdata, Len(rdata) - 4))
            Exit Sub
        Case "EQUI"
                If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "Y26")
                Exit Sub
                End If
                rdata = Right$(rdata, Len(rdata) - 4)
                If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                     If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
                Else
                    Exit Sub
                End If
                Call EquiparInvItem(UserIndex, val(rdata))
                Exit Sub
        Case "CHEA" 'Cambiar Heading ;-)
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) > 0 And val(rdata) < 5 Then
                UserList(UserIndex).Char.Heading = rdata
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            Exit Sub
        Case "SKSE" 'Modificar skills
            Dim sumatoria As Integer
            Dim incremento As Integer
            rdata = Right$(rdata, Len(rdata) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rdata, 44))
                
                If incremento < 0 Then
                    'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                    Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                    UserList(UserIndex).Stats.SkillPts = 0
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                sumatoria = sumatoria + incremento
            Next i
            
            If sumatoria > UserList(UserIndex).Stats.SkillPts Then
                'UserList(UserIndex).Flags.AdministrativeBan = 1
                'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rdata, 44))
                UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
                UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
                If UserList(UserIndex).Stats.UserSkills(i) > 100 Then UserList(UserIndex).Stats.UserSkills(i) = 100
            Next i
            Exit Sub
        Case "ENTR" 'Entrena hombre!
            
            If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
            
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 3 Then Exit Sub
            
            rdata = Right$(rdata, Len(rdata) - 4)
            
            If Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rdata) > 0 And val(rdata) < Npclist(UserList(UserIndex).flags.TargetNpc).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNpc).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNpc).Pos, True, False)
                        If SpawnedNpc <= MAXNPCS Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNpc
                            Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas = Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas + 1
                        End If
                End If
            Else
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
            End If
            
            Exit Sub
        Case "COMP"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNpc > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                       Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rdata = Right$(rdata, Len(rdata) - 5)
             'User compra el item del slot rdata
             If UserList(UserIndex).flags.Comerciando = False Then
                Call SendData(ToIndex, UserIndex, 0, "Y253")
                Exit Sub
             End If
             'listindex+1, cantidad
             Call NPCVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).flags.TargetNpc)
             Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNpc > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rdata = Right(rdata, Len(rdata) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
             Exit Sub
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNpc > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                       Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.charindex))
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rdata = Right$(rdata, Len(rdata) - 5)
             'User compra el item del slot rdata
             Call NPCCompraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
             Exit Sub
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "DEPO"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNpc > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rdata = Right(rdata, Len(rdata) - 5)
             'User deposita el item del slot rdata
             Call UserDepositaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
             Exit Sub
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
             
    End Select
    
    Select Case UCase$(Left$(rdata, 5))
        Case "DEMSG"
            If UserList(UserIndex).flags.TargetObj > 0 Then
            rdata = Right$(rdata, Len(rdata) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rdata, 176)
            msg = ReadField(2, rdata, 176)
            Dim n2 As Integer, loopme As Integer
            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = val(GetVar(f, "INFO", "CantMSG"))
                If num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                    num = 0
                End If
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rdata, 6))
        Case "DESPHE" 'Mover Hechizo de lugar
            rdata = Right(rdata, Len(rdata) - 6)
            Call DesplazarHechizo(UserIndex, CInt(ReadField(1, rdata, 44)), CInt(ReadField(2, rdata, 44)))
            Exit Sub
        Case "DESCOD" 'Informacion del hechizo
                rdata = Right$(rdata, Len(rdata) - 6)
                Call UpdateCodexAndDesc(rdata, UserIndex)
                Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(rdata, 7))
    Case "OFRECER"
            rdata = Right$(rdata, Len(rdata) - 7)
            Arg1 = ReadField(1, rdata, Asc(","))
            Arg2 = ReadField(2, rdata, Asc(","))

            If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
                Exit Sub
            End If
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
                'sigue vivo el usuario ?
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else
                'esta vivo ?
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Sub
                End If
                '//Tiene la cantidad que ofrece ??//'
                If val(Arg1) = FLAGORO Then
                    'oro
                    If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                        Call SendData(ToIndex, UserIndex, 0, "Y210")
                        Exit Sub
                    End If
                Else
                    'inventario
                    If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                        Call SendData(ToIndex, UserIndex, 0, "Y210")
                        Exit Sub
                    End If
                End If
                '[Consejeros]
                If UserList(UserIndex).ComUsu.Objeto > 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y254")
                    Exit Sub
                End If
                UserList(UserIndex).ComUsu.Objeto = val(Arg1)
                UserList(UserIndex).ComUsu.Cant = val(Arg2)
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Sub
                Else
                    '[CORREGIDO]
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha cambiado su oferta." & FONTTYPE_TALK)
                    End If
                    '[/CORREGIDO]
                    'Es la ofrenda de respuesta :)
                    Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Exit Sub
    End Select
    '[/Alejo]
    
    Select Case UCase$(Left$(rdata, 8))
        Case "ACEPPEAT"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call AcceptPeaceOffer(UserIndex, rdata)
            Exit Sub
        Case "PEACEOFF"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call RecievePeaceOffer(UserIndex, rdata)
            Exit Sub
        Case "PEACEDET"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SendPeaceRequest(UserIndex, rdata)
            Exit Sub
        Case "ENVCOMEN"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SendPeticion(UserIndex, rdata)
            Exit Sub
        Case "ENVPROPP"
            Call SendPeacePropositions(UserIndex)
            Exit Sub
        Case "DECGUERR"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call DeclareWar(UserIndex, rdata)
            Exit Sub
        Case "DECALIAD"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call DeclareAllie(UserIndex, rdata)
            Exit Sub
        Case "NEWWEBSI"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SetNewURL(UserIndex, rdata)
            Exit Sub
        Case "ACEPTARI"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call AcceptClanMember(UserIndex, rdata)
            Exit Sub
        Case "RECHAZAR"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call DenyRequest(UserIndex, rdata)
            Exit Sub
        Case "ECHARCLA"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call EacharMember(UserIndex, rdata)
            Exit Sub
        Case "ACTGNEWS"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call UpdateGuildNews(rdata, UserIndex)
            Exit Sub
        Case "1HRINFO<"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call SendCharInfo(rdata, UserIndex)
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rdata, 9))
        Case "SOLICITUD"
             rdata = Right$(rdata, Len(rdata) - 9)
             Call SolicitudIngresoClan(UserIndex, rdata)
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 11))
      Case "CLANDETAILS"
            rdata = Right$(rdata, Len(rdata) - 11)
            Call SendGuildDetails(UserIndex, rdata)
            Exit Sub
    End Select
    
Procesado = False
    
End Sub
