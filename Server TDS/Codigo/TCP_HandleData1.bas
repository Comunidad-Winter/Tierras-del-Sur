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

'********************Misery_Ezequiel 28/05/05********************'
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
        
        rdata = Right$(rdata, Len(rdata) - 1)
        If Mid(rdata, 1, 1) = "." And UserList(UserIndex).Stats.GlobAl = 2 Then
        If charlageneral = False Then
        Call Senddata(ToIndex, UserIndex, 0, "||El Global esta desactivado." & "~190~190~190~0~1~")
        Else
            If UserList(UserIndex).Stats.ELV < 10 Then
            Call Senddata(ToIndex, UserIndex, 0, "||Debes ser nivel 10 o superior." & "~190~190~190~0~1~")
            Else
                If UserList(UserIndex).Stats.CALLADO = True Then
                Call Senddata(ToIndex, UserIndex, 0, "||Se te ha prohibido hablar." & "~190~190~190~0~1~")
                Else
                For i = 1 To LastUser
                     If UserList(i).Stats.GlobAl = 2 Then
                     Call Senddata(ToIndex, i, 0, "||" & UserList(UserIndex).Name & "> " & Replace(Right$(rdata, Len(rdata) - 1), "~", "?") & "~190~190~190~0~1~")
                     End If
                Next
                End If
            End If
        End If
        Else
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
            rdata = Replace(rdata, " pt ", " capo ")
            rdata = Replace(rdata, " nw ", " pro ")
            rdata = Replace(rdata, " PT ", " CAPO ")
            rdata = Replace(rdata, " NW ", " PRO ")
            rdata = Replace(rdata, " Pt ", " Capo ")
            rdata = Replace(rdata, " Nw ", " Pro ")
            rdata = Replace(rdata, " pT ", " CaPo ")
            rdata = Replace(rdata, " nW ", " PrO ")
            rdata = Mid(rdata, 2, Len(rdata) - 2)
            '[/CDT]
            'piedra libre para todos los compas!
            
          
            
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                Call Senddata(ToIndex, UserIndex, 0, "Y288")
            End If
            
             If UserList(UserIndex).flags.Muerto = 1 Then
                  Call Senddata(ToDeadArea, UserIndex, UserList(UserIndex).Pos.Map, "||12632256°" & rdata & "°" & str(ind))
               Call Senddata(ToAdminsAreaButConsejeros, UserIndex, UserList(UserIndex).Pos.Map, "||12632256°" & rdata & "°" & str(ind))
           Else
              If Len(rdata) = 1 And rdata = " " Then
            Call Senddata(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & rdata & "°" & str(ind))
            Else
            Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & rdata & "°" & str(ind))
           
            End If
            Exit Sub 'ALto Buggggggggggg [WizARd]
            End If
       End If
       Exit Sub
        Case "-" 'Gritar
        
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y289")
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
            
                If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                Call Senddata(ToIndex, UserIndex, 0, "Y288")
            End If
            ind = UserList(UserIndex).Char.charindex
            If Len(rdata) = 0 Then
            Call Senddata(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "°" & rdata & "°" & str(ind))
            Exit Sub
            Else
            Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "°" & rdata & "°" & str(ind))
            Exit Sub
            End If
        Case "\" 'Susurrar al oido
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y289")
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 1)
            tName = ReadField(1, rdata, 32)
            tIndex = NameIndex(tName)
            If tIndex <> 0 Then
                If UserList(tIndex).flags.Privilegios > 0 And UserList(UserIndex).flags.Privilegios = 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y213")
                    Exit Sub
                End If
                If Len(rdata) <> Len(tName) Then
                    tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                If Not EstaPCarea(UserIndex, tIndex) Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y290")
                    Exit Sub
                End If
                ind = UserList(UserIndex).Char.charindex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                '[Consejeros]
                If UserList(UserIndex).flags.Privilegios > 1 Then
                    Call LogGM(UserList(UserIndex).Name, "Le dijo a '" & UserList(tIndex).Name & "' " & tMessage, True)
                End If
    
                Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                Call Senddata(ToIndex, tIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                '[CDT 17-02-2004]
                If UserList(UserIndex).flags.Privilegios < 2 Then
                    Call Senddata(ToAdminsAreaButConsejeros, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°" & "a " & UserList(tIndex).Name & "> " & tMessage & "°" & str(ind))
                End If
                '[/CDT]
                Exit Sub
            Else
            If UserDarPrivilegioLevel(tName) > 0 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y213")
            Else
            Call Senddata(ToIndex, UserIndex, 0, "Y214")
            End If
            End If
            Exit Sub
            
        Case "+"
        rdata = Right$(rdata, Len(rdata) - 1)
          If UserList(UserIndex).flags.ModoRol = True Then
           If rdata <> "" Then
              If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                 For i = 1 To LastUser
                 Call Senddata(ToIndex, i, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~0~150~255~1~0~" & ENDC)
                 Next
                 Exit Sub
              Else
                 For i = 1 To LastUser
                 Call Senddata(ToIndex, i, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~255~0~0~1~0~" & ENDC)
                 Next
                 Exit Sub
              End If
           End If
        End If
        
      
        
        Case "M" 'Moverse
           rdata = Right$(rdata, Len(rdata) - 1)
          ' Marche [4-20-04}
            If UserList(UserIndex).Counters.Saliendo Then
                Call Senddata(ToIndex, UserIndex, 0, "Y291")
                UserList(UserIndex).Counters.Saliendo = False
                UserList(UserIndex).Counters.Salir = 0
            End If
           'End Marche
            If UserList(UserIndex).flags.Paralizado = 0 Then
                If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando Then
                    Call MoveUserChar(UserIndex, val(rdata))
                ElseIf UserList(UserIndex).flags.Descansar Then
                  UserList(UserIndex).flags.Descansar = False
                  Call Senddata(ToIndex, UserIndex, 0, "DOK")
                  Call Senddata(ToIndex, UserIndex, 0, "Y215")
                  Call MoveUserChar(UserIndex, val(rdata))
                ElseIf UserList(UserIndex).flags.Meditando Then
                  UserList(UserIndex).flags.Meditando = False
                  Call Senddata(ToIndex, UserIndex, 0, "MEDOK")
                  Call Senddata(ToIndex, UserIndex, 0, "Y216")
                  UserList(UserIndex).Char.FX = 0
                  UserList(UserIndex).Char.loops = 0
                  Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & 0 & "," & 0)
                  Call MoveUserChar(UserIndex, val(rdata))
                End If
            Else    'paralizado
                UserList(UserIndex).flags.CountSH = 0
            End If
            If UserList(UserIndex).flags.Oculto = 1 Then
                If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y23")
                    UserList(UserIndex).flags.Oculto = 0
                    UserList(UserIndex).flags.Invisible = 0
                    Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                End If
            End If
        
    End Select
    
    Select Case UCase$(rdata)
        Case "RPU" 'Pedido de actualizacion de la posicion
            Call Senddata(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            Exit Sub
        Case "AT"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y292")
                Exit Sub
            End If

            If Not UserList(UserIndex).flags.ModoCombate Then
                Call Senddata(ToIndex, UserIndex, 0, "Y218")
            Else
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                            Call Senddata(ToIndex, UserIndex, 0, "Y219")
                        Exit Sub
                    End If
                End If
                
                   'Anti sh
                   
                    Call UsuarioAtaca(UserIndex)
                '[Misery_Ezequiel 29/06/05]
                If UCase$(UserList(UserIndex).Clase) = "CAZADOR" And UserList(UserIndex).flags.Oculto > 0 And UserList(UserIndex).Stats.UserSkills(Ocultarse) > 90 Then
                    If UserList(UserIndex).Invent.ArmourEqpObjIndex = 648 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
                        Exit Sub
                    End If
                End If
                '[\]Misery_Ezequiel 29/06/05]
                
                'piedra libre para todos los compas!
                If UserList(UserIndex).flags.Oculto > 0 Then
                    UserList(UserIndex).flags.Oculto = 0
                    UserList(UserIndex).flags.Invisible = 0
                    Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
                    Call Senddata(ToIndex, UserIndex, 0, "Y288")
                End If
                
            End If
            Exit Sub
        Case "AG"
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y293")
                    Exit Sub
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y220")
                    Exit Sub
            End If
            Call GetObj(UserIndex)
            Exit Sub
        Case "TAB" 'Entrar o salir modo combate
            If UserList(UserIndex).flags.ModoCombate Then
                Call Senddata(ToIndex, UserIndex, 0, "Y221")
            Else
                Call Senddata(ToIndex, UserIndex, 0, "Y222")
            End If
            UserList(UserIndex).flags.ModoCombate = Not UserList(UserIndex).flags.ModoCombate
            Exit Sub
        '[Misery_Ezequiel 12/06/05]
    
        
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(UserIndex).flags.Seguro Then
                Call Senddata(ToIndex, UserIndex, 0, "Y345")
            Else
                Call Senddata(ToIndex, UserIndex, 0, "SEGON")
                UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            End If
            Exit Sub
        '[\]Misery_Ezequiel 12/06/05]
        Case "ACTUALIZAR"
            Call Senddata(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
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
            Call Senddata(ToIndex, UserIndex, 0, "FINCOMOK")
            Exit Sub
        Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(UserIndex).ComUsu.DestUsu > 0 And _
                UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call Senddata(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(UserIndex).flags.Comerciando = False
            Call Senddata(ToIndex, UserIndex, 0, "FINBANOK")
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
                    Call Senddata(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Call Senddata(ToIndex, UserIndex, 0, "Y226")
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[/Alejo]
        
        
        

        
        Case "ENCARCEL"
        Call Encarcelar(UserIndex, TIEMPO_CARCEL_PIQUETE)
         
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
                Call Senddata(ToIndex, UserIndex, 0, "Y3")
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 2)
            Select Case val(rdata)
                Case Robar
                    Call Senddata(ToIndex, UserIndex, 0, "T01" & Robar)
              ' Pasado al cliente 'Case Magia
                 '   Call Senddata(ToIndex, UserIndex, 0, "T01" & Magia)
                Case Domar
                    Call Senddata(ToIndex, UserIndex, 0, "T01" & Domar)
                Case Ocultarse
                    If UserList(UserIndex).flags.Navegando = 1 Then
                              '[CDT 17-02-2004]
                              If Not UserList(UserIndex).flags.UltimoMensaje = 3 Then
                                Call Senddata(ToIndex, UserIndex, 0, "Y229")
                                UserList(UserIndex).flags.UltimoMensaje = 3
                              End If
                              '[/CDT]
                          Exit Sub
                    End If
                    If UserList(UserIndex).flags.Oculto = 1 Then
                              '[CDT 17-02-2004]
                              'If Not UserList(UserIndex).flags.UltimoMensaje = 2 Then
                                Call Senddata(ToIndex, UserIndex, 0, "Y2")
                               ' UserList(UserIndex).flags.UltimoMensaje = 2
                              'End If
                              '[/CDT]
                          Exit Sub
                    End If
                    Call DoOcultarse(UserIndex)
            End Select
            Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 3))
        Case "TCX"
        Dim hastacuanto As Integer
        rdata = Right$(rdata, Len(rdata) - 3)
        UserList(UserIndex).flags.TimesWalk = rdata
        If UserList(UserIndex).Stats.UserSkills(Mineria) <= 25 Then
        hastacuanto = 1
            ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 50 Then
            hastacuanto = 2
            ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 75 Then
            hastacuanto = 3
            Else
            hastacuanto = 4
        End If
        
        Call FundirMineral(UserIndex, hastacuanto)
        
         Case "UMH" ' Usa macro de hechizos
            Call Senddata(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " fue expulsado por Anti-macro de hechizos " & FONTTYPE_VENENO)
            Call Senddata(ToIndex, UserIndex, 0, "ERR Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
            Call CloseSocket(UserIndex)
            Exit Sub
            
         Case "PPP"
             rdata = Right$(rdata, Len(rdata) - 3)
             Dim Obj As ObjData
            Dim ObjIndex As Integer
            Dim TargObj As ObjData
            Dim MiObj As Obj
            Dim Slot As Long
            Slot = rdata
            If UserList(UserIndex).Invent.Object(Slot).Amount = 0 Then Exit Sub
            Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)
            If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call Senddata(ToIndex, UserIndex, 0, "Y287")
            Exit Sub
            End If
            '[Wizard 02/09/05] -> Arregla el bug de que sino tiene stamina labura = usando el f8.
                If Not UserList(UserIndex).Stats.MinSta > 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y167")
                    Exit Sub
                End If
            '[/Wizard]


            
            
            If Obj.ObjType = 24 Then
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                Else
                Call Senddata(ToIndex, UserIndex, 0, "Y350")
                Exit Sub
                End If
            End If
            If Not IntervaloPermiteUsar(UserIndex) Then
            Exit Sub
            End If
            ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
            UserList(UserIndex).flags.TargetObjInvSlot = Slot

            Select Case ObjIndex
                Case OBJTYPE_CAÑA, RED_PESCA
                    Call Senddata(ToIndex, UserIndex, 0, "PPP" & Pesca)
                Case HACHA_LEÑADOR
                    Call Senddata(ToIndex, UserIndex, 0, "PPP" & Talar)
                '[Misery_Ezequiel 27/05/05]
                Case HACHA_DORADA
                    Call Senddata(ToIndex, UserIndex, 0, "PPP" & Talar)
                '[\]Misery_Ezequiel 27/05/05]
                Case PIQUETE_MINERO
                    Call Senddata(ToIndex, UserIndex, 0, "PPP" & Mineria)
                Case MARTILLO_HERRERO
                    Call Senddata(ToIndex, UserIndex, 0, "PPP" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(UserIndex)
                    Call Senddata(ToIndex, UserIndex, 0, "SFC")
                'Case 192, 193, 194
                 '   If UserList(UserIndex).flags.Muerto = 1 Then
                  '  Call Senddata(ToIndex, UserIndex, 0, "Y26")
                   ' Exit Sub
                    'End If
                'Call Senddata(ToIndex, UserIndex, 0, "PPP" & FundirMetal)
           
                End Select
        Case "USA"
            'Marche
            If UserList(UserIndex).flags.Meditando Then
            Exit Sub
            End If
            'marche
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
                Call Senddata(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            Case Proyectiles
            If Not IntervaloPermiteAtacar(UserIndex, True) Then Exit Sub
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
                        Call Senddata(ToIndex, UserIndex, 0, "Y230")
                    End If
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                    Exit Sub
                End If
                DummyInt = 0
                'Quitamos stamina
                If UserList(UserIndex).Stats.MinSta >= ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia Then
                     Call QuitarSta(UserIndex, RandomNumber(1, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia))
                Else
                     Call Senddata(ToIndex, UserIndex, 0, "Y11")
                     Exit Sub
                End If
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
                TU = UserList(UserIndex).flags.TargetUser
                tN = UserList(UserIndex).flags.TargetNPC
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = UserIndex Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y231")
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
                            Call Senddata(ToIndex, UserIndex, 0, "Y232")
                            Exit Sub
                        End If
                    End If
                    Call UsuarioAtacaUsuario(UserIndex, TU)
                End If
            Case Magia
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(UserIndex).Pos.Map
                wp2.X = X
                wp2.Y = Y
                If UserList(UserIndex).flags.Hechizo > 0 Then
                  
                        Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                       'UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(UserIndex).flags.Hechizo = 0
                    'End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y233")
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
           
                
                If AuxInd <> OBJTYPE_CAÑA And AuxInd <> RED_PESCA Then
                        Call Cerrar_Usuario(UserIndex)
                        Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y234")
                    Exit Sub
                End If
                If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                    Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
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
                            Call Senddata(ToIndex, UserIndex, 0, "Y294")
                            Exit Sub
                        End If
                        Call DoPescarRed(UserIndex)
                    End Select
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y235")
                End If
                
            Case Robar
               If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                    'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                    If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                       If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = val(ReadField(1, rdata, 44))
                            wpaux.Y = val(ReadField(2, rdata, 44))
                            
                
                If UserList(UserIndex).flags.Oculto = 0 Then Exit Sub
                
                   If UCase(UserList(UserIndex).Clase) = "LADRON" Then
                            If UserList(UserIndex).Stats.ELV >= 25 Then
                                If Distancia(wpaux, UserList(UserIndex).Pos) > 5 Then
                                Call Senddata(ToIndex, UserIndex, 0, "Y5")
                                Exit Sub
                                End If
                            Else
                                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                                Call Senddata(ToIndex, UserIndex, 0, "Y5")
                                Exit Sub
                                End If
                            End If
                 Else
                        If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y5")
                        Exit Sub
                        End If
                End If
                                
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(UserIndex).flags.TargetUser).Pos.Map, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y).trigger = TRIGGER_ZONASEGURA Then
                                Call Senddata(ToIndex, UserIndex, 0, "Y236")
                                Exit Sub
                            End If
                            If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = TRIGGER_ZONASEGURA Then
                                Call Senddata(ToIndex, UserIndex, 0, "Y236")
                                Exit Sub
                            End If
                            Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                       End If
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y237")
                    End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y238")
                End If
                
            Case Talar
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                   
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y239")
                    Exit Sub
                End If
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    'Call CloseSocket(UserIndex)
                    'Exit Sub
                End If
                '[Misery_Ezequiel 28/05/05]
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_DORADA Then
                    'Call CloseSocket(UserIndex)
                    'Exit Sub
                End If
                '[\]Misery_Ezequiel 28/05/05]
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y5")
                        Exit Sub
                    End If
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(UserIndex).Pos) = 0 Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y240")
                        Exit Sub
                    End If
                    '[Misery_Ezequiel 28/05/05]
                    'Si no talas en un arbol de tejo
                    If AuxInd = 634 And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_DORADA Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y278")
                        Exit Sub
                    End If
                    If AuxInd <> 634 And UserList(UserIndex).Invent.HerramientaEqpObjIndex = HACHA_DORADA Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y279")
                        Exit Sub
                    End If
                    If AuxInd = 634 And UserList(UserIndex).Invent.HerramientaEqpObjIndex = HACHA_DORADA Then
                    If ObjData(AuxInd).ObjType = OBJTYPE_ARBOLES Then
                        Call Senddata(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                        Call DoTalar(UserIndex, 634)
                        Exit Sub
                    End If
                    End If
                    If AuxInd <> 0 And UserList(UserIndex).Invent.HerramientaEqpObjIndex = HACHA_LEÑADOR Or UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_DORADA Then
                    If ObjData(AuxInd).ObjType = OBJTYPE_ARBOLES Then
                        Call Senddata(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                        Call DoTalar(UserIndex, AuxInd)
                    End If
                    '[\]Misery_Ezequiel 27/05/05]
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "241")
                End If
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
                        Call Senddata(ToIndex, UserIndex, 0, "Y5")
                        Exit Sub
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(AuxInd).ObjType = OBJTYPE_YACIMIENTO Then
                        Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
                        Call DoMineria(UserIndex)
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y242")
                    End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y242")
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
              CI = UserList(UserIndex).flags.TargetNPC
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 2 Then
                                  Call Senddata(ToIndex, UserIndex, 0, "Y5")
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call Senddata(ToIndex, UserIndex, 0, "Y243")
                                  Exit Sub
                            End If
                            Call DoDomar(UserIndex, CI)
                        Else
                            Call Senddata(ToIndex, UserIndex, 0, "Y244")
                        Exit Sub
                        'Eze como vas a sacar eso!! Podes domar a cualquier criatura
                            '[Misery_Ezequiel 26/06/05]
                           ' Call DoDomar(UserIndex, CI)
                            '[\]Misery_Ezequiel 26/06/05]
                        End If
              Else
                     Call Senddata(ToIndex, UserIndex, 0, "Y245")
              End If

            Case FundirMetal
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex <> UserList(UserIndex).flags.TargetObjInvIndex Then
                            If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call Senddata(ToIndex, UserIndex, 0, "Y246")
                                Exit Sub
                            End If
                            ''FUISTE
                            'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                            'Call LogCheating(UserList(UserIndex).Name & " intento crear minerales a partir de otros: FlagSlot/usaba/usoconclick/cantidad/IP:" & UserList(UserIndex).flags.TargetObjInvSlot & "/" & UserList(UserIndex).flags.TargetObjInvIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount & "/" & UserList(UserIndex).ip)
                            'UserList(UserIndex).flags.Ban = 1
                            'Call SendData(ToAll, 0, 0, "||>>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                            Call Senddata(ToIndex, UserIndex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        Call Senddata(ToIndex, UserIndex, 0, "TT47")
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y247")
                    End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y247")
                End If
                
            Case Herreria
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call Senddata(ToIndex, UserIndex, 0, "SFH")
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y248")
                    End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y248")
                End If
            End Select
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
        Case "DEJ"
        UserList(UserIndex).TyTrabajo = 0
        UserList(UserIndex).TyTrabajoMod = 0
        UserList(UserIndex).Suerte = 0
        UserList(UserIndex).flags.Trabajando = False
        Case "TRA"
         On Error GoTo minaria
         If UserList(UserIndex).flags.Trabajando = True Then Call DejarDeTrabajar(UserIndex)
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
            If Not UserList(UserIndex).Stats.MinSta > 0 Then
            Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y167")
            Exit Sub
            End If
        
            If ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y287")
            Exit Sub
            End If
            Select Case UserList(UserIndex).Invent.HerramientaEqpObjIndex
            '//////////////////////////////////////////////
             Case OBJTYPE_CAÑA, RED_PESCA
                If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 Then
                    Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y234")
                    Exit Sub
                End If
                
                If HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
                    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = OBJTYPE_CAÑA Then
                    UserList(UserIndex).TyTrabajoMod = OBJTYPE_CAÑA
                    Else
                    UserList(UserIndex).TyTrabajoMod = RED_PESCA
                    End If
                    Call DameSuerte(UserIndex, 1)
                    UserList(UserIndex).TyTrabajo = 1
                    UserList(UserIndex).flags.Trabajando = True
                    Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "EMPT")
                Else
                    Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y235")
                    Exit Sub
                End If
                Dim auxinda As Integer
                Dim wpauxa As WorldPos
            '//////////////////////////////////////////////
            Case HACHA_LEÑADOR, HACHA_DORADA
                auxinda = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                If auxinda > 0 Then
                    wpauxa.Map = UserList(UserIndex).Pos.Map
                    wpauxa.X = UserList(UserIndex).flags.TargetObjX
                    wpauxa.Y = UserList(UserIndex).flags.TargetObjY
                    If Distancia(wpauxa, UserList(UserIndex).Pos) > 2 Then
                        Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y5")
                        Exit Sub
                    ElseIf Distancia(wpauxa, UserList(UserIndex).Pos) = 0 Then
                        Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y240")
                        Exit Sub
                    End If
                    
                    If ObjData(auxinda).ObjType <> OBJTYPE_ARBOLES Then
                    Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y381")
                    Exit Sub
                    End If
                    'Si no talas en un arbol de tejo
                    If auxinda = 634 And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_DORADA Then
                        Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y278")
                        Exit Sub
                    End If
                    
                    If auxinda <> 634 And UserList(UserIndex).Invent.HerramientaEqpObjIndex = HACHA_DORADA Then
                        Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y279")
                        Exit Sub
                    End If
                    
                    UserList(UserIndex).TyTrabajo = 2
                    UserList(UserIndex).TyTrabajoMod = auxinda
                    Call DameSuerte(UserIndex, 2)
                    UserList(UserIndex).flags.Trabajando = True
                   Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "EMPT")
                Else
                    Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y241")
                    Exit Sub
                End If
            '//////////////////////////////////////////////
            Case PIQUETE_MINERO, 684
                auxinda = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                If ObjData(auxinda).ObjType = OBJTYPE_YACIMIENTO Then
                    wpauxa.Map = UserList(UserIndex).Pos.Map
                    wpauxa.X = UserList(UserIndex).flags.TargetObjX
                    wpauxa.Y = UserList(UserIndex).flags.TargetObjY
                    If Distancia(wpauxa, UserList(UserIndex).Pos) > 2 Then
                        Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y5")
                        Exit Sub
                    End If
                    Call DameSuerte(UserIndex, 3)
                    UserList(UserIndex).TyTrabajo = 3
                    UserList(UserIndex).flags.Trabajando = True
                    Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "EMPT")
                Else
                    Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y242")
                    Exit Sub
                End If
            '//////////////////////////////////////////////
           Case hierrocrudo, orocrudo, platacruda
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                Call DameSuerte(UserIndex, 4)
                UserList(UserIndex).TyTrabajo = 4
                UserList(UserIndex).flags.Trabajando = True
                Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "EMPT")
                Else
                Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "Y247")
                Exit Sub
                End If
            End Select
            Else ' no tiene herramientas
            Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||Debes equiparte una herramienta para trabajar.~65~190~156~0~0")
            End If 'fin de tiene herramientas?¿
            Exit Sub
minaria:
            Call Senddata(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||Debes clickear algun elemento para extraer materia prima del mismo.~65~190~156~0~0")
            Exit Sub
        Case "LEC"
            rdata = Right$(rdata, Len(rdata) - 3)
            Call Senddata(ToAdmins, 0, 0, "||Servidor> " & ReadField(1, rdata, 44) & ". Fue echado del server por " & ReadField(2, rdata, 44) & "." & FONTTYPE_SERVER)
            LogCheats (ReadField(1, rdata, 44) & ". Fue echado del server por " & ReadField(2, rdata, 44) & ".")
            UserList(UserIndex).Stats.Veceshechado = UserList(UserIndex).Stats.Veceshechado + 1
            Exit Sub
        Case "CIG"
            rdata = Right$(rdata, Len(rdata) - 3)
            X = Guilds.Count
            If CreateGuild(UserList(UserIndex).Name, UserList(UserIndex).Reputacion.Promedio, UserIndex, rdata) Then
                If X = 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y249")
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "||Felicidades has creado el clan numero " & X + 1 & " de Argentum!!!." & FONTTYPE_INFO)
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
                        Call Senddata(ToIndex, UserIndex, 0, "Y250")
                        Call Senddata(ToIndex, UserIndex, 0, "||Nombre:" & Hechizos(H).Nombre & FONTTYPE_INFO)
                        Call Senddata(ToIndex, UserIndex, 0, "||Descripcion:" & Hechizos(H).Desc & FONTTYPE_INFO)
                        Call Senddata(ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(H).MinSkill & " de magia." & FONTTYPE_INFO)
                        Call Senddata(ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(H).ManaRequerido & FONTTYPE_INFO)
                        Call Senddata(ToIndex, UserIndex, 0, "||Stamina necesaria: " & Hechizos(H).StaRequerido & FONTTYPE_INFO)
                        Call Senddata(ToIndex, UserIndex, 0, "Y251")
                    End If
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y252")
                End If
                Exit Sub
                
                
        'el usuario reporta un md5
        Case "RMDC"
            UserList(UserIndex).flags.MD5Reportado = UCase$(Right$(rdata, Len(rdata) - 4))
            Exit Sub
        Case "EQUI"
                If UserList(UserIndex).flags.Muerto = 1 Then
                Call Senddata(ToIndex, UserIndex, 0, "Y26")
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
            If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Sub
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 3 Then Exit Sub
            rdata = Right$(rdata, Len(rdata) - 4)
            If Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rdata) > 0 And val(rdata) < Npclist(UserList(UserIndex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNPC).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Pos, True, False)
                        If SpawnedNpc <= MAXNPCS Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNPC
                            Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas = Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas + 1
                        End If
                End If
            Else
                Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            End If
            Exit Sub
        Case "COMP"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call Senddata(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                       Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rdata = Right$(rdata, Len(rdata) - 5)
             'User compra el item del slot rdata
             If UserList(UserIndex).flags.Comerciando = False Then
                Call Senddata(ToIndex, UserIndex, 0, "Y253")
                Exit Sub
             End If
             'listindex+1, cantidad
             Call NPCVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).flags.TargetNPC)
             Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call Senddata(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 4 Then
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
                       Call Senddata(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                       Call Senddata(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
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
                       Call Senddata(ToIndex, UserIndex, 0, "Y3")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿El NPC puede comerciar?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 4 Then
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
                        Call Senddata(ToIndex, UserIndex, 0, "Y210")
                        Exit Sub
                    End If
                Else
                    If ObjData(UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex).Newbie = 1 Then
                    Exit Sub
                    End If
                    
                    If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y210")
                        Exit Sub
                    End If
                End If
                '[Consejeros]
                If UserList(UserIndex).ComUsu.Objeto > 0 Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y254")
                    Exit Sub
                End If
                UserList(UserIndex).ComUsu.Objeto = val(Arg1)
                UserList(UserIndex).ComUsu.cant = val(Arg2)
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Sub
                Else
                    '[CORREGIDO]
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call Senddata(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha cambiado su oferta." & FONTTYPE_TALK)
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

Public Sub DameSuerte(ByVal UserIndex As Integer, Laburo As Byte)
Dim Suerte As String

Select Case Laburo
Case 1 'PESCAR
            Select Case UserList(UserIndex).Stats.UserSkills(Pesca)
            Case 0:         Suerte = 28
            Case 1 To 10:   Suerte = 26
            Case 11 To 20:  Suerte = 24
            Case 21 To 30:  Suerte = 22
            Case 31 To 40:  Suerte = 20
            Case 41 To 50:  Suerte = 18
            Case 51 To 60:  Suerte = 16
            Case 61 To 70:  Suerte = 14
            Case 71 To 80:  Suerte = 12
            Case 81 To 90:  Suerte = 10
            Case 91 To 99:  Suerte = 7
            Case Else:      Suerte = 5
            End Select
Case 2 'TALAR
            Select Case UserList(UserIndex).Stats.UserSkills(Talar)
            Case 0:         Suerte = 28
            Case 1 To 10:   Suerte = 26
            Case 11 To 20:  Suerte = 24
            Case 21 To 30:  Suerte = 22
            Case 31 To 40:  Suerte = 20
            Case 41 To 50:  Suerte = 18
            Case 51 To 60:  Suerte = 16
            Case 61 To 70:  Suerte = 14
            Case 71 To 80:  Suerte = 12
            Case 81 To 90:  Suerte = 10
            Case 91 To 99:  Suerte = 7
            Case Else:      Suerte = 5
            End Select
Case 3 'MINAR
            Select Case UserList(UserIndex).Stats.UserSkills(Mineria)
            Case 0:         Suerte = 28
            Case 1 To 10:   Suerte = 26
            Case 11 To 20:  Suerte = 24
            Case 21 To 30:  Suerte = 22
            Case 31 To 40:  Suerte = 20
            Case 41 To 50:  Suerte = 18
            Case 51 To 60:  Suerte = 16
            Case 61 To 70:  Suerte = 14
            Case 71 To 80:  Suerte = 12
            Case 81 To 90:  Suerte = 10
            Case 91 To 99:  Suerte = 7
            Case Else:      Suerte = 5
            End Select
Case 4 'LINGOTEAR
            If UserList(UserIndex).Stats.UserSkills(Mineria) <= 25 Then
            Suerte = 1
            ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 50 Then
            Suerte = 2
            ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 75 Then
            Suerte = 3
            Else
            Suerte = 4
            End If
End Select
UserList(UserIndex).Suerte = Suerte
End Sub
