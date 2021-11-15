Attribute VB_Name = "Mod_TCP"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
Const clave = "!*%!&?¡\¿@°>$<" 'clave de encriptación
Public macmarch As Boolean
Public Listaintegrantes(0 To 20) As String
Public Listasolicitudes(0 To 20) As String
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean
Private Enum SentidoRotacion
    ROTIzquierda = 0
    ROTDerecha = 1
End Enum

Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim RetVal As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim CharIndex As Integer
    Dim TempInt As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim I As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    Dim T() As String
    
    Dim sData As String
    sData = UCase(Rdata)
    
    #If LOG_DEBUG = 1 Then
        LogDebug ("HandleData ---> " & Rdata)
    #End If
    Debug.Print "BEGIN>>> " & Rdata
    
    Select Case sData
        Case "LOGGED"            ' >>>>> LOGIN :: LOGGED
            logged = True
            UserCiego = False
            EngineRun = True
            IScombate = False
            Istrabajando = False
            UserDescansar = False
            Nombres = True
            If frmCrearPersonaje.Visible Then
                   Unload frmPasswd
                   Unload frmPasswdSinPadrinos
                   Unload frmCrearPersonaje
                   Unload frmConnect
                   frmMain.Show
            End If
            Call SetConnected
            'Mostramos el Tip
            If tipf = "1" And PrimeraVez Then
                 Call CargarTip
                 frmtip.Visible = True
                 PrimeraVez = False
            End If
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Call DoFogataFx
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "FINOK" ' Graceful exit ;))
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            frmMain.Visible = False
            logged = False
            frmMain.piquete.enabled = False
            frmMain.SoundFX.enabled = False
            IMC.Stop
            UserParalizado = False
            IScombate = False
             Istrabajando = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            frmConnect.Visible = True
            Call frmMain.StopSound
            frmMain.IsPlaying = plNone
            bRain = False
            '[Misery_Ezequiel 10/07/05]
            bSnow = False
            '[\]Misery_Ezequiel 10/07/05]
            bNoche = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For I = 1 To LastChar
                CharList(I).invisible = False
            Next I
            bO = 0
            bK = 0
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            I = 1
            Do While I <= UBound(UserInventory)
                If UserInventory(I).OBJIndex <> 0 Then
                        frmComerciar.List1(1).AddItem UserInventory(I).Name
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                I = I + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim ii As Integer
            ii = 1
            Do While ii <= UBound(UserInventory)
                If UserInventory(ii).OBJIndex <> 0 Then
                        frmBancoObj.List1(1).AddItem UserInventory(ii).Name
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                ii = ii + 1
            Loop
            
            I = 1
            Do While I <= UBound(UserBancoInventory)
                If UserBancoInventory(I).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(I).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                I = I + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For I = 1 To UBound(UserInventory)
                If UserInventory(I).OBJIndex <> 0 Then
                        frmComerciarUsu.List1.AddItem UserInventory(I).Name
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = UserInventory(I).Amount
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next I
            Comerciando = True
            frmComerciarUsu.Show , frmMain
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]
        Case "RECPASSOK"
            Call MsgBox("¡¡¡El password fue enviado con éxito!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "RECPASSER"
            Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, "La criatura fallo el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
            
            Exit Sub
    End Select

    Select Case Mid(sData, 1, 1)
        Case "Y"
            Rdata = right(Rdata, Len(Rdata) - 1)
            Dim values As Variant
            values = Split(Rdata, ",")
            Dim params As Variant
            params = Split(Mensaje(values(0)), "~")
            Dim msg_texto As String
            Dim msg_red As Integer, msg_green As Integer, msg_blue As Integer
            Dim msg_bold As Boolean, msg_italic As Boolean
            msg_texto = params(0)
            msg_red = params(1)
            msg_green = params(2)
            msg_blue = params(3)
            msg_bold = params(4)
            msg_italic = params(5)
            '<gorlok 2005-03-28> Agrego la recepción de valores parámetros.
            Dim iv As Integer
            For iv = 1 To UBound(values)
                msg_texto = Replace(msg_texto, "#", values(iv), 1, 1)
            Next iv
            '</gorlok 2005-03-28>
            Call AddtoRichTextBox(frmMain.RecTxt, msg_texto, msg_red, msg_green, msg_blue, msg_bold, msg_italic)
            Exit Sub
    End Select

    Select Case sData
         Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha matado!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, "Has rechazado el ataque con el escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, "El usuario rechazo el ataque con su escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, "Has fallado el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "REAU"
            Call frmMain.DibujarSatelite
            Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, ">>SEGURO ACTIVADO<<", 0, 255, 0, True, False, False)
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, ">>SEGURO DESACTIVADO<<", 255, 0, 0, True, False, False)
            Exit Sub
    End Select

    Select Case left(sData, 2)
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            'Obtiene la version del mapa
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , TempInt
                Close #1
                frmMain.Coord.Caption = Mapa(UserMap)
                If TempInt = Val(ReadField(2, Rdata, 44)) Then
                Terreno = ReadField(3, Rdata, 44)
              Zona = ReadField(4, Rdata, 44)
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
                    If bLluvia(UserMap) = 0 Then
                        If bRain Then
                            IMC.Stop
                        End If
                    Else
                         If bRain Then
                            IMC.Run
                        End If
                    End If
                    
            '[Misery_Ezequiel 10/07/05]
                If bNieva(UserMap) = 0 Then
                    If bSnow Then
                        frmMain.StopSound
                        frmMain.IsPlaying = plNone
                    End If
                End If
            '[\]Misery_Ezequiel 10/07/05]
                Else
                    'vers incorrecta
                    MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                    Call LiberarObjetosDX
                    Call UnloadAllForms
                    End
                End If
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call LiberarObjetosDX
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        Case "CI"
            Dim veces As Integer
            Dim monkey As Variant
            Dim monkey2 As Variant
            Rdata = right(Rdata, Len(Rdata) - 2)
            monkey = Split(Rdata, ";")
            veces = monkey(0)
            For I = 1 To veces - 1
            monkey2 = Split(monkey(I), ",")
            X = monkey2(1)
            Y = monkey2(2)
            MapData(X, Y).ObjGrh.GrhIndex = monkey2(0)
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Next I
        
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.X, UserPos.Y).CharIndex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
            CharList(UserCharIndex).Pos = UserPos
            Exit Sub
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case BTarget.bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado en la cabeza por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado el brazo derecho por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado la pierna izquierda por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado la pierna derecha por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado en el torso por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a la criatura por " & Rdata & "!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Rdata & " te ataco y fallo!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case BTarget.bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado en la cabeza por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado el brazo derecho por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado la pierna izquierda por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado la pierna derecha por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado en el torso por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case BTarget.bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la cabeza por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el brazo derecho por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la pierna izquierda por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la pierna derecha por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case BTarget.bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el torso por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
                'i = 1
                'Do While i <= iuser
                '    Dialogos.CrearDialogo ReadField(2, Rdata, 176), i, Val(ReadField(1, Rdata, 176))
                '    i = i + 1
                'Loop
            Else
                  If PuedoQuitarFoco Then _
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
            End If
            Exit Sub
            
        Case "|2"
        
    Rdata = right$(Rdata, Len(Rdata) - 2)
    Dim mens As String
 mens = "Ves a " & ReadField(1, Rdata, 44)
    Rdata = right$(Rdata, Len(Rdata) - (1 + Len(ReadField(1, Rdata, 44))))
           
    If ReadField(1, Rdata, 44) = 1 Then mens = mens + " <NEWBIE>"

          If Len(Rdata) > 3 Then  'No esta muerto
          
                If Mid(ReadField(3, Rdata, 44), 1, 1) = "1" Then 'caos
                mens = mens & RangoCaos(Mid(ReadField(3, Rdata, 44), 2, 1))
                    If Mid(ReadField(3, Rdata, 44), 3, 1) = "1" Then
                    mens = mens & " [CONSEJO DE LAS SOMBRAS]"
                    End If
                Else
                    If Mid(ReadField(2, Rdata, 44), 1, 1) = "1" Then 'ARMADA
                    mens = mens & RangoArmada(Mid(ReadField(2, Rdata, 44), 2, 1)) '[Wizard 03/09/05]=> El error de los rangos reales estaba en que buscaba el 3 digito, que es el q respecta al Consejo de bander, el correcto es el 2.
                        If Mid(ReadField(2, Rdata, 44), 3, 1) = "1" Then
                        mens = mens & " [CONSEJO DE BANDERBILL]"
                        End If
                    End If
                End If
                            
                If ReadField(4, Rdata, 44) <> Empty Then
                mens = mens & " <" & ReadField(4, Rdata, 44) & ">"
                End If
                
                If ReadField(5, Rdata, 44) <> Empty Then
                mens = mens & " - " & ReadField(5, Rdata, 44)
                End If

                Select Case ReadField(6, Rdata, 44)
                Case 1
                AddtoRichTextBox frmMain.RecTxt, mens & " <CRIMINAL> ", 255, 0, 0, True, False
                Case 2
                AddtoRichTextBox frmMain.RecTxt, mens & " <CONSEJERO> ", 0, 150, 0, True, False
                Case 0
                AddtoRichTextBox frmMain.RecTxt, mens & " <CIUDADANO> ", 0, 0, 200, True, False
                Case 3
                AddtoRichTextBox frmMain.RecTxt, mens & " <GAME MASTER> ", 0, 185, 0, True, False
                Case 4
                AddtoRichTextBox frmMain.RecTxt, mens & " <GAME MASTER> ", 255, 255, 100, True, False
                Case 6 'Consejo de banderbill
                AddtoRichTextBox frmMain.RecTxt, mens & " <CIUDADANO> ", 0, 195, 255, True, False
                Case 7 'Consilio de las sombras.
                AddtoRichTextBox frmMain.RecTxt, mens & " <CRIMINAL> ", 255, 50, 0, True, False
                Case 8 'Combat
                AddtoRichTextBox frmMain.RecTxt, mens, 80, 80, 80, True, False
                Case 9 'Combat
                AddtoRichTextBox frmMain.RecTxt, mens, 200, 200, 200, True, False
                Case 10 'combat
                AddtoRichTextBox frmMain.RecTxt, mens, 220, 220, 220, True, False
                Case 11 'combat
                AddtoRichTextBox frmMain.RecTxt, mens, 250, 100, 150, True, False
                End Select
           
            Else 'Esta muerto pobrecito
            AddtoRichTextBox frmMain.RecTxt, mens & " <MUERTO> ", 192, 192, 192, True, False
            End If
            
Exit Sub
        Case "PT" 'mensajes de clan NUUEVO Marche
        Rdata = right$(Rdata, Len(Rdata) - 2)
        If Not Activado Then
            AddtoRichTextBox frmMain.RecTxt, Rdata, 228, 199, 27, 0, 0, False
            Else
                If Len(Rdata) < 85 Then
                clantext5 = clantext4
                clantext4 = clantext3
                clantext3 = clantext2
                clantext2 = clantext1
                clantext1 = Rdata
                    Else
                    clantext5 = clantext4
                    clantext4 = clantext3
                    clantext3 = clantext2
                    clantext2 = Mid(Rdata, 1, 84) & "-"
                    clantext1 = Mid(Rdata, 85, Len(Rdata))
                End If
      End If
      Exit Sub
        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = CharList(UserCharIndex).Pos
            Exit Sub
            
        Case "RE"
            TiempoReto = 4
            frmMain.Pasarsegundo.enabled = True
           Exit Sub
        Case "CC"              ' >>>>> Crear un Personaje :: CC
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            
            CharList(CharIndex).Fx = Val(ReadField(9, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            CharList(CharIndex).Nombre = ReadField(12, Rdata, 44)
            CharList(CharIndex).Criminal = Val(ReadField(13, Rdata, 44))
            CharList(CharIndex).priv = Val(ReadField(14, Rdata, 44))
            
            Call MakeChar(CharIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
            Call RefreshAllChars
            Exit Sub
            
        
            
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "M1"             ' >>>>> Mover un Personaje ARRIBA
      
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
                
            If Fx = 0 Then
                If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    If UserCharIndex <> CharIndex Then
                    CharList(CharIndex).Difpos.Y = (CharList(CharIndex).Pos.Y + 1) - UserPos.Y
                    End If
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y - 1)
            Call RefreshAllChars
            Exit Sub
        Case "M2"             ' >>>>> Mover un Personaje DERECHA
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    If UserCharIndex <> CharIndex Then
                     CharList(CharIndex).Difpos.X = (CharList(CharIndex).Pos.X + 1) - UserPos.X
                    End If
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X + 1, CharList(CharIndex).Pos.Y)
            Call RefreshAllChars
            Exit Sub
        Case "M3"             ' >>>>> Mover un Personaje ABAJO
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                   If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    If UserCharIndex <> CharIndex Then
                     CharList(CharIndex).Difpos.Y = (CharList(CharIndex).Pos.Y - 1) - UserPos.Y
                    End If
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y + 1)
            Call RefreshAllChars
            Exit Sub
        Case "M4"             ' >>>>> Mover un Personaje IZQUIERDA
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    If UserCharIndex <> CharIndex Then
                    CharList(CharIndex).Difpos.X = (CharList(CharIndex).Pos.X - 1) - UserPos.X
                    End If
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X - 1, CharList(CharIndex).Pos.Y)
            Call RefreshAllChars
            Exit Sub
        Case "M5"             ' >>>>> Mover un Personaje ARRIBA DERECHA
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X + 1, CharList(CharIndex).Pos.Y + 1)
            Call RefreshAllChars
            Exit Sub
        Case "M6"             ' >>>>> Mover un Personaje ABAJO DERECHA
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X + 1, CharList(CharIndex).Pos.Y - 1)
            Call RefreshAllChars
            Exit Sub
        Case "M7"             ' >>>>> Mover un Personaje ABAJO IZQUIERDA
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X - 1, CharList(CharIndex).Pos.Y - 1)
            Call RefreshAllChars
            Exit Sub
        Case "M8"             ' >>>>> Mover un Personaje ARRIBA IZQUIERDA
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                    Call DoPasosFx(CharIndex)
                End If
            End If
            Call MoveCharbyPos(CharIndex, CharList(CharIndex).Pos.X - 1, CharList(CharIndex).Pos.Y + 1)
            Call RefreshAllChars
            Exit Sub
           '
        Case "FF"
            frmMain.Label2.Visible = True 'cmsg
            frmMain.Label3.Visible = True  'rmsg
            frmMain.Label5.Visible = True  'invi
            frmMain.Label9.Visible = True  'panelgm
            frmMain.Label11.Visible = True 'show sos y Trabajando
            Exit Sub
        Case "FZ"
        Call Retos.Show(vbModeless, frmMain)
        Exit Sub
        Case "FC"
    
            Rdata = LTrim(right(Rdata, Len(Rdata) - 2))
            Dim valuess As Variant
            Dim paramss As Variant
            paramss = Split(Rdata, "~")
            Dim msg_textos As String
            Dim msg_reds As Integer, msg_greens As Integer, msg_blues As Integer
            Dim msg_bolds As Boolean, msg_italics As Boolean
            msg_textos = paramss(0)
            msg_reds = paramss(1)
            msg_greens = paramss(2)
            msg_blues = paramss(3)
            msg_bolds = paramss(4)
            msg_italics = paramss(5)
            Call AddtoRichTextBox(Partym.RecTxt, msg_textos, msg_reds, msg_greens, msg_blues, msg_bolds, msg_italics)
      
        Exit Sub
        
        Case "GH"
        gh = True
        Exit Sub
        Case "SS" 'fundo party
        ss = True
        Exit Sub
                Case "MP"            ' >>>>> Mover un Personaje :: MP
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Fx = 0 Then
                'If Not UserNavegando And Val(ReadField(4, Rdata, 44)) <> 0 Then
                        If CharList(CharIndex).priv = 0 Or CharList(CharIndex).priv = 4 Then
                            Call DoPasosFx(CharIndex)
                        End If
                'Else
                        'FX navegando
                'End If
            End If
            Call MoveCharbyPos(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            Call RefreshAllChars
            Exit Sub
            
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).muerto = Val(ReadField(3, Rdata, 44)) = 500
            CharList(CharIndex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
            CharList(CharIndex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            CharList(CharIndex).Heading = Val(ReadField(4, Rdata, 44))
            CharList(CharIndex).Fx = Val(ReadField(7, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            TempInt = Val(ReadField(5, Rdata, 44))
            If TempInt <> 0 Then CharList(CharIndex).Arma = WeaponAnimData(TempInt)
            TempInt = Val(ReadField(6, Rdata, 44))
            If TempInt <> 0 Then CharList(CharIndex).Escudo = ShieldAnimData(TempInt)
            TempInt = Val(ReadField(9, Rdata, 44))
            If TempInt <> 0 Then CharList(CharIndex).Casco = CascoAnimData(TempInt)
            Call RefreshAllChars
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            If Musica = 0 Then
                Rdata = right$(Rdata, Len(Rdata) - 2)
                If Val(ReadField(1, Rdata, 45)) <> 0 Then
                    Stop_Midi
                    If Musica = 0 Then
                        'frmmain.Winsock1.SendData "A" & "2.mid"
                        Dim CurMidi As String
                        CurMidi = Val(ReadField(1, Rdata, 45)) & ".mid"
                        LoopMidi = Val(ReadField(2, Rdata, 45))
                        Call CargarMIDI(DirMidi & CurMidi)
                        Call Play_Midi
                    End If
                End If
            End If
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW
            If Fx = 0 Then
                Rdata = right$(Rdata, Len(Rdata) - 2)
               PlayWaveDS (Rdata & ".wav")
            End If
            Exit Sub
        Case "GL" 'Lista de guilds
            Rdata = right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            '[CODE 001]:MatuX
                If frmMain.IsPlaying <> plFogata Then
                    frmMain.StopSound
                    Call frmMain.Play("fuego.wav", True)
                    frmMain.IsPlaying = plFogata
                End If
            '[END]'
            Exit Sub
    End Select

    Select Case left(sData, 3)
         
        Case "VAL"                  ' >>>>> Validar Cliente :: VAL
            Rdata = right$(Rdata, Len(Rdata) - 3)
            'If frmBorrar.Visible Then
            bK = CLng(ReadField(1, Rdata, Asc(",")))
            bO = 100 'CInt(ReadField(1, Rdata, Asc(",")))
            bRK = ReadField(2, Rdata, Asc(","))
            If EstadoLogin = BorrarPj Then
                Call SendData("BORR" & frmBorrar.txtNombre.Text & "," & frmBorrar.txtPasswd.Text & "," & ValidarLoginMSG(CInt(Rdata)))
            ElseIf EstadoLogin = Normal Or EstadoLogin = CrearNuevoPj Then
                Call Login(ValidarLoginMSG(CInt(bRK)))
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje.Show vbModal
            End If
            Exit Sub
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "LLU"                  ' >>>>> LLuvia!
            If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            If Not bRain Then
                bRain = True
                If bLluvia(UserMap) <> 0 Then
                   Play_Song ("Lluviatds")
               End If
            Else
               bRain = False
            End If
            Exit Sub
'[Misery_Ezequiel 10/07/05]
        Case "NIE"                  ' >>>>> Nieve!
            If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            If Not bSnow Then
                bSnow = True
            Else
               If bNieva(UserMap) <> 0 Then
                        'Call frmMain.StopSound
                        'Call frmMain.Play("nieve.wav", False)
                        frmMain.IsPlaying = plNone
               End If
               bSnow = False
            End If
            Exit Sub
'[\]Misery_Ezequiel 10/07/05]
        Case "NOC" 'nocheeeee
'            Debug.Print Rdata
            IMC.Stop
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).Fx = Val(ReadField(2, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim N As String, n2 As String
            Rdata = right$(Rdata, Len(Rdata) - 3)
            N = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg N, n2
            frmMSG.Show , frmMain
            Exit Sub
            'marche 10 -9
        '[Misery_Ezequiel 05/06/05]
'****************************************************************
'****************************************************************
'****************************************************************
'NO TOCAR NI MODIFICAR, POSIBLES ERRORES SI LO HACES
'****************************************************************
'****************************************************************
'****************************************************************
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola Trabajar :: AYMT
            Rdata = right$(Rdata, Len(Rdata) - 3)
            N = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSGT.CrearGMmSg N, n2
            frmMSGT.Show , frmMain
            Exit Sub
'****************************************************************
'****************************************************************
'****************************************************************
'NO TOCAR NI MODIFICAR, POSIBLES ERRORES SI LO HACES
'****************************************************************
'****************************************************************
'****************************************************************
        '[\]Misery_Ezequiel 05/06/05]
        Case "T08"
        Call Mod_General.DejarDeTrabajars
        
        

        Case "PNI"
            For I = 0 To 20
            If Listasolicitudes(I) = "" Or Listasolicitudes(I) = right$(Rdata, Len(Rdata) - 3) Then
            Listasolicitudes(I) = right$(Rdata, Len(Rdata) - 3)
            Exit For
            End If
            Next
        '///////////////////////////////////////////////////////////
        Case "MAN" 'Reduccion de lag By Marche, Actualia la mana.
            Rdata = right$(Rdata, Len(Rdata) - 3)
            UserMinMAN = Rdata
            frmMain.Label14 = UserMinMAN & "/" & UserMaxMAN
            
        If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
        Else
                frmMain.MANShp.Width = 0
        End If
            
        If frmMain.MANShp.Width > 80 Then
            frmMain.Label14.ForeColor = &H0&
        ElseIf frmMain.MANShp.Width > 65 And frmMain.MANShp.Width <= 80 Then
            frmMain.Label14.ForeColor = &H404040
        ElseIf frmMain.MANShp.Width > 19 And frmMain.MANShp.Width <= 65 Then
            frmMain.Label14.ForeColor = &H808080
        ElseIf frmMain.MANShp.Width > 12 And frmMain.MANShp.Width <= 19 Then
            frmMain.Label14.ForeColor = &HC0C0C0
        ElseIf frmMain.MANShp.Width > 5 And frmMain.MANShp.Width <= 12 Then
            frmMain.Label14.ForeColor = &HE0E0E0
        ElseIf frmMain.MANShp.Width < 5 Then
            frmMain.Label14.ForeColor = &HFFFFFF
        End If
        
        Case "VID"
            UserMinHP = right$(Rdata, Len(Rdata) - 3)
            frmMain.Label15 = UserMinHP & "/" & UserMaxHP
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            If frmMain.Hpshp.Width > 68 Then
                frmMain.Label15.ForeColor = &H0&
            ElseIf frmMain.Hpshp.Width > 59 And frmMain.Hpshp.Width <= 68 Then
                frmMain.Label15.ForeColor = &H404040
            ElseIf frmMain.Hpshp.Width > 36 And frmMain.Hpshp.Width <= 59 Then
                frmMain.Label15.ForeColor = &HC0C0C0
            ElseIf frmMain.Hpshp.Width > 25 And frmMain.Hpshp.Width <= 36 Then
                frmMain.Label15.ForeColor = &HE0E0E0
            ElseIf frmMain.Hpshp.Width < 25 Then
                frmMain.Label15.ForeColor = &HFFFFFF
            End If
        
            If UserMinHP = 0 Then
                UserEstado = 1
                '[Wizard 03/09/05] Usando esto evitamos un envio insano de paquetes:P
                UserEstupido = False
                '[/Wizard]
            Else
                UserEstado = 0
            End If
            Exit Sub
        
        Case "ENE"
            UserMinSTA = right$(Rdata, Len(Rdata) - 3)
            frmMain.Label13 = UserMinSTA & "/" & UserMaxSTA
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
        
            If frmMain.STAShp.Width > 68 Then
                frmMain.Label13.ForeColor = &H0&
            ElseIf frmMain.STAShp.Width > 59 And frmMain.STAShp.Width <= 68 Then
                frmMain.Label13.ForeColor = &H404040
            ElseIf frmMain.STAShp.Width > 25 And frmMain.STAShp.Width <= 59 Then
                frmMain.Label13.ForeColor = &H808080
            ElseIf frmMain.STAShp.Width > 17 And frmMain.STAShp.Width <= 25 Then
                frmMain.Label13.ForeColor = &HC0C0C0
            ElseIf frmMain.STAShp.Width > 9 And frmMain.STAShp.Width <= 17 Then
                frmMain.Label13.ForeColor = &HE0E0E0
            ElseIf frmMain.STAShp.Width < 9 Then
                frmMain.Label13.ForeColor = &HFFFFFF
            End If
        
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = right$(Rdata, Len(Rdata) - 3)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
            frmMain.Exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
            '[Misery_Ezequiel 1/06/05]
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            frmMain.Label13 = UserMinSTA & "/" & UserMaxSTA
            frmMain.Label14 = UserMinMAN & "/" & UserMaxMAN
            frmMain.Label15 = UserMinHP & "/" & UserMaxHP
            frmMain.Label16 = UserMinHAM & "/" & UserMaxHAM
            frmMain.Label17 = UserMinAGU & "/" & UserMaxAGU
            '[\]Misery_Ezequiel 1/06/05]
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl

        If frmMain.MANShp.Width > 80 Then
            frmMain.Label14.ForeColor = &H0&
        ElseIf frmMain.MANShp.Width > 65 And frmMain.MANShp.Width <= 80 Then
            frmMain.Label14.ForeColor = &H404040
        ElseIf frmMain.MANShp.Width > 19 And frmMain.MANShp.Width <= 65 Then
            frmMain.Label14.ForeColor = &H808080
        ElseIf frmMain.MANShp.Width > 12 And frmMain.MANShp.Width <= 19 Then
            frmMain.Label14.ForeColor = &HC0C0C0
        ElseIf frmMain.MANShp.Width > 5 And frmMain.MANShp.Width <= 12 Then
            frmMain.Label14.ForeColor = &HE0E0E0
        ElseIf frmMain.MANShp.Width < 5 Then
            frmMain.Label14.ForeColor = &HFFFFFF
        End If

        If frmMain.Hpshp.Width > 68 Then
            frmMain.Label15.ForeColor = &H0&
        ElseIf frmMain.Hpshp.Width > 59 And frmMain.Hpshp.Width <= 68 Then
            frmMain.Label15.ForeColor = &H404040
        ElseIf frmMain.Hpshp.Width > 36 And frmMain.Hpshp.Width <= 59 Then
            frmMain.Label15.ForeColor = &HC0C0C0
        ElseIf frmMain.Hpshp.Width > 25 And frmMain.Hpshp.Width <= 36 Then
            frmMain.Label15.ForeColor = &HE0E0E0
        ElseIf frmMain.Hpshp.Width < 25 Then
            frmMain.Label15.ForeColor = &HFFFFFF
        End If
'Para la barra de stamina
        If frmMain.STAShp.Width > 68 Then
            frmMain.Label13.ForeColor = &H0&
        ElseIf frmMain.STAShp.Width > 59 And frmMain.STAShp.Width <= 68 Then
            frmMain.Label13.ForeColor = &H404040
        ElseIf frmMain.STAShp.Width > 25 And frmMain.STAShp.Width <= 59 Then
            frmMain.Label13.ForeColor = &H808080
        ElseIf frmMain.STAShp.Width > 17 And frmMain.STAShp.Width <= 25 Then
            frmMain.Label13.ForeColor = &HC0C0C0
        ElseIf frmMain.STAShp.Width > 9 And frmMain.STAShp.Width <= 17 Then
            frmMain.Label13.ForeColor = &HE0E0E0
        ElseIf frmMain.STAShp.Width < 9 Then
            frmMain.Label13.ForeColor = &HFFFFFF
        End If
'[\]Misery_Ezequiel 10/07/05]
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            Exit Sub
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
            End Select
            Exit Sub
           Case "PPP"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = right$(Rdata, Len(Rdata) - 3)
            PPP = Val(Rdata)
            Select Case PPP
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
            End Select
            Exit Sub

        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserInventory(slot).Name = ReadField(3, Rdata, 44)
            UserInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserInventory(slot).Equipped = ReadField(5, Rdata, 44)
            UserInventory(slot).GrhIndex = Val(ReadField(6, Rdata, 44))
            UserInventory(slot).ObjType = Val(ReadField(7, Rdata, 44))
            UserInventory(slot).MaxHit = Val(ReadField(8, Rdata, 44))
            UserInventory(slot).MinHit = Val(ReadField(9, Rdata, 44))
            UserInventory(slot).Def = Val(ReadField(10, Rdata, 44))
            UserInventory(slot).Valor = Val(ReadField(11, Rdata, 44))
        
            tempstr = ""
            If UserInventory(slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            If UserInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(slot).Amount & ") " & UserInventory(slot).Name
            Else
                tempstr = tempstr & UserInventory(slot).Name
            End If
            bInvMod = True
            Call frmMain.picInv.Refresh
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).ObjType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, Rdata, 44))
            tempstr = ""
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            bInvMod = True
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = right$(Rdata, Len(Rdata) - 3)
            For I = 1 To NUMATRIBUTOS
                UserAtributos(I) = Val(ReadField(I, Rdata, 44))
            Next I
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = right$(Rdata, Len(Rdata) - 3)
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = right$(Rdata, Len(Rdata) - 3)
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "OBR"
            Rdata = right$(Rdata, Len(Rdata) - 3)
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            Rdata = right(Rdata, Len(Rdata) - 3)
            For I = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(I + 1, Rdata, 44)
            Next I
            frmSpawnList.Show , frmMain
            Exit Sub
        Case "ERR"
            Rdata = right$(Rdata, Len(Rdata) - 3)
            frmOldPersonaje.MousePointer = 1
            frmPasswd.MousePointer = 1
            frmPasswdSinPadrinos.MousePointer = 1
            If Not frmCrearPersonaje.Visible Then
#If UsarWrench = 1 Then
                frmMain.Socket1.Disconnect
#Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
#End If
            End If
            MsgBox Rdata
            Exit Sub
    End Select
    
    Select Case left(sData, 4)
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).ObjType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            bInvMod = True
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmMain.Label13 = UserMinSTA & "/" & UserMaxSTA
            frmMain.Label14 = UserMinMAN & "/" & UserMaxMAN
            frmMain.Label15 = UserMinHP & "/" & UserMaxHP
            frmMain.Label16 = UserMinHAM & "/" & UserMaxHAM
            frmMain.Label17 = UserMinAGU & "/" & UserMaxAGU
'[Misery_Ezequiel 10/07/05]
'Para la barra de hambre
        If frmMain.COMIDAsp.Width > 68 Then
            frmMain.Label16.ForeColor = &H0&
        ElseIf frmMain.COMIDAsp.Width > 59 And frmMain.COMIDAsp.Width <= 68 Then
            frmMain.Label16.ForeColor = &H404040
        ElseIf frmMain.COMIDAsp.Width > 47 And frmMain.COMIDAsp.Width <= 59 Then
            frmMain.Label16.ForeColor = &H808080
        ElseIf frmMain.COMIDAsp.Width > 36 And frmMain.COMIDAsp.Width <= 47 Then
            frmMain.Label16.ForeColor = &HC0C0C0
        ElseIf frmMain.COMIDAsp.Width > 25 And frmMain.COMIDAsp.Width <= 36 Then
            frmMain.Label16.ForeColor = &HE0E0E0
        ElseIf frmMain.COMIDAsp.Width < 25 Then
            frmMain.Label16.ForeColor = &HFFFFFF
        End If
'Para la barra de sed
        If frmMain.AGUAsp.Width > 68 Then
            frmMain.Label17.ForeColor = &H0&
        ElseIf frmMain.AGUAsp.Width > 59 And frmMain.AGUAsp.Width <= 68 Then
            frmMain.Label17.ForeColor = &H404040
        ElseIf frmMain.AGUAsp.Width > 47 And frmMain.AGUAsp.Width <= 59 Then
            frmMain.Label17.ForeColor = &H808080
        ElseIf frmMain.AGUAsp.Width > 36 And frmMain.AGUAsp.Width <= 47 Then
            frmMain.Label17.ForeColor = &HC0C0C0
        ElseIf frmMain.AGUAsp.Width > 25 And frmMain.AGUAsp.Width <= 36 Then
            frmMain.Label17.ForeColor = &HE0E0E0
        ElseIf frmMain.AGUAsp.Width < 25 Then
            frmMain.Label17.ForeColor = &HFFFFFF
        End If
'[\]Misery_Ezequiel 10/07/05]
            Exit Sub
        Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
            Rdata = right$(Rdata, Len(Rdata) - 4)
            UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
            UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
            UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
            UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
            UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
            UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
            UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
            LlegoFama = True
            Exit Sub
        Case "MEST" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
                .CriminalesMatados = Val(ReadField(2, Rdata, 44))
                .UsuariosMatados = Val(ReadField(3, Rdata, 44))
                .NpcsMatados = Val(ReadField(4, Rdata, 44))
                .Clase = ReadField(5, Rdata, 44)
                .PenaCarcel = Val(ReadField(6, Rdata, 44))
            End With
            Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, "Hay " & Rdata & " npcs.", 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "TSOS"
            Rdata = right$(Rdata, Len(Rdata) - 4)
            frmMSGT.List1.AddItem Rdata
            Exit Sub
        Case "XXX1"
            Rdata = right$(Rdata, Len(Rdata) - 4)
            frmMSGT.List1.AddItem Rdata
            Exit Sub
         Case "XXX2"
            frmMSGT.Show , frmMain
            Exit Sub
        Case "EMPT"
        Istrabajando = True
        AddtoRichTextBox frmMain.RecTxt, "Empiezas a trabajar.", 65, 190, 156, False, False
        Case "DMPT"
        Call Mod_General.DejarDeTrabajars
        Case "TT47"
            Form3.Show , frmMain
        Case "ZZZZ"
            Dim mm As String
            mm = MsgBox("Hay una nueva version disponible. Si desea actualizarla pulse en si y el cliente se actualizara automaticamente. De lo contrario no podra seguir jugando.", vbExclamation + vbYesNo)
            If mm = vbYes Then
                If FileExist(App.Path & "\Updater.exe", vbNormal) Then
                Call Shell(App.Path & "\Updater.exe", vbNormalFocus)
                Else
                mm = MsgBox("El AutoUpdater-TDS no se encuentra instalado. Por favor descarguelo desde www.aotds.com.ar", vbCritical, "AutoUpdater - Tierras del Sur")
                End If
            End
            Else
            End If
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        '[Misery_Ezequiel 05/06/05]
'****************************************************************
'****************************************************************
'****************************************************************
'NO TOCAR NI MODIFICAR, POSIBLES ERRORES SI LO HACES
'****************************************************************
'****************************************************************
'****************************************************************
        Case "RSOST"             ' >>>>> Mensaje :: Trabajando
            Rdata = right$(Rdata, Len(Rdata) - 4)
            frmMSGT.List1.AddItem Rdata
            Exit Sub
        Case "TSOST"             ' >>>>> Mensaje :: Trabajando
            frmMSGT.Show , frmMain
            Exit Sub
'****************************************************************
'****************************************************************
'****************************************************************
'NO TOCAR NI MODIFICAR, POSIBLES ERRORES SI LO HACES
'****************************************************************
'****************************************************************
'****************************************************************
        '[\]Misery_Ezequiel 05/06/05]
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    Case "COSE"
            'UsandoSistemaPadrinos = ReadField(1, Rdata, 44)
            'PuedeCrearPjs = ReadField(1, Rdata, 44)
            'Exit Sub
        Case "RMDC"
            Rdata = UCase$(right$(Rdata, Len(Rdata) - 4))
            Dim asdf As String
            If UCase(Rdata) = "EXE" Then
                asdf = MD5File(App.Path & "\" & App.EXEName & ".exe")
                'asdf = "669423534a1063dcf5e0ba9d70983a2c"
            ElseIf FileExist(App.Path & "\" & Rdata, vbReadOnly) Then
                asdf = MD5File(App.Path & "\" & Rdata)
            Else
                asdf = "NO"
            End If
            #If UsarWrench = 1 Then
                    asdf = asdf & "-" & frmMain.Socket1.HostName & "-" & frmMain.Socket1.RemotePort & "-" & frmMain.Socket1.HostAddress
            #Else
                    asdf = asdf & "-" & frmMain.Winsock1.RemoteHostIP & "-" & frmMain.Winsock1.RemotePort & "-" & frmMain.Winsock1.RemoteHost
            #End If
            
            Call SendData("RMDC" & asdf)
            Exit Sub
    End Select
    
    Select Case left(sData, 5)
        Case "ZMOTD"
            Rdata = right$(Rdata, Len(Rdata) - 5)
            frmCambiaMotd.Show , frmMain
            frmCambiaMotd.txtMotd.Text = Rdata
            Exit Sub
        Case "DADOS"
            Rdata = right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)
                    
                    tempstr = ReadField(6, Rdata, 44)
                    If tempstr <> "" Then UsandoSistemaPadrinos = Val(tempstr)
                End If
            End With
            Exit Sub
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
        Case "NOVER"             ' >>>>> Invisible :: NOVER
            Rdata = right$(Rdata, Len(Rdata) - 5)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            Exit Sub
    End Select
    
    Select Case left(sData, 6)
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = right$(Rdata, Len(Rdata) - 6)
            For I = 1 To NUMSKILLS
                UserSkills(I) = Val(ReadField(I, Rdata, 44))
            Next I
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = right(Rdata, Len(Rdata) - 6)
            For I = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(I + 1, Rdata, 44)
            Next I
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case left(sData, 7)
        Case "GUILDNE"
            Rdata = right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
        Case "PEACEDE"
            Rdata = right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "PEACEPR"
            Rdata = right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
        Case "LEADERI"
            Rdata = right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "CLANDET"
            Rdata = right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                I = 1
                Do While I <= UBound(UserInventory)
                    If UserInventory(I).OBJIndex <> 0 Then
                            frmComerciar.List1(1).AddItem UserInventory(I).Name
                    Else
                            frmComerciar.List1(1).AddItem "Nada"
                    End If
                    I = I + 1
                Loop
                Rdata = right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
                Else
                        frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                I = 1
                Do While I <= UBound(UserInventory)
                    If UserInventory(I).OBJIndex <> 0 Then
                            frmBancoObj.List1(1).AddItem UserInventory(I).Name
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    I = I + 1
                Loop
                
                ii = 1
                Do While ii <= UBound(UserBancoInventory)
                    If UserBancoInventory(ii).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(ii).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    ii = ii + 1
                Loop
                
                Rdata = right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "ABPANEL"
            frmPanelGm.Show , frmMain
            Exit Sub
        Case "LISTUSU"
            Rdata = right(Rdata, Len(Rdata) - 7)
            T = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For I = LBound(T) To UBound(T)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem T(I)
                Next I
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
            End If
            Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase(left(Rdata, 9))
    Case "COMUSUINV"
        Rdata = right(Rdata, Len(Rdata) - 9)
        OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
        OtroInventario(1).Name = ReadField(3, Rdata, 44)
        OtroInventario(1).Amount = ReadField(4, Rdata, 44)
        OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
        OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
        OtroInventario(1).ObjType = Val(ReadField(7, Rdata, 44))
        OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
        OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
        OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
        OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
        frmComerciarUsu.List2.Clear
        frmComerciarUsu.List2.AddItem OtroInventario(1).Name
        frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
        frmComerciarUsu.lblEstadoResp.Visible = False
    End Select
End Sub

Private Sub Decrypt(ByVal s As String)
On Error GoTo errorH

Exit Sub
errorH:
Call LogCustom("error en decrypt: " & Err.description)

End Sub

Sub SendData(ByVal sdData As String)
Dim retcode
Dim AuxCmd As String

AuxCmd = UCase(left(sdData, 5))

'Debug.Print ">> " & sdData

bK = GenCrC(bK, sdData)

bO = bO + 1
If bO > 10000 Then bO = 100

'Agregamos el fin de linea
'sdData = sdData & "~" & bK & ENDC

sdData = sdData & "~" & bK
#If LOG_DEBUG = 1 Then
    LogDebug ("SendData:: enviar data: " & sdData)
#End If
' sdData = CryptStr(sdData) 'byGorlok
sdData = sdData & ENDC

'Para evitar el spamming
If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
    Exit Sub
ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
    Exit Sub
End If
#If UsarWrench = 1 Then
    retcode = frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If
End Sub

Sub Login(ByVal valcode As Integer)
'Personaje grabado
'If SendNewChar = False Then
If EstadoLogin = Normal Then
'marche


Dim nFic As Integer
Dim scadena As String
nFic = FreeFile
scadena = Space$(19)
Open (App.Path & "\oacslis.dll") For Binary As nFic
Get nFic, , scadena
scadena = Mod_TCP.Decripta(scadena)

Close nFic

SendData ("OLOGIN" & UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo & "," & Versiones(2) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & scadena)


'Crear personaje
'If SendNewChar = True Then
'Barrin 3/10/03
'mandamos diferentes datos de login nuevo a partir de si se esta usando o no el sistema de
'padrinos en el servidor
ElseIf EstadoLogin = CrearNuevoPj And UsandoSistemaPadrinos = 1 Then
    SendData ("NLOGIN" & UserName & "," & UserPassword _
    & "," & 0 & "," & 0 & "," _
    & App.Major & "." & App.Minor & "." & App.Revision & _
    "," & UserRaza & "," & UserSexo & "," & UserClase & "," & _
    UserAtributos(1) & "," & UserAtributos(2) & "," & UserAtributos(3) _
    & "," & UserAtributos(4) & "," & UserAtributos(5) _
     & "," & UserSkills(1) & "," & UserSkills(2) _
     & "," & UserSkills(3) & "," & UserSkills(4) _
     & "," & UserSkills(5) & "," & UserSkills(6) _
     & "," & UserSkills(7) & "," & UserSkills(8) _
     & "," & UserSkills(9) & "," & UserSkills(10) _
     & "," & UserSkills(11) & "," & UserSkills(12) _
     & "," & UserSkills(13) & "," & UserSkills(14) _
     & "," & UserSkills(15) & "," & UserSkills(16) _
     & "," & UserSkills(17) & "," & UserSkills(18) _
     & "," & UserSkills(19) & "," & UserSkills(20) _
     & "," & UserSkills(21) & "," & UserEmail & "," _
     & UserHogar & "," & PadrinoName & "," & PadrinoPassword & "," & Versiones(1) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & valcode & MD5HushYo)
ElseIf EstadoLogin = CrearNuevoPj And UsandoSistemaPadrinos = 0 Then
    SendData ("NLOGIN" & UserName & "," & UserPassword _
    & "," & 0 & "," & 0 & "," _
    & App.Major & "." & App.Minor & "." & App.Revision & _
    "," & UserRaza & "," & UserSexo & "," & UserClase & "," & _
    UserAtributos(1) & "," & UserAtributos(2) & "," & UserAtributos(3) _
    & "," & UserAtributos(4) & "," & UserAtributos(5) _
     & "," & UserSkills(1) & "," & UserSkills(2) _
     & "," & UserSkills(3) & "," & UserSkills(4) _
     & "," & UserSkills(5) & "," & UserSkills(6) _
     & "," & UserSkills(7) & "," & UserSkills(8) _
     & "," & UserSkills(9) & "," & UserSkills(10) _
     & "," & UserSkills(11) & "," & UserSkills(12) _
     & "," & UserSkills(13) & "," & UserSkills(14) _
     & "," & UserSkills(15) & "," & UserSkills(16) _
     & "," & UserSkills(17) & "," & UserSkills(18) _
     & "," & UserSkills(19) & "," & UserSkills(20) _
     & "," & UserSkills(21) & "," & UserEmail & "," _
     & UserHogar & "," & Versiones(1) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & valcode & MD5HushYo)
End If
End Sub

Public Function Decripta(ByVal strPassword As String) As String
Dim LongOrigen As Long
Dim I As Integer, j As Integer
Dim flag1 As String
Dim flag2 As String
Dim codigo As Long
Dim strTexto As String
LongOrigen = Len(strPassword)
For I = 1 To LongOrigen
If I > Len(clave) Then
j = 1
Else
j = I
End If
flag1 = Mid(clave, j, 1)
flag2 = Mid(strPassword, I, 1)
codigo = Asc(flag1) Xor Asc(flag2)
strTexto = strTexto & Chr(codigo)
Next
Decripta = strTexto
End Function
'********************Misery_Ezequiel 28/05/05********************'


Private Sub Cheats()

End Sub

Public Function STI(ByVal str As String, ByVal start As Byte) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim tempstr As String
    
    'Asergurarse sea válido
    If Len(str) < start - 1 Then Exit Function
    'Convertimos a hexa el valor ascii del segundo Byte
    tempstr = hex$(Asc(Mid$(str, start + 1, 1)))
    
    'Nos aseguramos tenga 2 Bytes (los ceros a la izquierda cuentan por ser el segundo Byte)
    While Len(tempstr) < 2
        tempstr = "0" & tempstr
    Wend
    
    'Convertimos a integer
    STI = Val("&H" & hex$(Asc(Mid$(str, start, 1))) & tempstr)
    
    'Vemos si el primer Byte era cero
    If STI And &H8000 Then _
        STI = STI Xor &H8001
    
    'Si el segundo Byte era cero
    If STI And &H4000 Then _
        STI = STI Xor &H4000
End Function

Public Function STI2(ByVal str As String, ByVal start As Single) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim tempstr As String
    
    'Asergurarse sea válido
    If Len(str) < start - 1 Then Exit Function
    'Convertimos a hexa el valor ascii del segundo Byte
    tempstr = hex$(Asc(Mid$(str, start + 1, 1)))
    
    'Nos aseguramos tenga 2 Bytes (los ceros a la izquierda cuentan por ser el segundo Byte)
    While Len(tempstr) < 2
        tempstr = "0" & tempstr
    Wend
    
    'Convertimos a integer
    STI2 = Val("&H" & hex$(Asc(Mid$(str, start, 1))) & tempstr)
    
    'Vemos si el primer Byte era cero
    If STI2 And &H8000 Then _
        STI2 = STI2 Xor &H8001
    
    'Si el segundo Byte era cero
    If STI2 And &H4000 Then _
        STI2 = STI2 Xor &H4000
End Function

