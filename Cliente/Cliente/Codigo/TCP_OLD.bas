Attribute VB_Name = "TCP_OLD"
#If False Then
'---------------------------------------------------------------------------------------
' Procedure : ProcesarPaquete
' DateTime  : 26/02/2007 21:35
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub ProcesarPaquete(ByVal Rdata As String)
Dim paquete As Byte
   'On Error GoTo ProcesarPaquete_Error

'On Error Resume Next
    If LenB(Rdata) = 0 Then Exit Sub
    If Grabando Then
        CrearAccion (Rdata)
        'Debug.Print TempPaq(PS).TC & "  " & TempPaq(PS).Rdata
        'Put #12, , TempPaq
        'DoEvents
    End If

    paquete = Asc(Left$(Rdata, 1))
    If paquete = 0 Then
    MsgBox "LA PUTA MADRE, EL PAQUETE ES CERO UUOOOUUOUOUOUOUOUOUOOU"
    End If

    TempStr = Left$(Rdata, 1)
    If Len(Rdata) > 1 Then 'Hay argumentos:O
        Rdata = Right$(Rdata, Len(Rdata) - 1)
    Else
        Rdata = vbNullString
    End If



    Debug.Print "LLEGADA PAQUTE>>> " & Asc(TempStr); Rdata

    Select Case paquete 'Asc(TempStr)
        '---------------------------------------------
        Case sPaquetes.pNpcInventory
           If Comerciando Then frmComerciar.NpcInventarioComercio.Cls
            For tempint = 1 To 25
                If Left(Rdata, 1) <> "ÿ" And LenB(Rdata) > 1 Then
                NPCInventory(tempint).OBJType = Asc(Left$(Rdata, 1))
                NPCInventory(tempint).Amount = STI(Rdata, 2)
                NPCInventory(tempint).GrhIndex = STI(Rdata, 4)
                NPCInventory(tempint).OBJIndex = STI(Rdata, 6)
                NPCInventory(tempint).MaxHit = STI(Rdata, 8)
                NPCInventory(tempint).MinHit = STI(Rdata, 10)
                NPCInventory(tempint).MinDef = StringToByte(Rdata, 12)
                NPCInventory(tempint).MaxDef = StringToByte(Rdata, 13)
                NPCInventory(tempint).Valor = StringToLong(Rdata, 14)
                NPCInventory(tempint).name = Objeto(NPCInventory(tempint).OBJIndex)
                Rdata = mid$(Rdata, 18)
             Else
                NPCInventory(tempint).Amount = 0
                NPCInventory(tempint).GrhIndex = 0
                NPCInventory(tempint).OBJIndex = 0
                NPCInventory(tempint).OBJType = 0
                NPCInventory(tempint).MaxHit = 0
                NPCInventory(tempint).MinHit = 0
                NPCInventory(tempint).MaxDef = 0
                NPCInventory(tempint).MinDef = 0
                NPCInventory(tempint).Valor = 0
                NPCInventory(tempint).name = ""
                Rdata = mid$(Rdata, 2)
            End If
            Next
            NPCInvDim = 25
              '  If NPCInventory(NPCInvDim).Name <> "" Then
                 '   frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
              '  Else
                '   frmComerciar.List1(0).AddItem "Nada"
               ' End If
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.TransOK
            If frmComerciar.Visible Then
            Call Dibujar(frmComerciar.ItemElegidoV, frmComerciar.ComercioInventario, UserInventory, 6)
            Call DibujarNpcInv(frmComerciar.ItemElegidoC, frmComerciar.NpcInventarioComercio, NPCInventory, 6, NPCInvDim)
            End If
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.pIniciarComercioNpc
          '  TempByte = 1
           ' Do While TempByte <= UBound(UserInventory)
             '   If UserInventory(TempByte).OBJIndex <> 0 Then
                '        frmComerciar.List1(1).AddItem UserInventory(TempByte).Name
              '  Else
                       ' frmComerciar.List1(1).AddItem "Nada"
              '  End If
               ' TempByte = TempByte + 1
          '  Loop
            If Comerciando Then Exit Sub
            Comerciando = True
            Call Dibujar(1, frmComerciar.ComercioInventario, UserInventory, 6)
            Call DibujarNpcInv(0, frmComerciar.NpcInventarioComercio, NPCInventory, 6, NPCInvDim)
            Call frmComerciar.Show(vbModeless, frmMain)
            Exit Sub
            '---------------------------------------------
        Case sPaquetes.pMensajeSimple 'Simple
            Rdata = (Asc(Rdata))
            Tempvar = Split(Mensaje(Rdata), "~")
            Rdata = Tempvar(0)
            Call AddtoRichTextBox(frmMain.RecTxt, Rdata, Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)), Int(Tempvar(5)))
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.pMensajeCompuesto 'Mensaje compuestos

            TempByte = Asc(Left$(Rdata, 1))
            'Numero de Mensaje
            TempStr = MensajesCompuestos(TempByte)
            ' Mensaje
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            'Sacamos el numero
            If TempByte = 39 Then
                TempStr = Replace(TempStr, "#1", Rdata)
            ElseIf InStr(1, Rdata, ",") Then
                Tempvar = Split(Rdata, ",")
                For TempByte2 = 0 To UBound(Tempvar)
                TempStr = Replace(TempStr, "#" & TempByte2 + 1, Tempvar(TempByte2))
                Next
            ElseIf LenB(Rdata) > 1 Then
                TempStr = Replace(TempStr, "#1", Rdata)
            End If
            Tempvar = Split(mid(TempStr, InStr(1, TempStr, "~") - 1), "~")
            Call AddtoRichTextBox(frmMain.RecTxt, mid(TempStr, 1, InStr(1, TempStr, "~") - 1), Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)), Int(Tempvar(5)))
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.EnPausa
            pausa = Not pausa
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.PrenderFogata 'Fogata
            bFogata = True
              '  If frmMain.IsPlaying <> plFogata Then
                  '  frmMain.StopSound
                  '  Call frmMain.Play("fuego.wav", True)
                  '  frmMain.IsPlaying = plFogata
              '  End If
            Exit Sub
            '---------------------------------------------
        'Case sPaquetes.MensajeForo 'Lee mensaje
            'frmForo.List.AddItem ReadField(1, Rdata, 176)
           ' frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
           ' Load frmForo.Text(frmForo.List.ListCount)
        'Exit Sub
            '---------------------------------------------
        'Case sPaquetes.MensajeForo2 'Carga foro
         '   If Not frmForo.Visible Then
          '        frmForo.Show
           ' End If
       ' Exit Sub
            '---------------------------------------------
        Case sPaquetes.WavSnd 'WAV
             Rdata = Asc(Rdata)
             If SoundActivated = 1 Then
                Call Audio.Sound_Play(val(Rdata))
            End If
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.MostrarCartel 'Muestra cartel
            Call InitCartel(ReadField(1, Rdata, 199), CInt(ReadField(2, Rdata, 199)))
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.VeObjeto 'Clickeo un Objeto
            Dim Quantity As Integer
            Quantity = STI(Rdata, 3)
            If Quantity <> 1 Then

               ' Call AddtoRichTextBox(frmMain.RecTxt, "Ves " & Quantity & " " & ObjetosPlural(STI(Rdata, 1)), 255, 2, 2, False, False)
            Else
                'Call AddtoRichTextBox(frmMain.RecTxt, "Ves 1 " & Objetos(STI(Rdata, 1)), 255, 2, 2, False, False)
            End If
        Exit Sub
              '---------------------------------------------
        Case sPaquetes.VeUser
            'Vasado en la idea de marce "|2" pero total
            'mente remodelado por mi[Wizard]
            If InStr(1, Rdata, "Ç") = 0 Then
            'Esta Muerto y leemos el (Newbie = not Newbie)
            'tenemos q agregar si es 0 o 1 por si ay
            'un tag de una letra..¬¬
                If Len(Rdata) > 2 Then
                TempStr = "Ves a " & CharList(STI(Right$(Rdata, Len(Rdata) - 1), 1)).Nombre & " <NEWBIE>"
                Else
                TempStr = "Ves a " & CharList(STI(Right$(Rdata, Len(Rdata)), 1)).Nombre
                End If
                TempStr = TempStr & " <MUERTO>"
                AddtoRichTextBox frmMain.RecTxt, TempStr, 192, 192, 192, True
            Else 'Esta vivo
                Tempvar = Split(Rdata, "Ç")
                tempint = STI(Tempvar(0), 1)

                TempStr = "Ves a " & Replace(CharList(STI(Tempvar(0), 3)).Nombre, "<", " <")

                If Right$(Rdata, 1) = "1" Then
                TempStr = TempStr & " <NEWBIE>"
                Tempvar(1) = Left$(Tempvar(1), Len(Tempvar(1)) - 1)
                End If

                If mid$(tempint, 2, 1) = "1" Then
                    TempStr = TempStr & " " & RangoArmada(mid$(tempint, 3, 1))
                ElseIf mid$(tempint, 2, 1) = "2" Then
                    TempStr = TempStr & " " & RangoCaos(mid$(tempint, 3, 1))
                End If
                'Agregamos el clan si tiene
                If Tempvar(1) <> "" Then TempStr = TempStr & " - " & Tempvar(1)
                'Terminamos el Mensaje agregamos el ultimo str y
                'Mandamos con color en especial
                Select Case mid(tempint, 1, 1)
                    Case 8  'Ciudadano
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " <CIUDADANO>", 0, 0, 200, True
                    Case 1 'Criminal
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " <CRIMINAL>", 255, 0, 0, True
                    Case 2 'Consejero
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " <CONSEJERO>", 0, 180, 0, True
                    Case 3 'Semidios
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " <SEMIDIOS>", 0, 230, 0, True
                    Case 4 'Dios
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " <DIOS>", 250, 250, 150, True
                    Case 5 'Administrador
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " <ADMINISTRADOR>", 255, 165, 0, True
                    Case 6 'Consejo de Bander
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " [CONSEJO DE BANDERBILL]", 0, 125, 200, True
                    Case 7 'Consilio de las sombras
                        AddtoRichTextBox frmMain.RecTxt, TempStr & " [CONCILIO DE LAS SOMBRAS]", 100, 100, 100, True
                    Case 9 'Mimetizado
                        AddtoRichTextBox frmMain.RecTxt, TempStr, 215, 215, 215, True
                End Select
            End If

            Exit Sub
            '---------------------------------------------
            Case sPaquetes.VeNpc
                tempint = STI(Rdata, 1)
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                If Len(Rdata) > 8 Then ' then es Mascota
                    'Call AddtoRichTextBox(frmMain.RecTxt, Npcs(tempint - 500) & " es mascota de " & Mid$(Rdata, 9) & "[" & StringToLong(Rdata, 1) & "/" & StringToLong(Rdata, 5) & "].", 255, 1, 1, False, False)
                Else
                    'Call AddtoRichTextBox(FrmMain.RecTxt, Npcs(tempint - 500) & " [" & StringToLong(Left$(Rdata, 4), 1) & "/" & StringToLong(Rdata, 5) & "]", 255, 1, 1, False, False)
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.DescNpc
                'Como me lo paso no anda, entonces lo hago parecido? Marce
                TempByte = Asc(Left$(Rdata, 1))
                'Numero de Mensaje
                 tempint = STI(Rdata, 2)
                'Charindex
                TempStr = NpcsMensajes(TempByte)
               ' Mensaje

                Rdata = Right$(Rdata, Len(Rdata) - 3)
                'Sacamos el numero y el charindex
                If InStr(1, Rdata, ",") Then
                    Tempvar = Split(Rdata, ",")
                    For TempByte2 = 0 To UBound(Tempvar)
                    TempStr = Replace(TempStr, "#" & TempByte2 + 1, Tempvar(TempByte2))
                    Next
                End If
                'miramos el fonttype
                Call Dialogos.CreateDialog(TempStr, tempint, mzWhite)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.DescNpc2
            Dialogos.CreateDialog Right$(Rdata, Len(Rdata) - 2), STI(Rdata, 1), mzWhite
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BloquearTile
                MapData(Asc(Left$(Rdata, 1)), Asc(mid$(Rdata, 2, 1))).Blocked = Right$(Rdata, 1)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.pEnviarSpawnList
                For TempByte = 1 To val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(TempByte + 1, Rdata, 44)
                Next
                frmSpawnList.Show , frmMain
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ModCeguera
                UserCiego = Not UserCiego
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ModEstupidez
                UserEstupido = Not UserEstupido
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BorrarObj
                MapData(Asc(Left$(Rdata, 1)), Asc(Right$(Rdata, 1))).ObjGrh.GrhIndex = 0
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.CrearObjeto
                TempByte = Asc(mid$(Rdata, 3, 1))
                TempByte2 = Asc(Right$(Rdata, 1))
                If STI(Rdata, 1) = 669 Then
                    If MapData(TempByte, TempByte2).Particles_groups(1) = 0 Then _
                        Engine_Particles.Particle_Group_Make 1, TempByte, TempByte2, 6
'                    Exit Sub
                ElseIf STI(Rdata, 1) = 1521 Then
                    If MapData(TempByte, TempByte2).luz = 0 Then
                        MapData(TempByte, TempByte2).luz = Engine_Landscape.Light_Create(TempByte, TempByte2, 255, 200, 0, 3, 1, LUZ_TIPO_FUEGO)
                    End If
                End If
                MapData(TempByte, TempByte2).ObjGrh.GrhIndex = STI(Rdata, 1)

                InitGrh MapData(TempByte, TempByte2).ObjGrh, STI(Rdata, 1)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ApuntarProyectil
                UserStats(SlotStats).UsingSkill = proyectiles
                frmMain.MousePointer = 2
                Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ApuntarTrb
                UserStats(SlotStats).UsingSkill = Asc(Rdata)
                frmMain.MousePointer = 2
                    Select Case UserStats(SlotStats).UsingSkill
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
                    End Select
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarArmasConstruibles

                For TempByte = 0 To UBound(ArmasHerrero)
                    ArmasHerrero(TempByte) = 0
                Next TempByte
                If Len(Rdata) = 0 Then Exit Sub
                tempint = Len(Rdata) / 6 'Sacamos la cantidad de Armas;)
                TempStr = ""
                For TempByte = 0 To tempint - 1
                    ArmasHerrero(TempByte) = STI(Rdata, ((6 * TempByte)) + 5)
                    TempStr = Objeto(ArmasHerrero(TempByte)) & " (" & STI(Rdata, ((6 * TempByte) + 1)) & "/" & STI(Rdata, ((6 * TempByte)) + 3) & ")"
                    frmHerrero.lstArmas.AddItem TempStr
                Next TempByte
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarObjConstruibles
                For TempByte = 0 To UBound(ObjCarpintero)
                    ObjCarpintero(TempByte) = 0
                Next TempByte

                tempint = Len(Rdata) / 6
                If tempint = 0 Then Exit Sub
                For TempByte = 0 To tempint - 1
                    ObjCarpintero(TempByte) = STI(Rdata, ((6 * TempByte) + 5))
                    TempStr = Objeto(ObjCarpintero(TempByte)) & " (" & StringToLong(Rdata, ((6 * TempByte) + 1)) & " leños)"
                    frmCarp.lstArmas.AddItem TempStr
                Next TempByte
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarArmadurasConstruibles
               If frmHerrero.Visible = True Then Exit Sub
               frmHerrero.lstArmaduras.Clear


                For TempByte = 0 To UBound(ArmadurasHerrero)
                    ArmadurasHerrero(TempByte) = 0
                Next TempByte
                If Len(Rdata) = 0 Then Exit Sub
                tempint = Len(Rdata) / 6 'Sacamos la cantidad de armaduras;)
                TempStr = ""
                For TempByte = 0 To tempint - 1
                    ArmadurasHerrero(TempByte) = STI(Rdata, ((6 * TempByte)) + 5)
                    TempStr = Objeto(ArmadurasHerrero(TempByte)) & " (" & STI(Rdata, ((6 * TempByte) + 1)) & "/" & STI(Rdata, ((6 * TempByte)) + 3) & ")"
                    frmHerrero.lstArmaduras.AddItem TempStr
                Next TempByte

               Exit Sub





                For TempByte = 0 To UBound(ArmadurasHerrero)
                    ArmadurasHerrero(TempByte) = 0
                Next TempByte
                tempint = Len(Rdata) / 6
                If tempint < 2 Then Exit Sub
                For TempByte = 0 To tempint - 2
                     ArmadurasHerrero(TempByte) = STI(Rdata, ((6 * TempByte)) + 5)
                    TempStr = Objeto(ArmadurasHerrero(TempByte)) & " (" & STI(Rdata, ((6 * TempByte) + 1)) & "/" & STI(Rdata, ((6 * TempByte)) + 3) & ")"
                    frmHerrero.lstArmaduras.AddItem TempStr
                Next TempByte
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ShowCarp
            Call frmCarp.Show(vbModeless, frmMain)
            Exit Sub
             '---------------------------------------------
            Case sPaquetes.InitComUsu
                If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
                If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
                    For TempByte = 1 To UBound(UserInventory)
                        If UserInventory(TempByte).OBJIndex <> 0 Then
                            frmComerciarUsu.List1.AddItem UserInventory(TempByte).name
                            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = UserInventory(TempByte).Amount
                        Else
                            frmComerciarUsu.List1.AddItem "Nada"
                            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                        End If
                    Next TempByte
                    Comerciando = True
                    frmMain.Enabled = False
                    Call frmComerciarUsu.Show(vbModeless, frmMain)
                Exit Sub
            Case sPaquetes.ComUsuInv

                frmComerciarUsu.List2.Clear

                For tempint = 1 To Len(Rdata) / 19
                OtroInventario(tempint).OBJIndex = STI(Rdata, 1)
                OtroInventario(tempint).Valor = StringToLong(Rdata, 12)
                OtroInventario(tempint).Amount = StringToLong(Rdata, 16)
                OtroInventario(tempint).Equipped = 0
                OtroInventario(tempint).GrhIndex = STI(Rdata, 3)
                OtroInventario(tempint).OBJType = Asc(mid$(Rdata, 7, 1))
                OtroInventario(tempint).MaxHit = Asc(mid$(Rdata, 8, 1))
                OtroInventario(tempint).MinHit = Asc(mid$(Rdata, 9, 1))
                OtroInventario(tempint).name = Objeto(OtroInventario(tempint).OBJIndex)
                frmComerciarUsu.List2.AddItem OtroInventario(tempint).name
                frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(tempint).Amount
                Rdata = Right(Rdata, Len(Rdata) - 19)
                Next
                frmComerciarUsu.lblEstadoResp.Visible = False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.FinComUsuOk
                frmComerciarUsu.List1.Clear
                frmComerciarUsu.List2.Clear
                Unload frmComerciarUsu
                frmMain.Enabled = True
                frmMain.SetFocus
                Comerciando = False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.InitBanco
            '    DibujarBOv
               Call Dibujar(0, frmBancoObj.BovBoveda, UserBancoInventory, 8)
               Call Dibujar(DibujarInventario.itemelegido, frmBancoObj.Bovinventario, UserInventory, 6)
              '  DibujarBOvInventario
                Bovedeando = True
                frmMain.Enabled = False
                Call frmBancoObj.Show(vbModeless, frmMain)
                Exit Sub
               ' Do While TempByte <= UBound(UserInventory)
               ' If UserInventory(TempByte).OBJIndex <> 0 Then
                      '  Call DrawGrhtoHdc(frmBancoObj.Picture1(24).hwnd, frmBancoObj.Picture1(24).Hdc + 1, UserBancoInventory(TempByte).GrhIndex, SR, DR)
                '        frmBancoObj.List1(1).AddItem UserInventory(TempByte).Name
                'Else
                 '       frmBancoObj.List1(1).AddItem "Nada"
               ' End If
                'TempByte = TempByte + 1
                'Loop


            'TempByte = 1
           ' Do While TempByte <= UBound(UserBancoInventory)
                'If UserBancoInventory(TempByte).OBJIndex <> 0 Then
                '        frmBancoObj.List1(0).AddItem UserBancoInventory(TempByte).Name
               ' Else
               '         frmBancoObj.List1(0).AddItem "Nada"
              '  End If
             '   TempByte = TempByte + 1
           ' Loop
           ' Comerciando = True
          '  frmBancoObj.Show
            'Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarBancoObj
                Call RecivirBancoObj(Rdata)
                 If frmBancoObj.Visible Then Call Dibujar(CInt(frmBancoObj.BBitemElegido), frmBancoObj.BovBoveda, UserBancoInventory, 8)
            Exit Sub
            Case sPaquetes.BancoOk
                If frmBancoObj.Visible Then
                'TempByte = 1
               ' Do While TempByte <= UBound(UserInventory)
                '    If UserInventory(TempByte).OBJIndex <> 0 Then
                 '           frmBancoObj.List1(1).AddItem UserInventory(TempByte).Name
                  '  Else
                   '         frmBancoObj.List1(1).AddItem "Nada"
                   ' End If
                   ' TempByte = TempByte + 1
                'Loop

                'TempByte = 1
                'Do While TempByte <= UBound(UserBancoInventory)
                '    If UserBancoInventory(TempByte).OBJIndex <> 0 Then
                 '           frmBancoObj.List1(0).AddItem UserBancoInventory(TempByte).Name
                  '  Else
                   '         frmBancoObj.List1(0).AddItem "Nada"
                    'End If
                    'TempByte = TempByte + 1
                'Loop

                'If Rdata = "0" Then
                 '       frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                'Else
                 '       frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
               ' End If
               Call Dibujar(frmBancoObj.BBitemElegido, frmBancoObj.BovBoveda, UserBancoInventory, 8)
               Call Dibujar(frmBancoObj.BintemElegido, frmBancoObj.Bovinventario, UserInventory, 6)
            End If
            Exit Sub
            'HORRIBLEMENTE HECHO NO CAMBIE NADA; MODIFICAR ESTO EN UNA FEATURE REALEASE
            Case sPaquetes.PeaceSolRequest
                Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
            Case sPaquetes.EnviarPeaceProp
                Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
            Case sPaquetes.PeticionClan
                Call frmUserRequest.recievePeticion(Rdata)
                Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
            Case sPaquetes.EnviarCharInfo
                 Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
            Case sPaquetes.EnviarLeaderInfo
                Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
            Case sPaquetes.EnviarGuildsList
                Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
            Case sPaquetes.EnviarGuildNews
                Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
            Case sPaquetes.EnviarGuildDetails
                Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
            '/////////////////////////FEO//////////////////
            Case sPaquetes.HechizoFX 'Grafico y Sonido
                tempint = STI(Rdata, 1)
                Call SetCharacterFx(tempint, StringToByte(Rdata, 3), STI(Rdata, 4))
                If SoundActivated = 1 And Len(Rdata) > 5 Then
                   Call Audio.Sound_Play(Asc(Right(Rdata, 1)))
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeTalk
                AddtoRichTextBox frmMain.RecTxt, Rdata, 255, 255, 255, True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeSpell
                AddtoRichTextBox frmMain.RecTxt, Rdata, 130, 150, 200, True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeFight
                AddtoRichTextBox frmMain.RecTxt, Rdata, 255, 0, 0, True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeInfo
                AddtoRichTextBox frmMain.RecTxt, Rdata, 65, 190, 156, False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.CambiarHechizo
                TempByte = Asc(Left$(Rdata, 1))
                If Len(Rdata) = 1 Then Rdata = Rdata + "  (Vacio)"
                If UserHechizos(TempByte) = 255 Then UserHechizos(TempByte) = 0
                If TempByte > frmMain.hlst.ListCount Then
                    frmMain.hlst.AddItem mid$(Rdata, 3)
                Else
                    frmMain.hlst.List(TempByte - 1) = mid$(Rdata, 3)
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.pCrearNPC
            'POMBUG
If Len(Rdata) = 10 Then
                Call MakeChar(STI(Rdata, 1), STI(Rdata, 3), STI(Rdata, 5), StringToByte(Rdata, 7), StringToByte(Rdata, 8), StringToByte(Rdata, 9), 0, 0, 0)
    Debug.Print "MAKENPC OK?"
Else
    LogError "EL PUTO ERROR DE pCrearNPC, DATA:[" & Rdata & "]"
    Debug.Print "MAKENPC FALLADO"
End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ChangeNpc
                tempint = STI(Rdata, 1)
                CharList(tempint).Body = BodyData(STI(Rdata, 3))
                CharList(tempint).Head = HeadData(STI(Rdata, 5))
                CharList(tempint).Heading = StringToByte(Rdata, 7)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BorrarNpc
                Call EraseChar(STI(Rdata, 1))
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MoveChar
            tempint = STI(Rdata, 1)
            'If Fx = 1 And Not CharList(tempint).iBody = 8 Then
             '   Call DoPasosFx(tempint) 'REDUNDATE
            'End If
            Call Engine_Extend.Char_Move_by_Pos(tempint, Asc(mid(Rdata, 3)), Asc(mid(Rdata, 4)))
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarNpclst
                TempByte2 = 1
            For TempByte = 1 To val(Left$(Rdata, 1))
                frmEntrenador.lstCriaturas.AddItem ReadField(TempByte + 1, Rdata, 44)
            Next TempByte
            frmEntrenador.Show , frmMain
            Exit Sub
            '||||||||||||||||COMBATE|||||||||||||||||||
            Case sPaquetes.COMBRechEsc
                Call AddtoRichTextBox(frmMain.RecTxt, "Has rechazado el ataque con el escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBNpcHIT
                TempByte = Asc(Left$(Rdata, 1))
                Select Case TempByte
                    Case bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado en la cabeza por " & DeCodify(Right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado el brazo izquierdo por " & DeCodify(Right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado el brazo derecho por " & DeCodify(Right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado la pierna izquierda por " & DeCodify(Right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado la pierna derecha por " & DeCodify(Right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado en el torso por " & DeCodify(Right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    End Select
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBMuereUser
                Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha matado!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBNpcFalla
                Call AddtoRichTextBox(frmMain.RecTxt, "La criatura fallo el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBUserFalla
                Call AddtoRichTextBox(frmMain.RecTxt, "Has fallado el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBEnemEscu
                Call AddtoRichTextBox(frmMain.RecTxt, "El usuario rechazo el ataque con su escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.SangraUser
                tempint = STI(Rdata, 1)
                'CharList(tempint).fx = 14
                'CharList(tempint).FxLoopTimes = 0
                Crear_Sangre tempint, 30, 8000, 32
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBUserImpcNpc
                Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a la criatura por " & DeCodify(Rdata) & "!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBEnemFalla
                Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Rdata & " te ataco y fallo!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBEnemHitUs ' <<--- user nos impacto
                TempByte = Asc(Left$(Rdata, 1))
                Select Case TempByte
                    Case bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Right$(Rdata, Len(Rdata) - 3) & " te ha pegado en la cabeza por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Right$(Rdata, Len(Rdata) - 3) & " te ha pegado el brazo izquierdo por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Right$(Rdata, Len(Rdata) - 3) & " te ha pegado el brazo derecho por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Right$(Rdata, Len(Rdata) - 3) & " te ha pegado la pierna izquierda por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Right$(Rdata, Len(Rdata) - 3) & " te ha pegado la pierna derecha por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Right$(Rdata, Len(Rdata) - 3) & " te ha pegado en el torso por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                End Select
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBUserHITUser ' <<--- impactamos un user
                TempByte = Asc(Left(Rdata, 1))

                Select Case TempByte
                    Case bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & Right$(Rdata, Len(Rdata) - 3) & " en la cabeza por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & Right$(Rdata, Len(Rdata) - 3) & " en el brazo izquierdo por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & Right$(Rdata, Len(Rdata) - 3) & " en el brazo derecho por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & Right$(Rdata, Len(Rdata) - 3) & " en la pierna izquierda por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & Right$(Rdata, Len(Rdata) - 3) & " en la pierna derecha por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & Right$(Rdata, Len(Rdata) - 3) & " en el torso por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                End Select
            Exit Sub
            '~~~~~~~~~~~~~~~~~~~~Combate~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '////////////////////////Trabajo///////////////////////////////////
            Case sPaquetes.Navega
                UserNavegando = Not UserNavegando
            Exit Sub
            '~~~~~~~~~~~~~~~~~~~~ Me canse¬¬
            Case sPaquetes.AuraFx
                SetCharacterFx STI(Rdata, 1), val(DeCodify(Right$(Rdata, Len(Rdata) - 2))), 999
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.Meditando
                UserMeditar = Not UserMeditar
                If UserMeditar Then
                AddtoRichTextBox frmMain.RecTxt, "Empiezas a meditar.", 65, 190, 156, False, False, False
                Else
                    If UserStats(SlotStats).UserMinMAN = UserMaxMAN Then
                    AddtoRichTextBox frmMain.RecTxt, "Has terminado de meditar.", 65, 190, 156, False, False, False
                    Else
                    AddtoRichTextBox frmMain.RecTxt, "Dejas de meditar.", 65, 190, 156, False, False, False
                    End If
                End If
                Exit Sub
            '---------------------------------------------
            Case sPaquetes.NoParalizado
                'CharList(STI(Rdata, 1)).Paralized = False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.Paralizado2
                UserStats(SlotStats).UserParalizado = True
                CharMap(UserPos.X, UserPos.Y) = 0
                UserPos.X = Asc(Left$(Rdata, 1))
                UserPos.Y = Asc(Right$(Rdata, 1))
                CharMap(UserPos.X, UserPos.Y) = UserCharIndex
                rm2a
               ' CharList(UserCharIndex).Paralized = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.NoParalizado2
                UserStats(SlotStats).UserParalizado = False
                'CharList(UserCharIndex).Paralized = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.invisible
                'El sig char esta invi!
                tempint = STI(Rdata, 1)
                 CharList(tempint).invisible = True
                 CharList(tempint).AlphaVal = 0
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.Visible
                tempint = STI(Rdata, 1)
                CharList(tempint).invisible = False
             '  CharList(tempint).InvisibleGuild = 0
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.pChangeUserChar
                tempint = STI(Rdata, 1)
                CharList(tempint).iBody = STI(Rdata, 3)
                CharList(tempint).Body = BodyData(STI(Rdata, 3))
                CharList(tempint).Head = HeadData(STI(Rdata, 5))
                CharList(tempint).Heading = StringToByte(Rdata, 7)
                If StringToByte(Rdata, 8) > 0 Then
                    CharList(tempint).Arma = WeaponAnimData(StringToByte(Rdata, 8))
                    CharList(tempint).Escudo = ShieldAnimData(StringToByte(Rdata, 9))
                End If

                CharList(tempint).Casco = CascoAnimData(StringToByte(Rdata, 13))

                SetCharacterFx tempint, StringToByte(Rdata, 10), STI(Rdata, 11)
             Exit Sub
            '---------------------------------------------
            Case sPaquetes.LevelUP
                SkillPoints = SkillPoints + STI(Rdata, 1)
                frmMain.Label1.Visible = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.SendSkills
                For TempByte = 1 To NUMSKILLS
                    UserSkills(TempByte) = StringToByte(Rdata, TempByte)
                    'In this way, evitamos enviar 48 caracteres pudiendo
                    'enviar 24.
                 Next TempByte
                LlegaronSkills = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.SendFama
                UserReputacion.AsesinoRep = StringToLong(Rdata, 1)
                UserReputacion.BandidoRep = StringToLong(Rdata, 5)
                UserReputacion.BurguesRep = StringToLong(Rdata, 9)
                UserReputacion.LadronesRep = StringToLong(Rdata, 13)
                UserReputacion.NobleRep = StringToLong(Rdata, 17)
                UserReputacion.PlebeRep = StringToLong(Rdata, 21)
                UserReputacion.promedio = ((-UserReputacion.AsesinoRep) + _
                                          (-UserReputacion.BandidoRep) + _
                                          UserReputacion.NobleRep + _
                                          UserReputacion.BurguesRep + _
                                          (-UserReputacion.LadronesRep) + _
                                          UserReputacion.PlebeRep) / 6
                LlegoFama = True
            Exit Sub
                '---------------------------------------------
            Case sPaquetes.SendAtributos
                For TempByte = 1 To NUMATRIBUTOS
                    UserAtributos(TempByte) = Asc(mid$(Rdata, TempByte, 1))
                Next TempByte
                  LlegaronAtrib = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MiniEst
               ' UserMiniEst.CiudasMuertos = STI(Rdata, 1)
                'UserMiniEst.CrimisMuertos = STI(Rdata, 3)
                'UserMiniEst.UsersMuertos = STI(Rdata, 5)
                'UserMiniEst.TiempoCarcel = Asc(Mid$(Rdata, 7, 1))
                'UserMiniEst.NpcsMuertos = STI(Rdata, 8)
                'UserMiniEst.Clase = Mid$(Rdata, 10)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BorrarUser
                Rdata = STI(Rdata, 1)
                Call EraseChar(val(Rdata))
                Call Dialogos.RemoveDialog(val(Rdata))
                Call RefreshAllChars

            Exit Sub
            '---------------------------------------------
            Case sPaquetes.CrearChar
                tempint = STI(Rdata, 1)
                SetCharacterFx tempint, StringToByte(Rdata, 3), 1 '999
                CharList(tempint).Nombre = mid$(Rdata, 16)
                TempByte = StringToByte(Rdata, 14)
                CharList(tempint).criminal = TempByte
                CharList(tempint).priv = StringToByte(Rdata, 15)
                Call MakeChar(tempint, STI(Rdata, 4), STI(Rdata, 6), StringToByte(Rdata, 8), Asc(mid$(Rdata, 9, 1)), Asc(mid$(Rdata, 10, 1)), StringToByte(Rdata, 11), StringToByte(Rdata, 12), StringToByte(Rdata, 13))

            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarPos
                'MapData(UserPos.X, UserPos.Y).CharIndex = 0
                'UserPos.X = Asc(left$(Rdata, 1))
                'UserPos.Y = Asc(right$(Rdata, 1))
                'MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
                'CharList(UserCharIndex).Pos = UserPos
                'CharMap(UserPos.x, UserPos.y) = 0
                UserPos.X = Asc(Left$(Rdata, 1))
                UserPos.Y = Asc(Right$(Rdata, 1))
                If UserCharIndex > 0 Then
                CharMap(UserPos.X, UserPos.Y) = UserCharIndex
                CharList(UserCharIndex).Pos = UserPos
                End If
             Exit Sub
            '---------------------------------------------
            Case sPaquetes.InvRefresh
                Call RecivirInvRefresh(Rdata)
                If Bovedeando Then Call Dibujar(CInt(frmBancoObj.BintemElegido), frmBancoObj.Bovinventario, UserInventory, 6)
                If Comerciando Then Call Dibujar(frmComerciar.ItemElegidoV, frmComerciar.ComercioInventario, UserInventory, 6)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarStat
            UserMaxHP = STI(Rdata, 1)
            UserStats(SlotStats).UserMinHP = STI(Rdata, 3)
            UserMaxMAN = STI(Rdata, 5)
            UserStats(SlotStats).UserMinMAN = STI(Rdata, 7)
            UserMaxSTA = STI(Rdata, 9)
            UserStats(SlotStats).UserMinSTA = STI(Rdata, 11)
            UserGLD = StringToLong(Rdata, 13)
            UserLvl = Asc(mid(Rdata, 17, 1))
            UserPasarNivel = StringToLong(Rdata, 18) * 100 + StringToByte(Rdata, 22)
            UserExp = StringToLong(Rdata, 23) * 100 + StringToByte(Rdata, 27)
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.Hpshp.Width = (((UserStats(SlotStats).UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            frmMain.Label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
            frmMain.Label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
            frmMain.Label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
            frmMain.Label16.Caption = UserMinHAM & "/" & UserMaxHAM
            frmMain.Label17.Caption = UserMinAGU & "/" & UserMaxAGU
            If UserMaxMAN > 0 Then
                frmMain.ManShp.Width = (((UserStats(SlotStats).UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.ManShp.Width = 0
            End If
            frmMain.stashp.Width = (((UserStats(SlotStats).UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl


      '  If frmMain.Hpshp.Width > 68 Then
           ' frmMain.label15.ForeColor = &H0&
       ' ElseIf frmMain.Hpshp.Width > 59 And frmMain.Hpshp.Width <= 68 Then
           ' frmMain.label15.ForeColor = &H404040
       ' ElseIf frmMain.Hpshp.Width > 36 And frmMain.Hpshp.Width <= 59 Then
           ' frmMain.label15.ForeColor = &HC0C0C0
       ' ElseIf frmMain.Hpshp.Width > 25 And frmMain.Hpshp.Width <= 36 Then
           ' frmMain.label15.ForeColor = &HE0E0E0
       ' ElseIf frmMain.Hpshp.Width < 25 Then
          '  frmMain.label15.ForeColor = &HFFFFFF
        'End If
'Para la barra de stamina

            If UserStats(SlotStats).UserMinHP <= 0 Then
                UserStats(SlotStats).UserEstado = 1
                UserStats(SlotStats).UserParalizado = False
                UserDescansar = False
                UserMeditar = False
                UserEstupido = False
                IsEnvenenado = False
                If val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = True
            Else
                UserStats(SlotStats).UserEstado = 0
                If val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = False
            End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarF
                'UserFuerza = Asc(Rdata)
               ' frmMain.LblFuerza.Caption = UserFuerza
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarA
               ' UserAgilidad = Asc(Rdata)
               'frmMain.LblAgilidad.Caption = UserAgilidad
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarOro
                UserGLD = DeCodify(Rdata)
                frmMain.GldLbl.Caption = UserGLD
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarHP
                UserStats(SlotStats).UserMinHP = DeCodify(Rdata)
                frmMain.Label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
                frmMain.Hpshp.Width = (((UserStats(SlotStats).UserMinHP / 100) / (UserMaxHP / 100)) * 94)

                'If frmMain.Hpshp.Width > 68 Then
                   ' frmMain.label15.ForeColor = &H0&
               ' ElseIf frmMain.Hpshp.Width > 59 And frmMain.Hpshp.Width <= 68 Then
                    'frmMain.label15.ForeColor = &H404040
               ' ElseIf frmMain.Hpshp.Width > 36 And frmMain.Hpshp.Width <= 59 Then
                   ' frmMain.label15.ForeColor = &HC0C0C0
               ' ElseIf frmMain.Hpshp.Width > 25 And frmMain.Hpshp.Width <= 36 Then
                   ' frmMain.label15.ForeColor = &HE0E0E0
               ' ElseIf frmMain.Hpshp.Width < 25 Then
                   ' frmMain.label15.ForeColor = &HFFFFFF
              '  End If

                If UserStats(SlotStats).UserMinHP <= 0 Then
                UserStats(SlotStats).UserEstado = 1
                UserEstupido = False
                UserStats(SlotStats).UserParalizado = False
                UserDescansar = False
                UserMeditar = False
                IsEnvenenado = False
               ' If Val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = True
                Else
                UserStats(SlotStats).UserEstado = 0
                'MsgBox "hola"
                If val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = False
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarMP
                UserStats(SlotStats).UserMinMAN = DeCodify(Rdata)
                frmMain.Label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN

                If UserMaxMAN > 0 Then
                    frmMain.ManShp.Width = (((UserStats(SlotStats).UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
                Else
                    frmMain.ManShp.Width = 0
                End If

            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarST
                UserStats(SlotStats).UserMinSTA = DeCodify(Rdata)
                frmMain.Label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                frmMain.stashp.Width = (((UserStats(SlotStats).UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)

            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarEXP
                UserExp = DeCodify(Rdata)
                frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
               ' tempstr = UserExp
               ' tempstr2 = UserPasarNivel
              '  If Len(tempstr) > 7 Then
               '     tempstr = left$(tempstr, Len(tempstr) - 4)
              '      tempstr2 = left$(tempstr2, Len(tempstr2) - 4)
               '' End If
                'TempLong = Round((Val(tempstr) * 100) / Val(tempstr2), 0)
                'frmMain.LblExpPorc.Caption = TempLong & "%"
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarSYM
                UserStats(SlotStats).UserMinSTA = STI(Rdata, 1)
                UserStats(SlotStats).UserMinMAN = STI(Rdata, 3)
                If UserStats(SlotStats).UserMinSTA > 0 Then
                    frmMain.stashp.Width = (((UserStats(SlotStats).UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
                Else
                    frmMain.stashp.Width = 0
                End If
                If UserStats(SlotStats).UserMinMAN > 0 Then
                  '  frmMain.MpShp.Width = (((UserStats(SlotStats).UserStats(SlotStats).UserMinMAN / 100) / (UserMaxMAN / 100)) * 93)
                Else
                  '  frmMain.MpShp.Width = 0
                End If

                'frmMain.LblSp.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
               ' frmMain.LblMp.Caption = UserStats(SlotStats).UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
                'frmMain.LblSp2.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                'frmMain.LblMp2.Caption = UserStats(SlotStats).UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarSYH
                UserStats(SlotStats).UserMinSTA = STI(Rdata, 1)
                UserStats(SlotStats).UserMinHP = STI(Rdata, 3)
                If UserStats(SlotStats).UserMinSTA > 0 Then
                    frmMain.stashp.Width = (((UserStats(SlotStats).UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
                Else
                    frmMain.stashp.Width = 0
                End If
                If UserStats(SlotStats).UserMinHP > 0 Then
                    frmMain.Hpshp.Width = (((UserStats(SlotStats).UserMinHP / 100) / (UserMaxHP / 100)) * 93)
                Else
                    frmMain.Hpshp.Width = 0
                End If
                If UserStats(SlotStats).UserMinHP <= 0 Then
                    UserStats(SlotStats).UserEstado = 1
                    UserStats(SlotStats).UserParalizado = False
                    UserDescansar = False
                    UserMeditar = False
                    UserEstupido = False
                    IsEnvenenado = False
                    'If Val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = True
                    Else
                    UserStats(SlotStats).UserEstado = 0
                    'If Val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = False
                    End If
               ' frmMain.LblSp.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                'frmMain.LblHp.Caption = UserStats(SlotStats).UserMinHp & "/" & UserMaxHP
              ' frmMain.LblSp2.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                'frmMain.LblHp2.Caption = UserStats(SlotStats).UserMinHp & "/" & UserMaxHP
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarFA
                'UserFuerza = Asc(left$(Rdata, 1))
                'UserAgilidad = Asc(right$(Rdata, 2))
               ' frmMain.LblAgilidad = UserAgilidad
              '  frmMain.LblFuerza = UserFuerza
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarHYS
                UserMaxAGU = 100
                UserMinAGU = StringToByte(Rdata, 1)
                UserMaxHAM = 100
                UserMinHAM = StringToByte(Rdata, 2)
                frmMain.Aguasp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
                frmMain.comidasp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
                frmMain.Label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                frmMain.Label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
                frmMain.Label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
                frmMain.Label16.Caption = UserMinHAM & "/" & UserMaxHAM
                frmMain.Label17.Caption = UserMinAGU & "/" & UserMaxAGU

                Exit Sub
            '---------------------------------------------
            Case sPaquetes.QDL
                Call Dialogos.RemoveDialog(STI(Rdata, 1))
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MDescansar
                UserDescansar = Not UserDescansar
                If UserDescansar Then AddtoRichTextBox frmMain.RecTxt, "Te acomodas junto a la fogata y comienzas a descansar.", 65, 190, 156, False, False, False
                If Not UserDescansar Then AddtoRichTextBox frmMain.RecTxt, "Has dejado de descansar.", 65, 190, 156, False, False, False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ChangeMap
                UserMap = Asc(Left$(Rdata, 1))
                frmMain.Coord.Caption = mapa(UserMap)
                Rdata = Right$(Rdata, Len(Rdata) - 3)
                Terreno = ReadField(1, Rdata, 44)
                Zona = ReadField(2, Rdata, 44)
                'Si es la vers correcta cambiamos el mapa
                AddtoRichTextBox frmMain.RecTxt, "CAMBIO DE MAPA!! -> " & UserMap, 255, 190, 156
                Call SwitchMap(UserMap)

                If bLluvia(UserMap) = 0 Then
                    If bRain Then
                        Cambiar_estado_climatico Clima_Normal
                        IMC.Stop
                    End If
                Else
                     If bRain Then
                        Cambiar_estado_climatico Clima_Lluvia_normal
                        If Not IMPos.CurrentPosition > 0 Then
                            IMC.Run
                        End If
                     End If
                End If

            Exit Sub

            '---------------------------------------------
            Case sPaquetes.ChangeMusic
                    If StringToByte(Rdata, 1) <> 0 Then
                        If Musica = 1 Then
                            CurMidi = StringToByte(Rdata, 1)
                            'FIXME Call Audio.PlayMIDI(CurMidi, -1)
                        End If
                    End If
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.QTDL
                Call Dialogos.RemoveAllDialogs
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.IndiceChar
                UserCharIndex = STI(Rdata, 1)
                UserPos = CharList(UserCharIndex).Pos
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.mBox
                NObajar = True

                If InStr(1, Msgboxes(Asc(Left(Rdata, 1))), "#") > 0 Then
                    frmConnect.texto = Replace(Msgboxes(Asc(Left(Rdata, 1))), "#", Right(Rdata, Len(Rdata) - 1))
                Else
                    frmConnect.texto = Msgboxes(Asc(Rdata))
                End If

                If EstadoLogin = crearnuevopj Then
                    frmCrearPersonaje.flash.SetVariable "formulario.errorm", Msgboxes(Asc(Rdata))
                Else
                    If frmOldPersonaje.Visible Then frmOldPersonaje.Hide

                    frmConnect.msgbox2.Visible = True
                    frmConnect.texto.Visible = True
                    frmConnect.Image1(2).Enabled = True
                End If
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.EnviarUI
                UserIndex = STI(Rdata, 1)
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.Loguea
                UserPrivilegios = StringToByte(Rdata, 1)
                Intervalos mid(Rdata, 2)

                UserCiego = False
                EngineRun = True
                IScombate = False
                UserDescansar = False
                Nombres = True
                UserEstupido = False
                If EstadoLogin = crearnuevopj Then
                'If frmCrearPersonaje.Visible Then
                 '  Unload frmPasswdSinPadrinos
                   Unload frmCrearPersonaje
                   'frmCrearPersonaje.Hide
                   Unload frmConnect
                   ShowCursor (True)
                   frmMain.Show
                Else
                Unload frmOldPersonaje
                End If

                Call SetConnected

                bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 8 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
                Call DoFogataFx
                Call SetMusicInfo("Jugando Tierras del Sur: " & UserName, "", "")
                Call frmMain.picInv.Refresh
            Exit Sub
            Case sPaquetes.Lluvia
                If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
                bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 8 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
'                If Not bRain Then
'                    bRain = True
'                    If bLluvia(UserMap) <> 0 Then
'                    Play_Song ("")
'                    End If
'                    Cambiar_estado_climatico Clima_Lluvia_normal
'                Else
'                    bRain = False
'                    Cambiar_estado_climatico Clima_Normal
'                End If
                Debug.Print "paquete de lluvia!"; "argumento:"; StringToByte(Rdata, 1)
                Cambiar_estado_climatico StringToByte(Rdata, 1)
            Exit Sub
            '...............................................
            Case sPaquetes.SOSAddItem
                frmMSG.List1.AddItem Rdata
            Exit Sub
            '...............................................
            Case sPaquetes.SOSViewList
                frmMSG.Caption = "Denuncias"
                frmMSG.Label1 = "Usuarios"
                frmMSG.Visible = True
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeServer
                AddtoRichTextBox frmMain.RecTxt, "Servidor> " & Rdata, 0, 185, 0, False, False
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeGMSG
                AddtoRichTextBox frmMain.RecTxt, Rdata, 0, 255, 0, False, True
            Exit Sub
            '...............................................
            Case sPaquetes.UserTalk
                Dialogos.CreateDialog Left$(Rdata, Len(Rdata) - 2), STI(Right$(Rdata, 2), 1), RGB(255, 255, 255)
            Exit Sub
            '...............................................
            Case sPaquetes.UserShout
                Dialogos.CreateDialog Left$(Rdata, Len(Rdata) - 2), STI(Right$(Rdata, 2), 1), RGB(255, 0, 0)
            Exit Sub
            '...............................................
            Case sPaquetes.UserWhisper
                Dialogos.CreateDialog Left$(Rdata, Len(Rdata) - 2), STI(Right$(Rdata, 2), 1), vbYellow
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToNorth
                tempint = STI(Rdata, 1)
                CharList(tempint).Heading = NORTH
                Call PararPj(tempint)
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToSouth
                tempint = STI(Rdata, 1)
                CharList(tempint).Heading = SOUTH
                Call PararPj(tempint)
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToEast
                tempint = STI(Rdata, 1)
                CharList(tempint).Heading = EAST
                Call PararPj(tempint)
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToWest
                tempint = STI(Rdata, 1)
                CharList(tempint).Heading = WEST
                Call PararPj(tempint)
            Exit Sub
            '...............................................
            Case sPaquetes.FinComOk
         '       frmComerciar.List1(0).Clear
           '     frmComerciar.List1(1).Clear
                NPCInvDim = 0
                Unload frmComerciar
                Comerciando = False
            Exit Sub
            '...............................................
            Case sPaquetes.FinBanOk
                'frmBancoObj.List1(0).Clear
               ' frmBancoObj.List1(1).Clear
                NPCInvDim = 0
                Unload frmBancoObj
                Bovedeando = False
                frmMain.Enabled = True
                frmMain.SetFocus
            Exit Sub
            '...............................................
            Case sPaquetes.SndDados
            With frmCrearPersonaje.flash
                If .Visible Then
                    .SetVariable "d1", Asc(mid$(Rdata, 1, 1))
                    .SetVariable "d2", Asc(mid$(Rdata, 2, 1))
                    .SetVariable "d3", Asc(mid$(Rdata, 3, 1))
                    .SetVariable "d4", Asc(mid$(Rdata, 4, 1))
                    .SetVariable "d5", Asc(mid$(Rdata, 5, 1))
                End If
            End With
            Exit Sub
            '...............................................
            Case sPaquetes.ShowHerreriaForm

                frmHerrero.Show vbModal, frmMain
            Exit Sub
            '...............................................
            Case sPaquetes.InitGuildFundation
                CreandoClan = True
                Call frmGuildFoundation.Show(vbModeless, frmMain)
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeClan1
            'Esto lo hize yo hace mucho y es feisimo, funciona. Pero esta mal echo. Algun dia
            'Lo tengo que hacer lindo 'Marce  6/3/2006
                    If Not activado Then
                    AddtoRichTextBox frmMain.RecTxt, Rdata, 228, 199, 27, 0, 0, False
                    Else
                        If Len(Rdata) < 90 Then
                        clantext5 = clantext4
                        clantext4 = clantext3
                        clantext3 = clantext2
                        clantext2 = clantext1
                        clantext1 = Rdata
                        Else
                        clantext5 = clantext3
                        clantext4 = clantext2
                        clantext3 = clantext1
                        clantext2 = mid(Rdata, 1, 87) & "-"
                        clantext1 = mid(Rdata, 88, Len(Rdata))
                        End If
                    End If
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeClan2
                AddtoRichTextBox frmMain.RecTxt, Rdata, 150, 50, 150
            Exit Sub
            '...............................................
            Case sPaquetes.SaidMagicWords
                Dialogos.CreateDialog mid$(Rdata, 3), STI(Rdata, 1), D3DColorXRGB(0, 192, 185), True
            Exit Sub
            '...............................................
            Case sPaquetes.MoveNpc
                    tempint = STI(Rdata, 1)
                If fx = 1 Then
                    DoPasosFx (tempint)
                End If
                Call Char_Move_by_Pos(tempint, Asc(mid(Rdata, 3)), Asc(mid(Rdata, 4)))
              Exit Sub
            '...............................................
            Case sPaquetes.pEnviarNpcInvBySlot
                TempByte = StringToByte(Rdata, 1)
                If Len(Rdata) > 1 Then
                    NPCInventory(TempByte).OBJType = Asc(Left$(Rdata, 2))
                    NPCInventory(TempByte).Amount = STI(Rdata, 3)
                    NPCInventory(TempByte).GrhIndex = STI(Rdata, 5)
                    NPCInventory(TempByte).OBJIndex = STI(Rdata, 7)
                    NPCInventory(TempByte).MaxHit = STI(Rdata, 9)
                    NPCInventory(TempByte).MinHit = STI(Rdata, 11)
                    NPCInventory(TempByte).MinDef = StringToByte(Rdata, 13)
                    NPCInventory(TempByte).MaxDef = StringToByte(Rdata, 14)
                    NPCInventory(TempByte).Valor = DeCodify(mid$(Rdata, 15))
                    NPCInventory(TempByte).name = Objeto(NPCInventory(TempByte).OBJIndex)
                Else
                    NPCInventory(TempByte).Amount = 0
                    NPCInventory(TempByte).GrhIndex = 0
                    NPCInventory(TempByte).OBJIndex = 0
                    NPCInventory(TempByte).OBJType = 0
                    NPCInventory(TempByte).MaxHit = 0
                    NPCInventory(TempByte).MinHit = 0
                    NPCInventory(TempByte).MaxDef = 0
                   NPCInventory(TempByte).MinDef = 0
                    NPCInventory(TempByte).Valor = 0
                    NPCInventory(TempByte).name = "(None)"
                End If
               ' For tempint = 1 To MAX_NPC_INVENTORY_SLOTS
                 '   frmComerciar.List1(0).AddItem NPCInventory(tempint).Name
                'Next tempint
            Exit Sub

            '...............................................
            Case sPaquetes.mTransError
                If frmConnect.Visible = True Then
                   ' MostrarTransCartel Rdata, vbRed
                End If
            Exit Sub
            '...............................................
            '...............................................
            Case sPaquetes.CrearObjetoInicio

                Dim Y As String
                Dim X As String
                Dim i As Integer
                Dim veces As Integer

                veces = STI(Left(Rdata, 2), 1)
                Rdata = Right(Rdata, Len(Rdata) - 2)
                For i = 0 To veces - 1
                X = Asc(mid(Rdata, 3 + (i * 4)))
                Y = Asc(mid(Rdata, 4 + (i * 4)))
                MapData(X, Y).ObjGrh.GrhIndex = STI2(Rdata, 1 + (i * 4))
                InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
                Next i
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.pMensajeSimple2 'Simple
                Rdata = (Asc(Rdata)) + 255
                Tempvar = Split(Mensaje(Rdata), "~")
                Rdata = Tempvar(0)
                Call AddtoRichTextBox(frmMain.RecTxt, Rdata, Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)), Int(Tempvar(5)))
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.noche
                Horas = StringToByte(Rdata, 1)
                Minutos = StringToByte(Rdata, 2)
                frmMain.tNoche.Enabled = True
                Exit Sub
            '...............................................
            '...............................................
            Case sPaquetes.SegOFF
                Call AddtoRichTextBox(frmMain.RecTxt, ">>SEGURO DESACTIVADO<<", 255, 0, 0, True, False, False)
                UserSeguro = False
                frmMain.IconoSeg = "X"
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.SegOn
                Call AddtoRichTextBox(frmMain.RecTxt, ">>SEGURO ACTIVADO<<", 0, 255, 0, True, False, False)
                UserSeguro = True
                frmMain.IconoSeg = ""
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.Nieva
                If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
                bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 8 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
'                If Not bSnow Then
'                    bSnow = True
'                Else
'                    If bNieva(UserMap) <> 0 Then
'                            'Call frmMain.StopSound
'                            'Call frmMain.Play("nieve.wav", False)
'                          '  frmMain.IsPlaying = plNone
'
'                    End If
'                    bSnow = False
'                End If
'                If bNieva(UserMap) Then
'                    Cambiar_estado_climatico Clima_Nieve
'                Else
'                    If bRain Then
'                        If bLluvia(UserMap) Then
'                            Cambiar_estado_climatico Clima_Lluvia_normal
'                        Else
'                            Cambiar_estado_climatico Clima_Normal
'                        End If
'                    Else
'                        Cambiar_estado_climatico Clima_Normal
'                    End If
'                End If
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.DejaDeTrabajar
                Call Mod_General.DejarDeTrabajars
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.TXA
                Engine_FX.FX_Hit_Create_Pos Asc(Left(Rdata, 1)), Asc(mid(Rdata, 2, 1)), STI(Rdata, 3), 3000, mzRed
                'Call AddTxtAtaque(Asc(Left(Rdata, 1)), Asc(mid(Rdata, 2, 1)), STI(Rdata, 3), STI(Rdata, 5))
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.mBox2
                'If PuedoQuitarFoco Then
                MsgBox Rdata, vbInformation, "Mensaje del servidor"
                'frmMensaje.Show
                'End If
                'Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.FXH
                tmpbyte = StringToByte(Rdata, 5)
'                Flecha comun = 2
'                Flecha incendiaria = 3
'                Flecha de tejo = 4
'                Magias >= 5

                Call AddFXHechizos(STI(Rdata, 1), STI(Rdata, 3), )
                
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.FundoParty
                Liderparty = True
                Partym.Boton(2).Enabled = True
                Partym.Boton(2).Visible = True
                Partym.Boton(1).Visible = False
                Partym.Boton(6).Enabled = False
                Partym.Boton(7).Enabled = False
                Partym.Label6.Visible = True
                Partym.List1.Visible = True
                Partym.List1.Enabled = True
                Partym.List2.Visible = True
                Partym.List2.Enabled = True
                Partym.Label4.Visible = False
                Partym.Label3.Visible = True
                Partym.Label2.Visible = True
                Partym.Boton(5).Enabled = True
                Call DameImagen(Partym.Boton(5), 29)
                Call DameImagen(Partym.Boton(2), 24)
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.PNI

                For i = 0 To 20
                If Listasolicitudes(i) = Rdata Then
                Listasolicitudes(i) = ""
                Exit For
                End If
                Next i

                For i = 0 To 20
                If Listaintegrantes(i) = "" Or Listaintegrantes(i) = Rdata Then
                Listaintegrantes(i) = Rdata
                Exit For
                End If
                Next
                Partym.List2.AddItem Rdata

                For i = 0 To Partym.List1.ListCount - 1
                If Partym.List1.List(i) = Rdata Then Partym.List1.RemoveItem i
                Next


            '...............................................
            '...............................................
                Case sPaquetes.Integranteparty
                gh = True
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.OnParty
                '0 ganas de programar asique lo hago asi nomas.. marce
                Dim informacion As String
                TempLong = 0
                For tempint = 1 To 5
                informacion = ReadField(tempint, Rdata, Asc(":"))
                    Partym.Label5(i).Caption = ReadField(3, informacion, Asc(";"))
                    Partym.Label7(i).Caption = ReadField(1, informacion, Asc(";"))
                    Partym.Label8(i).Caption = ReadField(2, informacion, Asc(";"))
                    TempLong = TempLong + val(Partym.Label7(i).Caption)
                    i = i + 1
                Next
                Partym.Label11.Caption = TempLong
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.Mest
                With UserEstadisticas
                    .ciudadanosMatados = STI(Rdata, 1)
                    .criminalesMatados = STI(Rdata, 3)
                    .UsuariosMatados = STI(Rdata, 5)
                    .NpcsMatados = StringToLong(Rdata, 7)
                    .Clase = Right(Rdata, Len(Rdata) - 11)
                    .PenaCarcel = StringToByte(Rdata, 11)
                End With
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.AnimGolpe
                    tempint = DeCodify(Rdata)
                    Char_Start_Anim tempint
                Exit Sub
            '...............................................
            '...............................................
            Case sPaquetes.AnimEscu
                tempint = STI(Rdata, 1)
                Char_Start_Anim_Escudo tempint
            Exit Sub
            '...............................................
            Case sPaquetes.CFXH
            Call AddFXList(STI(Rdata, 1), StringToByte(Rdata, 3), STI(Rdata, 4), Asc(mid(Rdata, 6, 1)))
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeGuild
            AddtoRichTextBox frmMain.RecTxt, Rdata, 255, 255, 255, True
            Audio.Sound_Play 43
            Exit Sub
            '...............................................
            Case sPaquetes.ClickObjeto
            If Len(Rdata) > 2 Then
            AddtoRichTextBox frmMain.RecTxt, Objeto(STI(Rdata, 1)) & " (" & STI(Rdata, 3) & ")", 65, 190, 156, False
            Else
            AddtoRichTextBox frmMain.RecTxt, Objeto(STI(Rdata, 1)), 65, 190, 156, False
            End If
            Exit Sub
            '...............................................
            Case sPaquetes.LISTUSU
            Tempvar = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For tempint = LBound(Tempvar) To UBound(Tempvar)
                    frmPanelGm.cboListaUsus.AddItem Tempvar(tempint)
                Next tempint
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
            End If
            Exit Sub
            '...............................................
            Case sPaquetes.Traba
            If LenB(Rdata) < 3 Then
            frmMSG.Caption = "Trabajando"
            frmMSG.Label1 = "Usuarios"
            frmMSG.Show , frmMain
            Else
            frmMSG.List1.AddItem Rdata
            End If
            Exit Sub
            '...............................................
            Case sPaquetes.UserTalkDead
            Dialogos.CreateDialog Left$(Rdata, Len(Rdata) - 2), STI(Right$(Rdata, 2), 1), D3DColorXRGB(120, 120, 120)
            Exit Sub
            '...............................................
            Case sPaquetes.TiempoRetos
            TempByte = StringToByte(Rdata, 1)
            If TempByte > 0 Then
                AddtoRichTextBox frmMain.RecTxt, "Reto> " & TempByte, 250, 250, 200, False
                TiempoReto = 1
            Else
                TiempoReto = 0
                AddtoRichTextBox frmMain.RecTxt, "Reto> " & "YA!", 220, 220, 220, False
            End If
            Exit Sub
           '...............................................
            Case sPaquetes.Pang
            AddtoRichTextBox frmMain.RecTxt, "Tiempo de retardo: " & Int(PingPerformanceTimer.Time) & " ms", 65, 190, 156, False
            'AddtoRichTextBox frmMain.RecTxt, "Tiempo de retardo: " & GetTickCount - PingTime & " ms", 65, 190, 156, False
            Exit Sub
           '...............................................
            Case sPaquetes.TalkQuest
            If OnQTalk = 1 Then
            TextoRey = Right(Rdata, Len(Rdata) - 4)
            BodyQuest = STI(Rdata, 1)
            HeadQuest = STI(Rdata, 3)
            If LimitarFPS = 1 Then
            TiempoRey = LenB(TextoRey) * 2
            Else
            TiempoRey = LenB(TextoRey) * 4
            End If
            Else
            AddtoRichTextBox frmMain.RecTxt, Right(Rdata, Len(Rdata) - 4), 100, 100, 100, True, False
            End If
            Exit Sub
            '...............................................
            Case sPaquetes.pChangeUserCharCasco
            tempint = STI(Rdata, 1)
            CharList(tempint).Casco = CascoAnimData(Asc(Right$(Rdata, 1)))
            Exit Sub
            '...............................................
            Case sPaquetes.pChangeUserCharEscudo
            tempint = STI(Rdata, 1)
            CharList(tempint).Escudo = ShieldAnimData(Asc(Right$(Rdata, 1)))
            '...............................................
            Case sPaquetes.pChangeUserCharArmadura
            tempint = STI(Rdata, 1)
            CharList(tempint).iBody = STI(Rdata, 3)
            CharList(tempint).Body = BodyData(STI(Rdata, 3))
            Exit Sub
            '...............................................
            Case sPaquetes.pChangeUserCharArma
            tempint = STI(Rdata, 1)
            If StringToByte(Rdata, 3) > 0 Then CharList(tempint).Arma = WeaponAnimData(StringToByte(Rdata, 3))
            Exit Sub
            '...............................................
            Case sPaquetes.EnCentinela
            UserStats(SlotStats).UserCentinela = Not UserStats(SlotStats).UserCentinela
            Exit Sub
            '...............................................
            Case sPaquetes.TXAII
            Engine_FX.FX_Hit_Create_Pos Asc(Left(Rdata, 1)), Asc(mid(Rdata, 2, 1)), STI(Rdata, 3), 4000, mzColorApu
            'Call AddTxtAtaqueII(Asc(Left(Rdata, 1)), Asc(mid(Rdata, 2, 1)), STI(Rdata, 3), STI(Rdata, 5))
            Exit Sub
            '...............................................
            Case sPaquetes.EnviarStatsBasicas
            UserStats(SlotStats).UserMinHP = STI(Rdata, 1)
            UserStats(SlotStats).UserMinMAN = STI(Rdata, 3)
            UserStats(SlotStats).UserMinSTA = STI(Rdata, 5)
            frmMain.Hpshp.Width = (((UserStats(SlotStats).UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            frmMain.Label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
            frmMain.Label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
            frmMain.Label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
            frmMain.Label16.Caption = UserMinHAM & "/" & UserMaxHAM
            frmMain.Label17.Caption = UserMinAGU & "/" & UserMaxAGU
            If UserMaxMAN > 0 Then
                frmMain.ManShp.Width = (((UserStats(SlotStats).UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.ManShp.Width = 0
            End If
            frmMain.stashp.Width = (((UserStats(SlotStats).UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
          Exit Sub
        '...............................................
        Case sPaquetes.MensajeArmadas
        AddtoRichTextBox frmMain.RecTxt, Rdata, 100, 100, 255, True, False
        Exit Sub
        '...............................................
        Case sPaquetes.MensajeCaos
        AddtoRichTextBox frmMain.RecTxt, Rdata, 255, 10, 10, True, False
        Exit Sub
        '...............................................
        Case sPaquetes.EmpiezaTrabajo
        Istrabajando = True
        AddtoRichTextBox frmMain.RecTxt, "Empiezas a trabajar.", 65, 190, 156, False, False
        Exit Sub
         '...............................................
        Case sPaquetes.MensajeGlobal '"~190~190~190~0~1~"
        Rdata = Replace(Rdata, "~", " ")
        AddtoRichTextBox frmMain.RecTxt, Rdata, 190, 190, 190, False, True
        Exit Sub
        '...............................................
        Case sPaquetes.PartyAcomodarS
        Dim Caden() As String
        Caden = Split(Rdata, "|")
            frmPartyPorc.SkillsL = StringToByte(Right(Rdata, 1), 1)
            If UBound(Caden) > 10 Then Exit Sub
            For TempByte = 1 To (UBound(Caden)) / 2
                frmPartyPorc.Pj(TempByte).Caption = Caden(TempByte * 2 - 2)
                frmPartyPorc.Pj(TempByte).Visible = True
                frmPartyPorc.Porc(TempByte).Text = Caden(TempByte * 2 - 1) * 100
                frmPartyPorc.Porc(TempByte).Visible = True
                frmPartyPorc.Lin(TempByte).Visible = True
            Next TempByte
            frmPartyPorc.Show vbModal
        Exit Sub
       '...............................................
       Case sPaquetes.PPI
                For TempByte = 0 To 20
                If Listasolicitudes(TempByte) = "" Or Listasolicitudes(TempByte) = Rdata Then
                Listasolicitudes(TempByte) = Rdata
                Exit For
                End If
                Next TempByte
       Exit Sub
       '...............................................
       Case sPaquetes.PPE
        gh = False
        Liderparty = False

        For TempByte = 0 To 20
        Listasolicitudes(TempByte) = ""
        Next TempByte

        For TempByte = 0 To 20
        Listaintegrantes(TempByte) = ""
        Next TempByte

       Exit Sub
       '...............................................
       Case sPaquetes.Sefuedeparty
        For TempByte = 0 To 20
            If Listaintegrantes(TempByte) = Rdata Then
            Listaintegrantes(TempByte) = ""
            Exit For
            End If
        Next TempByte
        Exit Sub
     '..............................................
     Case sPaquetes.MensajeBoveda
     '   frmBancoObj.msgboveda = Mensaje(Asc(Rdata))
        Exit Sub
    '..............................................
    Case sPaquetes.IniciarAutoUpdater
        Dim mm As String
        mm = MsgBox("Hay una nueva version disponible. Si desea actualizarla pulse en si y el cliente se actualizara automaticamente. De lo contrario no podra seguir jugando.", vbExclamation + vbYesNo)
        If mm = vbYes Then
            If FileExist(App.path & "\Updater.exe", vbNormal) Then
            Call Shell(App.path & "\Updater.exe", vbNormalFocus)
            End
            Else
            mm = MsgBox("El AutoUpdater-TDS no se encuentra instalado. Por favor descarguelo desde www.aotds.com.ar", vbCritical, "AutoUpdater - Tierras del Sur")
            End If
        End If
        Exit Sub
    '..............................................
    Case sPaquetes.EstaEnvenenado
        IsEnvenenado = Not IsEnvenenado
        Exit Sub
     '..............................................
     Case sPaquetes.Actualizarestado
        tempint = STI(Rdata, 2)

        Select Case Asc(Left(Rdata, 1)) 'Que actualizamos?

        Case 1 'el clan aceptado
        CharList(tempint).Nombre = CharList(tempint).Nombre & "<" & Right$(Rdata, Len(Rdata) - 3) & ">"
        Exit Sub
        Case 2 ' chau clan
        CharList(tempint).Nombre = Right$(Rdata, Len(Rdata) - 3)
        Exit Sub
        Case 3 'Crimi o ciudadano
        CharList(tempint).criminal = Right$(Rdata, Len(Rdata) - 3)

        End Select
        Exit Sub
    '..............................................
    Case sPaquetes.MoverMuerto
        tempint = STI(Rdata, 1)
        Call Char_Move_by_Head2(tempint, Right$(Rdata, 1))
            If tempint = UserCharIndex Then
            Call Engine_MoveScreen(Right$(Rdata, 1))
            DoFogataFx
            End If
    Exit Sub
    '..............................................
    Case sPaquetes.ocultar
        tempint = STI(Rdata, 1)
        CharList(tempint).Oculto = True
        CharList(tempint).AlphaVal = 0
    Exit Sub
    '..............................................
    Case sPaquetes.Desocultar
        tempint = STI(Rdata, 1)
        CharList(tempint).Oculto = False
    Exit Sub
    '..............................................
    Case sPaquetes.pNpcActualizarPrecios
    'Borro el inventario actuaal
      '  frmComerciar.NpcInventarioComercio.Cls
        TempLong = 1
        For tempint = 1 To 20
            If mid(Rdata, TempLong, 1) <> "X" Then
            NPCInventory(tempint).Valor = StringToLong(Rdata, TempLong)
                'Si es el selecionado actualizo la pantallita de los datos.
                If frmComerciar.ItemElegidoC = tempint Then
                frmComerciar.Label3.Caption = NPCInventory(frmComerciar.ItemElegidoC).name & " " & "Def: " & NPCInventory(frmComerciar.ItemElegidoC).MinDef & "/" & NPCInventory(frmComerciar.ItemElegidoC).MaxDef & " Hit: " & NPCInventory(frmComerciar.ItemElegidoC).MinHit & "/" & NPCInventory(frmComerciar.ItemElegidoC).MaxHit & " Valor: " & NPCInventory(frmComerciar.ItemElegidoC).Valor
                End If
                TempLong = TempLong + 4
            Else
            TempLong = TempLong + 1
            NPCInventory(tempint).Valor = 0
            End If
        Next
    Exit Sub
    '..............................................
    Case sPaquetes.ActualizaNick
    tempint = STI(Rdata, 1)
    CharList(tempint).criminal = val(mid(Rdata, 3, 1))
    CharList(tempint).Nombre = mid(Rdata, 4)
    Exit Sub
    '..............................................
    Case sPaquetes.EquiparItem
    tempint = Asc(Rdata)
    UserInventory(tempint).Equipped = 1
    re_render_inventario = True: Call frmMain.picInv.Refresh
    Exit Sub
    '..............................................
    Case sPaquetes.DesequiparItem
    tempint = Asc(Rdata)
    UserInventory(tempint).Equipped = 1
    re_render_inventario = True: Call frmMain.picInv.Refresh
    Exit Sub
    '..............................................
    Case sPaquetes.ActualizaCantidadItem
    tempint = Asc(Left$(Rdata, 1))
    TempLong = DeCodify(mid(Rdata, 2))
    With UserInventory(tempint)
        If TempLong = 0 Then
        .OBJIndex = 0
        .Amount = 0
        .Equipped = 0
        .GrhIndex = 0
        .OBJType = 0
        .MaxHit = 0
        .MinHit = 0
        .MinDef = 0
        .Valor = 0
        .name = "(Nada)"
        Else
        .Amount = TempLong
        End If
    End With
    re_render_inventario = True: Call frmMain.picInv.Refresh
    Exit Sub
    '..............................................
    Case sPaquetes.ActualizarAreaUser
    tempint = STI(Rdata, 1)
    TempByte = Asc(mid(Rdata, 3, 1))
    TempByte2 = Asc(mid(Rdata, 4, 1))

    With CharList(tempint)
        '.active = 1
        .Heading = Asc(mid(Rdata, 5, 1))
        .Pos.X = TempByte
        .Pos.Y = TempByte2

        .iBody = STI(Rdata, 7)
        .iHead = STI(Rdata, 9)

        .Head = HeadData(.iHead)
        .Body = BodyData(.iBody)

        .Arma = WeaponAnimData(StringToByte(Rdata, 11))

        .Arma.WeaponAttack = 0
        .Escudo.ShieldAttack = 0
        .Escudo = ShieldAnimData(StringToByte(Rdata, 12))

        .Casco = CascoAnimData(StringToByte(Rdata, 13))

        CharMap(.Pos.X, .Pos.Y) = tempint
    End With
    SetCharacterFx tempint, StringToByte(Rdata, 6), 999
    Exit Sub
'..............................................

    Case sPaquetes.ActualizarAreanpc
    tempint = STI(Rdata, 1)
    TempByte = Asc(mid(Rdata, 3, 1))
    TempByte2 = Asc(mid(Rdata, 4, 1))

    CharList(tempint).Pos.X = TempByte
    CharList(tempint).Pos.Y = TempByte2
    CharList(tempint).active = 1

    CharList(tempint).Heading = StringToByte(Rdata, 5)

    CharMap(CharList(tempint).Pos.X, CharList(tempint).Pos.Y) = tempint
    Exit Sub
'..............................................
    Case sPaquetes.CambiarHeadingNpc
        tempint = STI(Rdata, 1)
        CharList(tempint).Heading = mid(Rdata, 3, 1)
    Exit Sub
'..............................................
    Case sPaquetes.BorrarArea
       Call BorrarAreaB
    Exit Sub
'..............................................
    Case sPaquetes.Pong
        EnviarPaquete Paquetes.Pong2, ""
    Exit Sub
'..............................................
    Case sPaquetes.SonidoTomarPociones
            If fx = 1 Then
                Call Audio.Sound_Play(46)
            End If
    Exit Sub
'..............................................
    Case sPaquetes.infoLogin
        MinPacketNumber = Asc(mid(Rdata, 1, 1))
        MaxPacketNumber = Asc(mid(Rdata, 2, 1))
        PacketNumber = MinPacketNumber
        Call LoginInit
    Exit Sub


End Select
   On Error GoTo 0
   Exit Sub

ProcesarPaquete_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure ProcesarPaquete of Módulo TCP. Paquete numero " & Asc(TempStr) & " Anexo " & Rdata
End Sub

#End If

