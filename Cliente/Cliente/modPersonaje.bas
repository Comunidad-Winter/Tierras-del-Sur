Attribute VB_Name = "modPersonaje"
Option Explicit



Public Sub actualizarNick(ByRef Char As Char, nombreyClan As String)

    If right$(nombreyClan, 1) = ">" Then
        Char.Clan = right$(nombreyClan, Len(nombreyClan) - InStr(nombreyClan, "<") + 1)
        Char.Nombre = left$(nombreyClan, InStr(nombreyClan, "<") - 1)
        Char.flags = (Char.flags Or ePersonajeFlags.tieneClan)
    Else
        Char.Clan = ""
        Char.flags = Char.flags And Not ePersonajeFlags.tieneClan
        Char.Nombre = nombreyClan
    End If
    
End Sub

Public Function isGameMaster(ByRef Char As Char) As Boolean
    isGameMaster = Not (Char.priv = 0 Or Char.priv = 5)
End Function

Public Function getHexaColorByPrivForInterface(ByRef Char As Char) As Long
      Select Case Char.priv
        Case 0
            If (Char.alineacion = eAlineaciones.caos) Then
                getHexaColorByPrivForInterface = &HFF&
            ElseIf Char.alineacion = eAlineaciones.Real Then
                getHexaColorByPrivForInterface = &HFFFF00
            ElseIf Char.alineacion = eAlineaciones.Neutro Then
                getHexaColorByPrivForInterface = &HB0AAA3
            End If
        Case 1 ' Consejero
           getHexaColorByPrivForInterface = &HCAE8F6
        Case 2 ' Semidios
           getHexaColorByPrivForInterface = &HCAE8F6
        Case 3 'Dios
           getHexaColorByPrivForInterface = &HCAE8F6
        Case 4 ' Admin
           getHexaColorByPrivForInterface = &HCAE8F6
        Case 5 ' Concilio/ Consejos
            If (Char.alineacion = eAlineaciones.caos) Then
                getHexaColorByPrivForInterface = &H8080FF
            ElseIf Char.alineacion = eAlineaciones.Real Then
                getHexaColorByPrivForInterface = &HFFFF80
            ElseIf Char.alineacion = eAlineaciones.Neutro Then
                getHexaColorByPrivForInterface = &HFFDC73
            End If
        End Select
End Function

Public Function getHexaColorByPrivForDialog(ByRef Char As Char) As Long
      Select Case Char.priv
        Case 0
            getHexaColorByPrivForDialog = mzWhite
        Case 1 ' Consejero
           getHexaColorByPrivForDialog = mzWhite
        Case 2 ' Semidios
           getHexaColorByPrivForDialog = mzWhite
        Case 3 'Dios
           getHexaColorByPrivForDialog = mzWhite
        Case 4 ' Admin
           getHexaColorByPrivForDialog = mzWhite
        Case 5 ' Concilio/ Consejos
            If (Char.alineacion = eAlineaciones.caos) Then
                getHexaColorByPrivForDialog = &HFFF6363
            ElseIf Char.alineacion = eAlineaciones.Real Then
                getHexaColorByPrivForDialog = &HF7DEFFF
            Else
                getHexaColorByPrivForDialog = &HFFA500
            End If
        End Select
End Function
Public Sub setColorNombre(ByRef Char As Char)

    Select Case Char.priv
        Case 0
            If (Char.alineacion = eAlineaciones.caos) Then
                Set Char.NickLabel = New clsGUIText
                Char.NickLabel.text = Char.Nombre
                Char.NickLabel.color = D3DColorXRGB(255, 0, 0)
                Char.NickLabel.GradientMode = dSolid
                Char.NickLabel.Centrar = True
                
                If Not Char.Clan = "" Then
                    Set Char.NickClan = New clsGUIText
                    Char.NickClan.text = Char.Clan
                    Char.NickClan.color = D3DColorXRGB(255, 0, 0)
                    Char.NickClan.GradientMode = dSolid
                    Char.NickClan.Centrar = True
                Else
                    Set Char.NickClan = Nothing
                End If
            ElseIf Char.alineacion = eAlineaciones.Real Then
                Set Char.NickLabel = New clsGUIText
                Char.NickLabel.text = Char.Nombre
                Char.NickLabel.color = D3DColorXRGB(0, 128, 255)
                Char.NickLabel.GradientMode = dSolid
                Char.NickLabel.Centrar = True
                
                If Not Char.Clan = "" Then
                    Set Char.NickClan = New clsGUIText
                    Char.NickClan.text = Char.Clan
                    Char.NickClan.color = D3DColorXRGB(0, 128, 255)
                    Char.NickClan.GradientMode = dSolid
                    Char.NickClan.Centrar = True
                Else
                    Set Char.NickClan = Nothing
                End If
            Else
                Set Char.NickLabel = New clsGUIText
                Char.NickLabel.text = Char.Nombre
                Char.NickLabel.color = D3DColorXRGB(176, 170, 163)
                Char.NickLabel.GradientMode = dSolid
                Char.NickLabel.Centrar = True
                
                If Not Char.Clan = "" Then
                    Set Char.NickClan = New clsGUIText
                    Char.NickClan.text = Char.Clan
                    Char.NickClan.color = D3DColorXRGB(176, 170, 163)
                    Char.NickClan.GradientMode = dSolid
                    Char.NickClan.Centrar = True
                Else
                    Set Char.NickClan = Nothing
                End If
            End If
        Case 1 ' Consejero
            Set Char.NickLabel = New clsGUIText
            Char.NickLabel.text = Char.Nombre
            Char.NickLabel.color = D3DColorXRGB(30, 150, 0)
            Char.NickLabel.GradientMode = dSolid
            Char.NickLabel.Centrar = True
            
            If Not Char.Clan = "" Then
                Set Char.NickClan = New clsGUIText
                Char.NickClan.text = Char.Clan
                Char.NickClan.color = D3DColorXRGB(30, 150, 0)
                Char.NickClan.GradientMode = dSolid
                Char.NickClan.Centrar = True
            Else
                Set Char.NickClan = Nothing
            End If
        Case 2 ' Semidios
            Set Char.NickLabel = New clsGUIText
            Char.NickLabel.text = Char.Nombre
            Char.NickLabel.color = D3DColorXRGB(30, 255, 30)
            Char.NickLabel.GradientMode = dSolid
            Char.NickLabel.Centrar = True
            
            If Not Char.Clan = "" Then
                Set Char.NickClan = New clsGUIText
                Char.NickClan.text = Char.Clan
                Char.NickClan.color = D3DColorXRGB(30, 255, 30)
                Char.NickClan.GradientMode = dSolid
                Char.NickClan.Centrar = True
            Else
                Set Char.NickClan = Nothing
            End If
        Case 3 'Dios
            Set Char.NickLabel = New clsGUIText
            Char.NickLabel.text = Char.Nombre
            Char.NickLabel.color = D3DColorXRGB(250, 250, 150)
            Char.NickLabel.GradientMode = dSolid
            Char.NickLabel.Centrar = True
            
            If Not Char.Clan = "" Then
                Set Char.NickClan = New clsGUIText
                Char.NickClan.text = Char.Clan
                Char.NickClan.color = D3DColorXRGB(250, 250, 150)
                Char.NickClan.GradientMode = dSolid
                Char.NickClan.Centrar = True
            Else
                Set Char.NickClan = Nothing
            End If
        Case 4 ' Admin
            Set Char.NickLabel = New clsGUIText
            Char.NickLabel.text = Char.Nombre
            Char.NickLabel.color = D3DColorXRGB(255, 165, 0)
            Char.NickLabel.GradientMode = dSolid
            Char.NickLabel.Centrar = True
            
            If Not Char.Clan = "" Then
                Set Char.NickClan = New clsGUIText
                Char.NickClan.text = Char.Clan
                Char.NickClan.color = D3DColorXRGB(255, 165, 0)
                Char.NickClan.GradientMode = dSolid
                Char.NickClan.Centrar = True
            Else
                Set Char.NickClan = Nothing
            End If
        Case 5 ' Concilio/ Consejos
            Dim color As Long
            
            If Char.alineacion = eAlineaciones.caos Then
                color = D3DColorXRGB(128, 128, 128)
            ElseIf Char.alineacion = eAlineaciones.Real Then
                color = D3DColorXRGB(0, 195, 195)
            Else
                color = D3DColorXRGB(255, 165, 0)
            End If
            
            Set Char.NickLabel = New clsGUIText
            
            Char.NickLabel.text = Char.Nombre
            Char.NickLabel.color = color
            Char.NickLabel.GradientMode = dSolid
            Char.NickLabel.Centrar = True
    
            If Not Char.Clan = "" Then
                Set Char.NickClan = New clsGUIText
                Char.NickClan.text = Char.Clan
                Char.NickClan.color = color
                Char.NickClan.GradientMode = dSolid
                Char.NickClan.Centrar = True
            Else
                Set Char.NickClan = Nothing
            End If

        End Select

End Sub
