Attribute VB_Name = "CLI_Grabador"
Public Grabando As Boolean
Public RutaVideo As String

Public MinutoVideo As Byte
Public SegVideo As Byte
Public Sub IniciarGrabacion()
    Dim NumVid As Integer
    Dim i As Long
    
    frmMain.lGrabando.Caption = "Iniciando Grabación..."
    frmMain.lGrabando.Visible = True
    DoEvents
    PS = 0
    
    ReDim TempPaq(0)
    StartTC = GetTickCount
    NumVid = val(GetVar(getConfigFilePath, "INIT", "NumVid")) + 1
    
    Call WriteVar(getConfigFilePath, "INIT", "NumVid", CStr(NumVid))
    
    If Not FileExist(app.Path & "\Videos\", vbDirectory) Then
        MkDir app.Path & "\Videos\"
    End If
    
    RutaVideo = app.Path & "\Videos\VideoTDS-" & NumVid & ".vtd"
    
    Open (RutaVideo) For Binary As #12
    Dim a As String * 50
    Dim b As String * 2500
    Dim c As String * 30
    b = frmConsola.ConsolaFlotante.TextRTF
    
    c = UserName
    Put #12, , StartTC
    Put #12, , c
    Put #12, , b
    Put #12, , UserInventory
                            
    Put #12, , UserStats(SlotStats).UserMinSTA
    Put #12, , UserMaxSTA
                             
    Put #12, , UserStats(SlotStats).UserMinMAN
    Put #12, , UserMaxMAN
                               
    Put #12, , UserStats(SlotStats).UserMinHP
    Put #12, , UserMaxHP
                                 
    Put #12, , UserMinHAM
    Put #12, , UserMaxHAM
                                   
    Put #12, , UserMinAGU
    Put #12, , UserMaxAGU
                                     
    Put #12, , frmMain.hlst.ListIndex
                            
    If frmMain.CmdLanzar.Visible = False Then
        Put #12, , CByte(0)
    Else
        Put #12, , CByte(1)
    End If
                            
    For i = 0 To MAXHECHI - 1
        a = frmMain.hlst.list(i)
        Put #12, , a
    Next
                            
    For i = 1 To 10000
        If CharList(i).active Then
            Put #12, , i
            'FIXME Put #12, , CharList(i)
        End If
    Next i
                                
    TempStr = 12345
    
    Put #12, , CLng(0)
    Put #12, , UserCharIndex
    Put #12, , UserPos
    
'    For x = 1 To 100
'         For y = 1 To 100
'            Put #12, , MapData(x, y)
'         Next y
'    Next x
    
    Grabando = Not Grabando
    frmMain.lGrabando.Caption = "Grabando video..."

End Sub

Public Sub FinalizarGrabacion()
    frmMain.lGrabando.Caption = "Finalizando Grabación..."
    CrearAccion ("FIN")
    
    ReDim Preserve TempPaq(UBound(TempPaq))
                            
    TempPaq(UBound(TempPaq)).TC = 0
    TempPaq(UBound(TempPaq)).Rdata = ""
    
    For i = 1 To UBound(TempPaq)
            Put #12, , TempPaq(i)
    Next i
                    
    Seek #12, 1
    Put #12, , GetTickCount - StartTC
                    
    DoEvents
    
    frmMain.lGrabando.Visible = False
    Close #12
    
    Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Video grabado exitosamente en " & RutaVideo, 0, 200, 200, False, False, False)
    
    Grabando = False
    MinutoVideo = 0
    SegVideo = 0
End Sub

