VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vWCDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements vWindow

Private vX As Integer
Private vY As Integer
Private vW As Integer
Private vH As Integer

Private alpha As Byte

Private Progreso As Single

Private InicioTick As Long
Private FinTick As Long

Private Matar As Boolean

Private I1 As Single
Private I2 As Single


Private wLabel As clsGUIText

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow
'/Implementacion


Private WithEvents Downloader As asyncDownload
Attribute Downloader.VB_VarHelpID = -1

Private NumeroDescargandoScript As Integer
Private estado As New vWControlChat

Private Bmax As Long
Private bMin As Long

Private DownloadStack As clsStack
Private TmpFile As CDM_File

Private Sub Class_Terminate()
    Set wLabel = Nothing

End Sub

Private Sub Class_Initialize()
Set wLabel = New clsGUIText
Set Downloader = frmCDM.asyncDownload1
ReDim scr(0)
InicioTick = GetTimer
FinTick = InicioTick + 0
I1 = 0.25
I2 = 1

wLabel.text = vTextoAlerta
Titulo = vAlerteTitle
estado.CantidadDialogos = 4
End Sub

Private Sub Downloader_DownloadComplete(bytes As Long)
On Error GoTo enda
    TmpFile = DownloadStack.Peek
    If Len(TmpFile.FileName) = 0 Then
        TmpFile.FileName = (GetTickCount Xor (Not RENDERCRC)) & ".tmp"
    End If
    With TmpFile
        On Local Error Resume Next
            If FileExist(CDM_TMP_PATH & .FileName & ".backup", vbNormal) Then Kill CDM_TMP_PATH & .FileName & ".backup"
            DoEvents
            Name (CDM_TMP_PATH & .FileName) As (CDM_TMP_PATH & .FileName & ".backup")
        On Local Error GoTo enda
        
        DoEvents
        
        Downloader.SaveAs CDM_TMP_PATH & .FileName
        
        If FileExist(CDM_TMP_PATH & .FileName, vbNormal) Then

            
            DoEvents
            If LCase(MD5File(CDM_TMP_PATH & .FileName)) = LCase(.MD5) Then
                Dim Pak As clsEnpaquetado
                
                If .Tipo = CDM_TIPO_UPDATE.CDM_Upd_Grafico Then Set Pak = pakGraficos
                If .Tipo = CDM_TIPO_UPDATE.CDM_Upd_Sonidos Then Set Pak = pakSonidos
                If .Tipo = CDM_TIPO_UPDATE.CDM_Upd_Mapas Then Set Pak = pakMapasME
                
                If Not Pak Is Nothing Then
                    If Pak.Parchear(.numero, CDM_TMP_PATH & .FileName, .id) Then
                        Dim TmpIH As INFOHEADER
                        
                        If Pak.IH_Get(.numero, TmpIH) Then
                            TmpIH.complemento_1 = .C1
                            TmpIH.complemento_2 = .C2
                            TmpIH.complemento_3 = .C3
                            TmpIH.complemento_4 = .C4
                            TmpIH.owner = .user
                            TmpIH.privs = .PrivsPublicos
                            Pak.IH_Mod .numero, TmpIH
                        End If
                        
                        If Not ((.user = CDM_UserID And CDM_UserID > 0) Or CDM_UserPrivs = -1 Or ((.PrivsPublicos And CDM_Privs.PuedeExtraerArchivos) And (CDM_UserPrivs And CDM_Privs.PuedeExtraerArchivos))) Then
                            'Marce On local error resume next
                            Kill CDM_TMP_PATH & .FileName
                            On Local Error GoTo enda
                        End If
                        
                    End If
                    
                    If .Tipo = CDM_TIPO_UPDATE.CDM_Upd_Grafico Then BorrarTexturaDeMemoria .numero
                    If .Tipo = CDM_TIPO_UPDATE.CDM_Upd_Mapas Then ActualizarPuedoMapas
                End If
                
                Set Pak = Nothing
                .EstaListo = True
                
                If .id > CDM_Revision Then
                    Call WriteVar(app.Path & "\ME.ini", "CEREBRO_DE_MONO", "Revision", .id)
                    CDM_Revision = .id
                End If
                
                
            Else
                Kill CDM_TMP_PATH & .FileName
                CDMOutput "El archivo " & .FileName & " está corrupto o no es el mismo que en el servidor. " & LCase(MD5File(CDM_TMP_PATH & .FileName)) & "-" & LCase(.MD5)
                        MsgBox MD5File(CDM_TMP_PATH & .FileName) & "-" & .MD5

            End If
            DownloadStack.Pop
        End If
    End With
enda:
    
DescargarProximo

End Sub

Friend Sub SetStack(stack As clsStack)
    Set DownloadStack = stack
    DescargarProximo
End Sub

Private Sub DescargarProximo()
    
    If Downloader.Busy = False Then
        If Not DownloadStack Is Nothing Then
            If DownloadStack.Count Then
                TmpFile = DownloadStack.Peek

                If Len(TmpFile.Remote) = 0 Then
                    DownloadStack.Pop
                    If DownloadStack.Count Then DescargarProximo
                Else
                    estado.PushBackText "Descargando " & Chr$(255) & TmpFile.FileName, -1
                    Downloader.DownloadStart TmpFile.Remote
                End If
            End If
            If DownloadStack.Count = 0 Then
                vWindow_Hide
            End If
        End If
    End If

End Sub

Private Sub Downloader_DownloadFailed(numero As Long, desc As String)
    TmpFile.Fallo = True
End Sub

Private Sub Downloader_DownloadProgress(b1 As Long, b2 As Long, max As Long)
    Bmax = max
    bMin = (b1 + b2) / 2
End Sub

Private Property Get vWindow_GetControl() As vControles

End Property

Private Sub vWindow_Hide()
    I1 = 1
    I2 = 0
    InicioTick = GetTimer
    FinTick = InicioTick + 300
    Matar = True
End Sub

Private Function vWindow_Render() As Boolean

    DescargarProximo

    Dim TmpTick As Long
    
    TmpTick = GetTimer
    'Decimos "Si, la ventana sigue viva."
    vWindow_Render = True
    
    If TmpTick < FinTick Then
        Progreso = CosInterp(I1, I2, (TmpTick - InicioTick) / (FinTick - InicioTick))
    Else
        Progreso = 1
        If I2 = 0 Then
            'Decimos "La ventana acaba de morirse wachin"
            vWindow_Render = Not Matar
            If vWindow_Render = False Then Exit Function
        End If
    End If
    
    w = 300
    vW = (w + 16) * Progreso
    vH = (h + 96) * Progreso
    vX = D3DWindow.BackBufferWidth - vW - 16
    vY = CosInterp(512, D3DWindow.BackBufferHeight - vH - 16, Progreso)

    'Dibujo la ventana
    GUI_RenderDialog vX, vY, vW, vH, "CDM", Me, Progreso
    
    If Progreso = 1 Then
        Engine.Draw_FilledBox vX, vY + 24, ((bMin + 1) / (Bmax + 1)) * vW, vH - 24, &H7F007F00, 0, 0
        estado.Draw vX + 8, vY + 24 + 8
    End If
    
End Function

Private Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)

End Sub

Private Sub vWindow_Show()

End Sub

Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function
Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
