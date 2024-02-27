Attribute VB_Name = "CLI_CurrentLInk"
Option Explicit

Private timeExpiredLink As Long
Private destinationLink As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub setLink(texto As String, url As String)

    destinationLink = url
    timeExpiredLink = GetTickCount() + 60000
     
    frmMain.lblLink.Caption = texto
    frmMain.lblLink.Visible = True
   
End Sub

Public Sub pasarTiempo()
    If GetTickCount() > timeExpiredLink Then
        frmMain.lblLink.Caption = vbNullString
        frmMain.lblLink.Visible = False
        
        destinationLink = vbNullString
        timeExpiredLink = 0
    End If
End Sub

Public Sub clickLink()
    Dim url As String
    
    If destinationLink = vbNullString Then
        Exit Sub
    End If
    
    If InStr(1, destinationLink, "http", vbTextCompare) <> 1 Then
        url = "http://" & destinationLink
    Else
        url = destinationLink
    End If
    
    Call openUrl(url)
End Sub

Public Sub openUrl(url As String)
    ShellExecute frmMain.hWnd, "open", url, vbNullString, vbNullString, vbNormalFocus
End Sub
