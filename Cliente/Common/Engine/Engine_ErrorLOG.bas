Attribute VB_Name = "Engine_ErrorLOG"
'ESTE MODULO ESTA COMPARTIDO

Option Explicit

Public Sub LogError(Desc As String)
    '</EhHeader>
    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    Debug.Print Desc
    Open app.Path & "\errores.log" For Append As #nFile
    Print #nFile, Desc
    Close #nFile

    '</EhFooter>
End Sub

Public Sub LogDebug_Iniciar()
    If FileExist(app.Path & "\debug.log", vbNormal) Then
        Kill app.Path & "\debug.log"
    End If
    
    LogDebug "  "
    LogDebug "  "
    LogDebug "[Tierras Del Sur] " & Date
    LogDebug "  "
End Sub

Public Sub LogDebug(Desc As String)
    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    Open app.Path & "\debug.log" For Append As #nFile
    Print #nFile, GetTickCount() & "->" & Desc
    Close #nFile
End Sub

Public Sub LogCustomCDM(Desc As String, Optional file As String = "CerebroDeMono.log")
    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    Open app.Path & "\" & file For Append As #nFile
    Print #nFile, Desc
    Close #nFile
End Sub


