Attribute VB_Name = "Mod_ErrorLOG"
Option Explicit

Public Sub LogError(Desc As String)
'on error Resume Next
Dim nFile As Integer
nFile = FreeFile ' obtenemos un canal
Open App.path & "\errores.log" For Append As #nFile
Print #nFile, Desc
Close #nFile
End Sub


Public Sub IniciarDebug()
If FileExist(App.path & "\debug.log", vbNormal) Then Kill App.path & "\debug.log"
End Sub

Public Sub LogDebug(Desc As String)
'on error Resume Next
Dim nFile As Integer
nFile = FreeFile ' obtenemos un canal
Open App.path & "\debug.log" For Append As #nFile
Print #nFile, Desc
Close #nFile
End Sub


