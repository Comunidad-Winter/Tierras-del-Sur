Attribute VB_Name = "CLI_Updater"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const HTTP_URL = "https://tierrasdelsur.cc"
Public Const UPDATER_PATH = "version"

Public Sub juegoDesactualizado(Optional ByVal versionActual As Integer = 0, Optional ByVal versionNueva As Integer = 0)
    Dim mm As String
    Dim comando As String
    
    If versionNueva > 0 Then
        mm = MsgBox("Hay una nueva version disponible. Usted posee la versión " & versionActual & " y la última versión es la " & versionNueva & ". Si desea actualizarla pulse en si y el cliente se actualizara automaticamente. De lo contrario no podrá seguir jugando.", vbExclamation + vbYesNo)
    Else
        mm = MsgBox("Hay una nueva version disponible. Si desea actualizarla pulse en si y el cliente se actualizara automaticamente. De lo contrario no podrá seguir jugando.", vbExclamation + vbYesNo)
    End If
    
    If mm = vbYes Then
        If FileExist(app.Path & "\Updater.exe", vbNormal) Then
            comando = Chr$(34) & app.Path & "\Updater.exe" & Chr$(34)
            Call ejecutarUpdater(comando)
            End
        Else
            mm = MsgBox("El AutoUpdater-TDS no se encuentra instalado. Por favor descarguelo desde www.tierrasdelsur.cc", vbCritical, "AutoUpdater - Tierras del Sur")
        End If
    End If
End Sub

Private Sub ejecutarUpdater(comando As String)
    On Error GoTo hayerror:
    Dim result As Long
    
    result = ShellExecute(0, "runas", "Updater.exe", "", CurDir$(), vbNormalFocus)
    
    If Not (result < 0 Or result > 32) Then
        MsgBox "Surgió un error al momento de ejecutar el Updater. Proba ejecutando el programa Updater.exe de manera manual", vbCritical
    End If
    
    Exit Sub
hayerror:
    LogDebug (Err.Description)
    MsgBox "Surgió un error al momento de ejecutar el Updater. Proba ejecutando el programa Updater.exe de manera manual", vbCritical
End Sub


