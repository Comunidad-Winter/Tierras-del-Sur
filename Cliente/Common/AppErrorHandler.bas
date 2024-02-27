Attribute VB_Name = "AppErrorHandler"
Option Explicit

Public senderror As Boolean
Public error_string As String
Public dontpharsenext As Boolean
Public endthen As Boolean

Public funcion_actual As Long



Public Enum fnc
    E_Engine_Init
    E_Engine_Init_D3DDevice
    E_Get_Capabilities
    E_Render
    E_crons
    E_Engine_Calc_Screen_Moviment
    E_Map_Render
    E_Char_Render
    E_Engine_GetAngle
    E_Map_render_2array
    E_Main
    E_WEB_CONNECT
    E_WEB_INIT
    E_Set_Res
    E_WEB_CONNECTD
    E_RENDER_UI
    E_LOADGRH
    E_LOADMAP
End Enum

Public cfnc As Long

Public exerl As Long



Public Sub LogError(desc As String, Optional ByVal Comunicate As Boolean = False)
On Error GoTo ErrHandler

Dim nFile As Integer
nFile = FreeFile 'obtenemos un canal
Open App.path & "\ClientError.log" For Append Shared As #nFile
Print #nFile, Date & " " & Time & " " & desc
Debug.Print Date & " " & Time & " " & desc
Close #nFile

ErrHandler:

    
End Sub

Public Sub Log(desc As String)
On Error GoTo ErrHandler

Dim nFile As Integer
nFile = FreeFile 'obtenemos un canal
Open App.path & "\logs\LOG.txt" For Append Shared As #nFile
Print #nFile, Date & " " & Time & " " & desc
Debug.Print Date & " " & Time & " " & desc
Close #nFile
ErrHandler:
End Sub

Public Sub CriticError(desc As String)
On Error GoTo ErrHandler

Dim nFile As Integer
nFile = FreeFile
Open App.path & "\logs\ClientError.log" For Append Shared As #nFile
Print #nFile, Date & " " & Time & " CRITICO:" & get_machine_desc & desc
Debug.Print Date & " " & Time & " CRITICO:" & desc
Close #nFile

ErrHandler:
    Call MsgBox("Se ha producido un error cr�tico:" & vbNewLine & desc & vbNewLine & vbNewLine & "Por favor envienos el registro de errores """"Arduz/logs/ClientError.log"""" mediante nuestro foro. http://www.arduz.com.ar/", , "Arduz II")
    #If Debuging = 0 Then
        endthen = True
    #End If
End Sub

Private Function get_machine_desc() As String
Dim tmp$

get_machine_desc = tmp
End Function

Public Sub send_error(desc As String, Optional cerrar_programa As Boolean = False)
Dim tmp$
error_string = get_machine_desc & vbNewLine & desc
senderror = True
endthen = cerrar_programa

DoEvents
End Sub
