Attribute VB_Name = "API_Manager"
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Option Explicit

Private IDmanager As Long

Private Const WM_COPYDATA = &H4A

Private Declare Function FindWindow Lib "user32" Alias _
         "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
         As String) As Long

Private Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal _
         wParam As Long, lParam As Any) As Long

'Copies a block of memory from one location to another.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
         
Public Const NOMBRE_APP_MANAGER = "MANAGER_TIERRAS_DEL_SUR"

Public Enum eManagerPaquetes
    eApuestas = 1
End Enum

Public Function iniciarManager() As Boolean
    IDmanager = FindWindow(vbNullString, NOMBRE_APP_MANAGER)
    
    iniciarManager = (IDmanager > 0)
End Function

Public Function enviarMensaje(paquete As eManagerPaquetes, mensaje As String) As Long
    Dim cds As COPYDATASTRUCT
    Dim buf(1 To 1024) As Byte

    If Len(mensaje) < 1000 Then
        ' Copy the string into a byte array, converting it to ASCII
        Call CopyMemory(buf(1), ByVal (Chr$(paquete) & mensaje), Len(mensaje) + 1)
        cds.dwData = 3
        cds.cbData = Len(mensaje) + 1 + 1
        cds.lpData = VarPtr(buf(1))
        
        enviarMensaje = SendMessage(IDmanager, WM_COPYDATA, 0, cds)
    Else
        enviarMensaje = 0
        Call LogError("No es posible el envio de un mensaje a través del API por ser demasiado largo.")
    End If
End Function
