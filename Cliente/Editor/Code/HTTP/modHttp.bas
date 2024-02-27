Attribute VB_Name = "modHttp"
Option Explicit

Public Enum eHttpMethod
    httpGET
    httppost
    httpPUT
    httpDELETE
    httpPATCH
End Enum

' http://support.microsoft.com/kB/181050
Private Const REGISTRO_TIMEOUT_PATH = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
Private Const REGISTRO_TIMEOUT_CLAVE = "ReceiveTimeout"
Private Const REGISTRO_TIMEOUT_VALOR_MINIMO = 1800000 ' 30 minutos * 60 segundos * 1000


Public Function iniciar() As Boolean

Dim valorActual As Long

valorActual = CDMRegistry.GetDWord(HKEY_CURRENT_USER, REGISTRO_TIMEOUT_PATH, REGISTRO_TIMEOUT_CLAVE)

If valorActual < REGISTRO_TIMEOUT_VALOR_MINIMO Then
    Call CDMRegistry.SaveDWord(HKEY_CURRENT_USER, REGISTRO_TIMEOUT_PATH, REGISTRO_TIMEOUT_CLAVE, REGISTRO_TIMEOUT_VALOR_MINIMO)
    iniciar = False
Else
    iniciar = True
End If

End Function
