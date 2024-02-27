VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallbakCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements callBackUrl

Private Nombre As String
Private password As String

Public Sub setDatos(nombreCuenta As String, passwordCuenta As String)
    Nombre = nombreCuenta
    password = passwordCuenta
End Sub
Private Function callBackUrl_EndRequest(request As CHTTPRequest, response As CHTTPResponse) As Boolean

    If Not response.Code = 201 Then
    
        If response.Code = 500 Then
            Call modDibujarInterface.mostrarError(0, "La cuenta o el mail ya se están usando.")
        Else
            Call modDibujarInterface.mostrarError(0, getErrorFromResponse(response))
        End If
        
        Exit Function
    End If
    
    Call frmConnect.conectar(Nombre, password)
End Function

Private Function callBackUrl_EndRequestError(request As CHTTPRequest, error As String, errornumber As Integer) As Boolean
    Call modDibujarInterface.mostrarError(0, error)
End Function


