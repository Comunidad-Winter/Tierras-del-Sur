VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallBackConectarCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements callBackUrl

Private Function callBackUrl_EndRequest(request As CHTTPRequest, response As CHTTPResponse) As Boolean

    If Not response.Code = 200 Then
       Call modDibujarInterface.mostrarError(0, getErrorFromResponse(response))
       Exit Function
    End If
    
    Dim respuesta As Dictionary
    Set respuesta = response.bodyJSON
    
    Dim cuenta As cuenta
    Set cuenta = New cuenta
    cuenta.id = respuesta.item("usuario").item("id")
    cuenta.Nombre = respuesta.item("usuario").item("usuario")
    cuenta.Token = respuesta.item("token")
    cuenta.premium = respuesta.item("usuario").item("premium")
   
    Set MiCuenta.cuenta = cuenta
    
    frmConnect.cargarPersonajes
End Function

Private Function callBackUrl_EndRequestError(request As CHTTPRequest, error As String, errornumber As Integer) As Boolean
    Call modDibujarInterface.mostrarError(0, error)
End Function

