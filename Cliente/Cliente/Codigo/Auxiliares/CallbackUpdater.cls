VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallbackUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements callBackUrl

Public status As Integer

Private Function CallbackURL_EndRequest(request As CHTTPRequest, response As CHTTPResponse) As Boolean
    On Error GoTo hayerror:
    Dim resultado As Dictionary
    
    Set resultado = response.bodyJSON
        
    If resultado.item("version") > Configuracion_Usuario.versionActual Then
        Call juegoDesactualizado(Configuracion_Usuario.versionActual, resultado.item("version"))
    Else
        status = 1
    End If
hayerror:
        status = 2
End Function

Private Function callBackUrl_EndRequestError(request As CHTTPRequest, error As String, errornumber As Integer) As Boolean
    status = 2
End Function

Private Sub Class_Initialize()
    status = 0
End Sub
