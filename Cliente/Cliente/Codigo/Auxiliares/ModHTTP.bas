Attribute VB_Name = "ModHTTP"
Option Explicit

Public Enum eHttpMethod
    httpGET
    httppost
    httpPUT
    httpDELETE
    httpPATCH
    httpUNLOCK
End Enum


Public Function getErrorFromResponse(response As CHTTPResponse) As String

    Dim respuesta As Dictionary
    
    On Error GoTo hayerror:
    
    If response.Code = 500 Then GoTo hayerror:
    
    Set respuesta = response.bodyJSON
    
    getErrorFromResponse = respuesta.item("message")
    

    Exit Function
hayerror:
       getErrorFromResponse = "Error desconocido. Intente nuevamente."
       Call LogError(response.body)
       Exit Function
End Function


