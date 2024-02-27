Attribute VB_Name = "EnlaceWeb"
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private conectorWeb As inet
Private conectorWebTimeOutConnectar As timer

Public Enum eEnlanceWebErrores
    ejcDireccionInalcanzable = 1
    ejcServerRetornaError = 2
    ejcErrorInterno = 3
End Enum

Public Sub iniciarConector()
    Set conectorWeb = frmConnect.inetConectorWeb
    Set conectorWebTimeOutConnectar = frmConnect.tmrInnetConnect
    
    'Configuro
    conectorWeb.Protocol = icHTTP
    conectorWeb.RemotePort = 80
    
    conectorWeb.RequestTimeout = 10 'Tiempo maximo entre que solicito los datos y el webserver se los da
    conectorWebTimeOutConnectar.Interval = 10 * 1000 ' Tiempo entre que se resuelve el DNS y se conecta
End Sub


'
' Retorno
' TODO OK: Un string con la respuesta del servidor
' MAL:  numerError es mayor a 0 y ejecutarConsulta retorna un valor de detalle del tipo de error

'           NumeroError     funcion     significado
'   ejcDireccionInalcanzable    0           No se pudo resolver la IP del loginServer
'   ejcServerRetornaError    Response code  El servidor devolvio un codigo de error.
'   ejcErrorInterno            Tipo error     Error interno
Public Function ejecutarConsulta(url As String, header As String, variables As String, numeroError As Byte) As String
    Dim sheader As String
    Dim retorno As String
    
    On Error GoTo error:
    
    ' Ejecuto la consulta
    sheader = "Content-type: application/json" & vbCrLf
    sheader = sheader & "User-Agent: " & header & vbCrLf

    ' Me aseguro de que este en otro hilo
    
    Dim pepe As Integer
    pepe = 0
    While conectorWeb.StillExecuting Or pepe < 10000
        DoEvents
        pepe = pepe + 1
    Wend
    
    conectorWebTimeOutConnectar.Enabled = True 'Activamos el timeout
   
    'Mando
    Debug.Print variables
    Call conectorWeb.Execute(url & "?data=" & variables, "post", sheader)

    ' Espero la respuesta
    While conectorWeb.StillExecuting
        DoEvents
    Wend
    
    conectorWebTimeOutConnectar.Enabled = False
    
    If conectorWeb.ResponseCode > 0 Then
        'Varios: No pudo resolver el nombre, TimeOut de recibir info, falla.
        numeroError = eEnlanceWebErrores.ejcServerRetornaError
        ejecutarConsulta = conectorWeb.ResponseCode
    ElseIf conectorWeb.tag = "icConnectingTimeOut" Then
        'No pudo resolver la IP del dominio. Problemas con los DNS.
        numeroError = eEnlanceWebErrores.ejcDireccionInalcanzable
        ejecutarConsulta = 0
    Else 'Esta todo bien
    
        'Obtengamos el retorno
        retorno = conectorWeb.GetChunk(1024, icString)
    
        Do While Len(retorno) = 1024
            
            ejecutarConsulta = ejecutarConsulta & retorno
            
            retorno = conectorWeb.GetChunk(1024, icString)
        Loop
            
        ejecutarConsulta = ejecutarConsulta & retorno
    End If
     
    Debug.Print retorno
    
    Exit Function
error:
    numeroError = eEnlanceWebErrores.ejcErrorInterno
    ejecutarConsulta = Err.Number
    
End Function
