VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavegadorWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private WithEvents conectorWeb As inet
Attribute conectorWeb.VB_VarHelpID = -1
Private WithEvents timerTimeOut As timer
Attribute timerTimeOut.VB_VarHelpID = -1

Private callback As callBackUrl
Private httpRequest As CHTTPRequest

Private estado As String


Public Function consultaPendiente() As Boolean
    consultaPendiente = conectorWeb.StillExecuting
End Function

'CSEH: Con linea
Public Sub ejecutarConsulta(inet As inet, timer As timer, request As CHTTPRequest, callBackUrl As callBackUrl)
    Set callback = callBackUrl
    Set httpRequest = request
    Set conectorWeb = inet
    Set timerTimeOut = timer
    
    timerTimeOut.Enabled = True 'Activamos el timeout
    timerTimeOut.Interval = 10 * 1000
    
    Debug.Print "Ejecutando consulta: " & request.getURL
    ' Ejecutamos
    Call conectorWeb.Execute(request.getURL, request.getMethod, request.GetQuery, request.getHeaders)
End Sub

Private Sub Class_Terminate()
    Debug.Print "Terminado Navegador Web"
End Sub

Private Sub conectorWeb_StateChanged(ByVal State As Integer)
    
    Select Case State
          
            ' Obteniendo la direccion IP. Resolviendo DNS.
            Case icResolvingHost
                estado = "icResolvingHost"
            
            'Ya se tiene la IP
            Case icHostResolved
                estado = "icHostResolved"
              
            'Conectando al servidor
            Case icConnecting
                estado = "icConnecting"
            
            'Se conecto al servidor
            Case icConnected
                estado = "icConnected"
          
            'Enviando la consulta
            Case icRequesting
                estado = "icRequesting"
            
            'Solicitud enviada
            Case icRequestSent
                estado = "icRequestSent"
          
            'Recibiendo respuesta
            Case icReceivingResponse
                estado = "icReceivingResponse"
                
            'Respuesta recibida
            Case icResponseReceived
                estado = "icResponseReceived"
            
            'Descoenctando
            Case icDisconnecting
                estado = "icDisconnecting"
            
            'Descoenctado
            Case icDisconnected
                estado = "icDisconnected"
            'Fin
            Case icResponseCompleted
                estado = "icResponseCompleted"
                Call ejecutarCallBack
            'Error desconocido
            Case icError
                estado = "Error"
                Call ejecutarCallBack
        End Select

End Sub
Private Sub ejecutarCallBack()
    Dim intError As Integer
    Dim strError As String
    
    timerTimeOut.Enabled = False

    If conectorWeb.ResponseCode > 0 Then
        intError = conectorWeb.ResponseCode
        strError = conectorWeb.ResponseInfo
        'Varios: No pudo resolver el nombre, TimeOut de recibir info, falla.
        Call callback.EndRequestError(httpRequest, strError, intError)
    ElseIf estado = "icConnectingTimeOut" Then
        'No se pudo conectar despoues de un tiempo prudencial
        intError = 1
        strError = "No se ha podido conectar al Servidor."
        Call callback.EndRequestError(httpRequest, strError, intError)
    Else 'Esta todo bien

        'Obtengamos el retorno
        Dim retorno As String
        Dim Contenido As String
        
        retorno = conectorWeb.GetChunk(1024)

        Contenido = ""
        Do While Len(retorno) = 1024

            Contenido = Contenido & retorno

            retorno = conectorWeb.GetChunk(1024, icString)
        Loop

        Contenido = Contenido & retorno
        
        'Creamos
        Dim response As CHTTPResponse
        Set response = New CHTTPResponse
        Call response.crear(conectorWeb.GetHeader, Contenido)
        
        Call callback.EndRequest(httpRequest, response)
    End If
    
    Set callback = Nothing
    
End Sub

Private Sub timerTimeOut_Timer()
    Dim intError As Integer
    Dim strError As String
    
    'EL DNS no se pudo resolver aun?
    Call conectorWeb.Cancel
    
    intError = 1
    strError = "No se ha podido conectar al Servidor."
    
    LogError ("Timeout al buscar actualizaciones")
    
    Call callback.EndRequestError(httpRequest, strError, intError)

    timerTimeOut.Enabled = False
End Sub

