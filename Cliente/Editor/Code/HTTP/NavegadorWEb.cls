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
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : NavegadorWeb
'    Project    : CerebroDeMono2
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private WithEvents conectorWeb As Inet
Attribute conectorWeb.VB_VarHelpID = -1
Private WithEvents timerTimeOut As Timer
Attribute timerTimeOut.VB_VarHelpID = -1

Private intError As Integer
Private strError As String

Private estado As String



Public Sub crear(control_ As Inet, timer_ As Timer)

    Set conectorWeb = control_
    Set timerTimeOut = timer_
    'Configuro
    
    conectorWeb.Protocol = icHTTP
    conectorWeb.RemotePort = 80

    conectorWeb.RequestTimeout = 300 'Tiempo maximo entre que solicito los datos y el webserver se los da
    
    timerTimeOut.Interval = 30 * 1000 ' Tiempo entre que se resuelve el DNS y se conecta

End Sub


Public Function ejecutarConsulta(request As CHTTPRequest) As CHTTPResponse
    Dim sheader As String
    Dim retorno As String
    Dim Contenido As String
    
    On Error GoTo error:
    
    Set ejecutarConsulta = Nothing
    
    ' Me aseguro de que este en otro hilo
    While conectorWeb.StillExecuting
        DoEvents
    Wend
        
    timerTimeOut.Enabled = True 'Activamos el timeout
       
    ' Ejecutamos
    Call conectorWeb.Execute(request.getURL, request.getMethod, request.GetQuery, request.getHeaders)

    ' Espero la respuesta
    While conectorWeb.StillExecuting
        DoEvents
    Wend
    
    timerTimeOut.Enabled = False

    If conectorWeb.ResponseCode > 0 Then
        intError = conectorWeb.ResponseCode
        strError = conectorWeb.ResponseInfo
        'Varios: No pudo resolver el nombre, TimeOut de recibir info, falla.
    ElseIf estado = "icConnectingTimeOut" Then
        'No se pudo conectar despoues de un tiempo prudencial
        intError = 1
        strError = "No se ha podido conectar al Servidor."
    Else 'Esta todo bien

        'Obtengamos el retorno
        retorno = conectorWeb.GetChunk(1024, icString)

        Contenido = ""
        Do While Len(retorno) = 1024

            Contenido = Contenido & retorno

            retorno = conectorWeb.GetChunk(1024, icString)
        Loop

        Contenido = Contenido & retorno
        
        'Creamos
        Set ejecutarConsulta = New CHTTPResponse
        Call ejecutarConsulta.crear(conectorWeb.GetHeader, Contenido)
        
    End If

    Exit Function
error:
Debug.Print Err.description
End Function

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
            
            'Error desconocido
            Case icError
                estado = "Error"
        End Select

End Sub

Private Sub timerTimeOut_Timer()
    'EL DNS no se pudo resolver aun?
    If estado = "icResolvingHost" Or estado = "icConnecting" Then
        estado = "icConnectingTimeOut"
        Call conectorWeb.Cancel
    End If
End Sub

Public Property Get error() As Integer
    error = intError
End Property

Public Property Get errorDescripcion() As String
    errorDescripcion = strError
End Property

