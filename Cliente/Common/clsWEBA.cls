VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWEBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
'Option Explicit
'Good just incase you misspell a variable
'Option Compare Binary 'AAA is equal to aaa

Private ACTION_POST As Boolean
Private raw As Boolean
Private send_data As String
Private URL As String
Private code As String

Event RecibeDatosWeb(ByRef datos As String, ByRef raw As Boolean)
Event Progreso(ByVal val As Long, ByVal max As Long)

Private WithEvents Winsock1 As MSWinsockLib.Winsock
Attribute Winsock1.VB_VarHelpID = -1

Private NIP As String
Private AIP As String

Public hay_conexion As Boolean
Private conter As Integer

Private es_binario As Boolean

Private iniciado As Boolean

Private Type typHOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 127) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function apiGetHostByName Lib "wsock32" Alias "gethostbyname" (ByVal hostname As String) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long

Private A_ENVIAR As String

Private Type Cole
    datos As String
    URL As String
    uri As String
End Type

Private Cola() As Cole
Private Ultimo As Integer
Private total As Integer
Private ENCola As Integer



Private BufferEntrada As String

Public conectadoAT As Long

Private no_procesar As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public BQ As clsByteQueue
Public TengoArchivo As Boolean

Private Sub Class_Terminate()

    If iniciado Then
        If Winsock1.State <> sckClosed Then Winsock1.close
        Set Winsock1 = Nothing
    End If

End Sub

Private Sub Connect()


    If Puedo Then
        Winsock1.Protocol = sckTCPProtocol
        If Len(AIP) Then
            Winsock1.Connect AIP, 80 ' host_web, 80
          Else
            Winsock1.Connect host_web, 80
        End If
    End If

End Sub

Private Function GetHttpHeaderValue(sHttpHeader As String, sValueName As String) As String

  Dim sHeaders() As String
  Dim vHeaders As Variant

  ''gets the value for corresponding header ie.
  ''File Content: pkzip file
  ''if sValueName is File Content than returns pkzip file

    sHeaders = Split(sHttpHeader, vbCrLf)
    For Each vHeaders In sHeaders
        vHeaders = LCase(vHeaders)
        If InStr(vHeaders, LCase(sValueName)) > 0 Then
            GetHttpHeaderValue = Trim$(mid$(vHeaders, InStr(vHeaders, Chr(32))))
            Exit Function
        End If
    Next

End Function

Public Function GetIPFromHost(ByVal hostname As String) As String
  Dim udtWSAData As WSADATA
  Dim HostAddress As Long
  Dim HostInfo As typHOSTENT
  Dim IPLong As Long
  Dim IPBytes() As Byte
  Dim i As Integer

    'Marce On error resume next

        If WSAStartup(257, udtWSAData) Then
            MsgBox "Error initializing winsock on WSAStartup!"
            GetIPFromHost = hostname
            Exit Function
        End If

        'Make sure a HTTP:// or FTP:// something wasn't added... some people like to do that
        If UCase$(Left$(hostname, 7)) = "HTTP://" Then
            hostname = Right$(hostname, Len(hostname) - 7)
          ElseIf UCase$(Left$(hostname, 6)) = "FTP://" Then
            hostname = Right$(hostname, Len(hostname) - 6)
        End If

        'If we were already passed an IP, just abort since we have what we want
        If IsIP(hostname) Then
            GetIPFromHost = hostname
            Exit Function
        End If

        'Get the host address
        HostAddress = apiGetHostByName(hostname)

        'Failure!
        If HostAddress = 0 Then Exit Function

        'Move the memory around to get it in a format we can read
        apiCopyMemory HostInfo, HostAddress, LenB(HostInfo)
        apiCopyMemory IPLong, HostInfo.hAddrList, 4

        'Get the number of parts to the IP (will always be 4 as far as I know)
        ReDim IPBytes(1 To HostInfo.hLength)

        'Convert the address, stored in the format of a long, to 4 bytes (just simple long -> byte array conversion)
        apiCopyMemory IPBytes(1), IPLong, HostInfo.hLength

        'Add in the periods
        For i = 1 To HostInfo.hLength
            GetIPFromHost = GetIPFromHost & IPBytes(i) & "."
        Next

        'Remove the final period
        GetIPFromHost = Left$(GetIPFromHost, Len(GetIPFromHost) - 1)

        'Clean up the socket
        WSACleanup

    'Marce 'Marce 'Marce On error goto 0

End Function

Public Function Initialize(WS As MSWinsockLib.Winsock) As Boolean
    If Not (WS Is Nothing) Then Set Winsock1 = WS
    hay_conexion = Not (Winsock1 Is Nothing)
    iniciado = hay_conexion
    AIP = GetIPFromHost(host_web)
    Initialize = hay_conexion
End Function

Private Function IsIP(ByVal IpAddress As String) As Boolean
  Dim s() As String
  Dim i As Long

  'If there are no periods, I have no idea what we have...

    If InStr(1, IpAddress, ".") = 0 Then Exit Function

    'Split up the string by the periods
    s = Split(IpAddress, ".")

    'Confirm we have ubound = 3, since xxx.xxx.xxx.xxx has 4 elements and we start at index 0
    If UBound(s) <> 3 Then Exit Function

    'Check that the values are numeric and in a valid range
    For i = 0 To 3
        If val(s(i)) < 0 Then Exit Function
        If val(s(i)) > 255 Then Exit Function
    Next i

    'Looks like we were passed a valid IP!
    IsIP = True

End Function

#If IsServer = 1 Then

Private Sub Loga(desc As String)

    If frmMain.sulog.Value = vbUnchecked Then Exit Sub
    On Error GoTo ErrHandler

  Dim nFile As Integer
    nFile = FreeFile 'obtenemos un canal
    Open app.Path & "\debug_" & Puerto & ".txt" For Append Shared As #nFile
    Print #nFile, date & " " & Time & " [" & menduz & "] " & desc
    Close #nFile
ErrHandler:

End Sub

#End If

Private Sub Logan(desc As String)

    On Error GoTo ErrHandler

  Dim nFile As Integer
    nFile = FreeFile 'obtenemos un canal
    Open app.Path & "\WEBA.txt" For Append Shared As #nFile
    Print #nFile, date & " " & Time & " " & desc
    Debug.Print date & " " & Time & " " & desc
    Close #nFile
ErrHandler:

End Sub

Public Function oURL(URL As String, Optional raw_data As String) As Boolean

    ACTION_POST = Len(raw_data) > 0
    send_data = raw_data
    URL = URL
    raw = True
    oURL = True
    
    TryRequest

End Function

Private Sub procesar_datos()

    Dim cabecera() As String
    
    Const CabezalArchivo As String = vbCrLf & vbCrLf & "FILE:" & vbCrLf & vbCrLf

    If InStr(1, BufferEntrada, vbCrLf & vbCrLf, vbTextCompare) <> 0 And (InStr(1, BufferEntrada, "HTTP/1.0 200 OK", vbTextCompare) Or InStr(1, BufferEntrada, "HTTP/1.1 200 OK", vbTextCompare)) Then
        cabecera = Split(BufferEntrada, vbCrLf & vbCrLf, 2)
        
        If InStr(cabecera(1), CabezalArchivo) <> 0 Then
            If BQ.Length Then
                'Marce On local error resume next
                BQ.ReadASCIIStringFixed InStr(1, cabecera(1), CabezalArchivo) + Len(CabezalArchivo) + Len(cabecera(0)) + 3
                cabecera(1) = Left(cabecera(1), InStr(1, cabecera(1), CabezalArchivo))
                TengoArchivo = True
                'Marce 'Marce On local error goto 0
            End If
        End If
        
        RaiseEvent RecibeDatosWeb(cabecera(1), False)
    Else
        RaiseEvent RecibeDatosWeb(BufferEntrada, True)
    End If
    BufferEntrada = ""

End Sub

Public Property Get Puedo() As Boolean

    If iniciado Then
        Puedo = (Winsock1.State = sckClosed)
        If Puedo = True Then Puedo = hay_conexion
    End If

End Property

Private Function Pop() As String

  Dim i As Integer
    If ENCola > 0 Then
        
        
        Pop = Cola(1).datos
        Cola(1).datos = ""
        ENCola = ENCola - 1
        For i = 0 To ENCola
            Cola(i) = Cola(i + 1)
        Next i
        
        If ENCola < 0 Then ENCola = 0
        ReDim Preserve Cola(maxl(ENCola + 1, 1)) As Cole
    End If
End Function

Private Function Peek() As String
    If ENCola > 0 Then
        Peek = Cola(1).datos
        If Len(Peek) = 0 Then Pop
    End If
End Function

Public Function SendEXT(URL As String, cReq As CHTTPRequest, codigo As String) As Boolean
    cReq.Host = host_web
    cReq.UserAgent = WebUserAgent
    cReq.Path = "/" & WEBPATH & ClientIDs & ";" & app.Major & "." & app.Minor & ";" & codigo & "/" & URL

    ENCola = ENCola + 1
    ReDim Preserve Cola(0 To ENCola + 1) As Cole
    
    Cola(ENCola).datos = cReq.GetPOSTQuery
    Cola(ENCola).URL = URL
    'If vWindowCDM Is Nothing Then
    '    Set vWindowCDM = New vWCDM
        'GUI_Load vWindowCDM
    'End If
    GUI_SetFocus vWindowCDM
    
    
    TryRequest
End Function

Public Function Send(URL As String, Optional raw_data As String, Optional codigo As String = vbNullString, Optional ByVal es_bin As Boolean = False) As Boolean
    Send = True

    ACTION_POST = Len(raw_data) > 0
    
    Dim buffer As String

    If ACTION_POST Then
        buffer = "POST "
    Else
        buffer = "GET "
    End If

    buffer = buffer & "/" & WEBPATH & ClientIDs & ";" & app.Major & "." & app.Minor & ";" & codigo & "/" & URL & " HTTP/1.0" & vbCrLf & _
             "Host: " & host_web & vbCrLf & _
             "User-Agent: " & WebUserAgent & vbCrLf & _
             "Connection: Close" & vbCrLf
    
    If ACTION_POST Then
        buffer = buffer & _
             "Content-Length: " & CStr(Len(raw_data)) & vbCrLf & _
             "Content-Type: application/x-www-form-urlencoded" & vbCrLf & vbCrLf & _
             raw_data & vbCrLf & vbCrLf
    Else
        buffer = buffer & vbCrLf
    End If

    ENCola = ENCola + 1
    ReDim Preserve Cola(0 To ENCola + 1) As Cole
    
    Cola(ENCola).datos = buffer

    'If vWindowCDM Is Nothing Then
    '    Set vWindowCDM = New vWCDM
        'GUI_Load vWindowCDM
    'End If
   ' vWindowCDM.Show
        GUI_SetFocus vWindowCDM

    frmCDM.FilenameActual = "Conectando..."

    TryRequest
End Function

Sub TryRequest()
    If ENCola > 0 And Puedo = True And Len(Peek) > 0 Then
        Connect
    Else
        If conectadoAT <> 0 Then
            If conectadoAT + 5000 < (GetTickCount() And &H7FFFFFFF) Then
            ' timeout
                If Winsock1.State <> sckClosed And Winsock1.State <> sckConnected Then
                    Winsock1.close
                    #If IsServer = 1 Then
                        If OFICIAL = 1 Then
                            Loga "TIMED_OUT"
                        End If
                    #End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Winsock1_Close()
    Winsock1.close
    procesar_datos
    conectadoAT = 0
    
    TryRequest
End Sub

Private Sub Winsock1_Connect()

    #If esCLIENTE = 1 Then
        cfnc = fnc.E_WEB_CONNECTD
    #End If
    #If IsServer = 1 Then
        If OFICIAL = 1 Then
            Loga "ENVIADO { " & Cola(1).uri & " }"
        End If
    #End If
Debug.Print "ENVIADO { " & Cola(1).datos & " }"

frmCDM.FilenameActual = "Enviando datos..."

    Winsock1.SendData Pop
    
    If BQ Is Nothing Then
        Set BQ = New clsByteQueue
    End If
    
    If BQ.Length Then
        BQ.ReadASCIIStringFixed BQ.Length
        TengoArchivo = False
    End If
    
    conectadoAT = GetTickCount() And &H7FFFFFFF
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)

  Dim datos As String
    Dim RD() As Byte
    
    Winsock1.GetData datos, vbString
    
    RD = StrConv(datos, vbFromUnicode)
    BQ.WriteBlock RD

    #If esCLIENTE = 1 Then
        If InStr(1, datos, "Transfer-Encoding: chunked", vbTextCompare) Then
            send_error "WEBA: CHUNKED -----------------REQUEST:" & vbNewLine & A_ENVIAR & vbNewLine & "-----------------RESPONSE:" & vbNewLine & datos
            Logan "CABECERA INCORRECTA."
        End If
    #End If

    BufferEntrada = BufferEntrada & datos

    hay_conexion = True

End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    If Winsock1.State <> sckClosed Then _
       Winsock1.close
    DoEvents
    TryRequest

End Sub

Public Sub clear_buffers()
ENCola = 0
ReDim Cola(2)
    If Winsock1.State <> sckClosed Then _
       Winsock1.close
conectadoAT = 0
End Sub


Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
conectadoAT = GetTickCount() And &H7FFFFFFF
RaiseEvent Progreso(bytesSent, bytesRemaining)
End Sub


