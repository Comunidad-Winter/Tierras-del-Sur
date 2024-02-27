Attribute VB_Name = "wskapiAO"
Option Explicit

''
' Modulo para manejar Winsock
'


'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_BOTH As Long = &H2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192
''
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
'
' @param Sock sock
' @param slot slot
'
Public Type tSockCache
    Sock As Long
    slot As Long
End Type

Public WSAPISock2Usr As New Collection

' ====================================================================================
' ====================================================================================

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

' ====================================================================================
' ====================================================================================

Public SockListen As Long
Public LastSockListen As Long


' ====================================================================================
' ====================================================================================


'---------------------------------------------------------------------------------------
' Procedure : IniciaWsApi
' DateTime  : 18/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub IniciaWsApi(ByVal hwndParent As Long)

Dim desc As String

#If WSAPI_CREAR_LABEL Then
hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
#Else
hWndMsg = hwndParent
#End If 'WSAPI_CREAR_LABEL

OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

Call StartWinsock(desc)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LimpiaWsApi
' DateTime  : 18/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub LimpiaWsApi(ByVal hWnd As Long)

If WSAStartedUp Then
    Call EndWinsock
End If

If OldWProc <> 0 Then
    SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
    OldWProc = 0
End If

#If WSAPI_CREAR_LABEL Then
If hWndMsg <> 0 Then
    DestroyWindow hWndMsg
End If
#End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : BuscaSlotSock
' DateTime  : 18/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'CSEH: Nada
Public Function BuscaSlotSock(ByVal s As Long, Optional ByVal CacheInd As Boolean = False) As Long

On Error GoTo hayError

BuscaSlotSock = WSAPISock2Usr.Item(CStr(s))

Exit Function

hayError:
BuscaSlotSock = -1
End Function

'---------------------------------------------------------------------------------------
' Procedure : AgregaSlotSock
' DateTime  : 18/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal slot As Long)

 If WSAPISock2Usr.Count > MaxUsers Then
    Call CloseSocket(slot)
     Exit Sub
End If

WSAPISock2Usr.Add CStr(slot), CStr(Sock)

End Sub


' Remueve la relacion entre el Socket y el Index
Public Sub BorraSlotSock(ByVal Sock As Long, Optional ByVal CacheIndice As Long)
WSAPISock2Usr.Remove CStr(Sock)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : WndProc
' DateTime  : 18/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim Ret As Long
Dim Tmp As String

Dim s As Long, E As Long
Dim N As Integer
    
Dim UltError As Long

WndProc = 0

Select Case msg
Case 1025

    s = wParam
    E = WSAGetSelectEvent(lParam)

    
    Select Case E
    
    Case FD_ACCEPT
        
        Debug.Print "Acepto una nueva conexion " & s
               
        If s = SockListen Then
            Call EventoSockAccept(s)
        End If
        
    Case FD_READ
        
        N = BuscaSlotSock(s)
                
        If N < 0 And s <> SockListen Then
            Call WSApiCloseSocket(s)
            Exit Function
        End If
        
        '4k de buffer
        'buffer externo
        Tmp = Space$(SIZE_RCVBUF)   'si cambias este valor, tambien hacelo mas abajo
                            'donde dice ret = 8192 :)
        
        Ret = recv(s, Tmp, Len(Tmp), 0)
        ' Comparo por = 0 ya que esto es cuando se cierra
        ' "gracefully". (mas abajo)
        
        ' Si se produce un error NO hay q llamar a CloseSocket() directamente,
        ' Ya que pueden abusar de algun error para desconectarse sin los 10segs
                
        If Ret < 0 Then
            UltError = Err.LastDllError
            If UltError = WSAEMSGSIZE Then
                Debug.Print "WSAEMSGSIZE"
                Ret = SIZE_RCVBUF
            Else
                Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)

                Call CierreForzadoPorDesconexion(N)

                Exit Function
            End If
        ElseIf Ret = 0 Then
            Call CierreForzadoPorDesconexion(N)
            Exit Function
        End If
        
        Tmp = Left(Tmp, Ret)
               
        Call EventoSockRead(N, Tmp)
        
    Case FD_CLOSE
        N = BuscaSlotSock(s)
        If s <> SockListen Then Call apiclosesocket(s)
              
        If N > 0 Then
            Call CierreForzadoPorDesconexion(N)
        End If
        
    End Select
Case Else
    WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
End Select

End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
'---------------------------------------------------------------------------------------
' Procedure : WsApiEnviar
' DateTime  : 18/02/2007 19:50
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function WsApiEnviar(ByVal slot As Integer, ByVal str As String) As Long
 
Dim Ret As String
Dim UltError As Long
Dim Retorno As Long
Dim i As Integer

Retorno = 0

' Sólo enviamo si tiene un socket valido
If Not UserList(slot).ConnID = INVALID_SOCKET Then

    Ret = send(ByVal UserList(slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)

    If Ret < 0 Then
        UltError = Err.LastDllError
        Retorno = UltError
    End If
End If

WsApiEnviar = Retorno

End Function

Public Sub EventoSockAccept(ByVal SockID As Long)
'==========================================================
'USO DE LA API DE WINSOCK
'========================
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long
    Dim NuevoSock As Long
    Dim i As Long
    Dim sa As sockaddr

    Tam = sockaddr_size
    '=============================================
    'SockID es en este caso es el socket de escucha,
    'a diferencia de socketwrench que es el nuevo
    'socket de la nueva conn
    
    'Modificado por Maraxus
    Ret = accept(SockID, sa, Tam)

    If Ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    
    NuevoSock = Ret
    
    'Seteamos el tamaño del buffer de salida
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
    End If

    If False Then
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = obtenerUserIndexLibre ' Nuevo indice
   
    If Not NewIndex = -1 Then
                
        UserList(NewIndex).ConnID = NuevoSock
        UserList(NewIndex).InicioConexion = GetTickCount()
        
        UserList(NewIndex).PacketNumber = 1
        UserList(NewIndex).MinPacketNumber = 1
        
        UserList(NewIndex).FechaIngreso = Now
        
        ' Relaciones Socket y UserIndex
        Call AgregaSlotSock(NuevoSock, NewIndex)
    Else
        ' Cierro Socket de una
        Call WSApiCloseSocket(NuevoSock)
        Call Admin.servidorComienzaAtaque
    End If
End Sub

Public Sub EventoSockRead(ByVal slot As Integer, ByRef datos As String)
Dim LastPos As Long
Dim RD As String
Dim bytesEnCola As Integer
Dim longitud As Integer

TCPESStats.BitesRecibidosMinuto = TCPESStats.BitesRecibidosMinuto + LenB(datos) * 8
TCPESStats.PaquetesRecibidosMinuto = TCPESStats.PaquetesRecibidosMinuto + 1

If slot < 1 Then Exit Sub
    If UserList(slot).ConnID Then
        RD = UserList(slot).RDBuffer & datos
        bytesEnCola = Len(RD)
      
        LastPos = 1
    
        'Explicado en el cliente
        Do
            
            longitud = Asc(mid$(RD, LastPos, 1)) + 1
                  
            If longitud = 256 Then
                If bytesEnCola - LastPos > 257 Then
                    longitud = STILong(RD, LastPos + 1) + longitud
                    LastPos = LastPos + 2
                Else
                    Exit Do
                End If
            End If
            
            'Tengo el paquete completo?
            If bytesEnCola - LastPos >= longitud Then
                Call sHandleData(mid$(RD, LastPos + 1, longitud), slot)
                LastPos = LastPos + longitud + 1
            Else
                If longitud > 255 Then LastPos = LastPos - 2
                Exit Do
            End If
            
        Loop Until bytesEnCola < LastPos

        UserList(slot).RDBuffer = Right$(RD, Len(RD) - LastPos + 1)
End If

End Sub

' Esta funcion se ejecuta cuando se produce un corte de la conexión
' entre el usuario y el personaje
Public Sub CierreForzadoPorDesconexion(ByVal slot As Integer)

    ' Liberamos el Socket Ya
    Call CloseSocketSL(UserList(slot))
    
    'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
    If Not UserList(slot).flags.UserLogged Then
        ' Si no está jugando con ningún personaje, libero ya el Slot
        If Not CloseSocket(slot) Then Call LogError("cierre forzado")
    Else
        ' Voy a encolar el cierre de este personaje
        Call Cerrar_Usuario_Forzadamente(UserList(slot))
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WSApiReiniciarSockets
' DateTime  : 18/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub WSApiReiniciarSockets()
    Dim i As Long

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Cierra todas las conexiones
    For i = 1 To MaxUsers
        If Not UserList(i).ConnID = INVALID_SOCKET Then
            Call CloseSocket(i)
        End If
    Next i
    
    ' No 'ta el PRESERVE :p
    ReDim UserList(1 To MaxUsers)
    
    For i = 1 To MaxUsers
        UserList(i).ConnID = INVALID_SOCKET
        UserList(i).InicioConexion = 0
        UserList(i).ConfirmacionConexion = 0
    Next i
    
    LastUser = 1
    NumUsers = 0
    NumUsersPremium = 0
    
    Call LimpiaWsApi(frmMain.hWnd)
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WSApiCloseSocket
' DateTime  : 18/02/2007 19:50
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub WSApiCloseSocket(ByVal Socket As Long)

Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
Call ShutDown(Socket, SD_BOTH)

End Sub
