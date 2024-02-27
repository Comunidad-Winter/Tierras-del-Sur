Attribute VB_Name = "WSKSOCK"
'date stamp: sept 1, 1996 (for version control, please don't remove)

'Visual Basic 4.0 Winsock "Header"
'   Alot of the information contained inside this file was originally
'   obtained from ALT.WINSOCK.PROGRAMMING and most of it has since been
'   modified in some way.
'
'Disclaimer: This file is public domain, updated periodically by
'   Topaz, SigSegV@mail.utexas.edu, Use it at your own risk.
'   Neither myself(Topaz) or anyone related to alt.programming.winsock
'   may be held liable for its use, or misuse.
'
'Declare check Aug 27, 1996. (Topaz, SigSegV@mail.utexas.edu)
'   All 16 bit declarations appear correct, even the odd ones that
'   pass longs inplace of in_addr and char buffers. 32 bit functions
'   also appear correct. Some are declared to return integers instead of
'   longs (breaking MS's rules.) however after testing these functions I
'   have come to the conclusion that they do not work properly when declared
'   following MS's rules.
'
'NOTES:
'   (1) I have never used WS_SELECT (select), therefore I must warn that I do
'       not know if fd_set and timeval are properly defined.
'   (2) Alot of the functions are declared with "buf as any", when calling these
'       functions you may either pass strings, byte arrays or UDT's. For 32bit I
'       I recommend Byte arrays and the use of memcopy to copy the data back out
'   (3) The async functions (wsaAsync*) require the use of a message hook or
'       message window control to capture messages sent by the winsock stack. This
'       is not to be confused with a CallBack control, The only function that uses
'       callbacks is WSASetBlockingHook()
'   (4) Alot of "helper" functions are provided in the file for various things
'       before attempting to figure out how to call a function, look and see if
'       there is already a helper function for it.
'   (5) Data types (hostent etc) have kept there 16bit definitions, even under 32bit
'       windows due to the problem of them not working when redfined following the
'       suggested rules.
Option Explicit

Public Const FD_SETSIZE = 64
Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const hostent_size = 16

Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type
Public Const servent_size = 14

Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type
Public Const protoent_size = 10

Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17

Public Const INADDR_NONE = &HFFFFFFFF
Public Const INADDR_ANY = &H0

Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Public Const sockaddr_size = 16
Public saZero As sockaddr


Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

'Agregado por Maraxus
Type WSABUF
    dwBufferLen As Long
    lpBuffer    As Long
End Type

'Agregado por Maraxus
Type FLOWSPEC
    TokenRate           As Long     'In Bytes/sec
    TokenBucketSize     As Long     'In Bytes
    PeakBandwidth       As Long     'In Bytes/sec
    Latency             As Long     'In microseconds
    DelayVariation      As Long     'In microseconds
    ServiceType         As Integer  'Guaranteed, Predictive,
                                    'Best Effort, etc.
    MaxSduSize          As Long     'In Bytes
    MinimumPolicedSize  As Long     'In Bytes
End Type

'Agregado por Maraxus
Public Const WSA_FLAG_OVERLAPPED = &H1

'Agregados por Maraxus
Public Const CF_ACCEPT = &H0
Public Const CF_REJECT = &H1

Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1

Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2

Public Const MAXGETHOSTSTRUCT = 1024

Public Const AF_INET = 2
Public Const PF_INET = 2

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type
' Windows Sockets definitions of regular Microsoft C error constants
Global Const WSAEINTR = 10004
Global Const WSAEBADF = 10009
Global Const WSAEACCES = 10013
Global Const WSAEFAULT = 10014
Global Const WSAEINVAL = 10022
Global Const WSAEMFILE = 10024
' Windows Sockets definitions of regular Berkeley error constants
Global Const WSAEWOULDBLOCK = 10035
Global Const WSAEINPROGRESS = 10036
Global Const WSAEALREADY = 10037
Global Const WSAENOTSOCK = 10038
Global Const WSAEDESTADDRREQ = 10039
Global Const WSAEMSGSIZE = 10040
Global Const WSAEPROTOTYPE = 10041
Global Const WSAENOPROTOOPT = 10042
Global Const WSAEPROTONOSUPPORT = 10043
Global Const WSAESOCKTNOSUPPORT = 10044
Global Const WSAEOPNOTSUPP = 10045
Global Const WSAEPFNOSUPPORT = 10046
Global Const WSAEAFNOSUPPORT = 10047
Global Const WSAEADDRINUSE = 10048
Global Const WSAEADDRNOTAVAIL = 10049
Global Const WSAENETDOWN = 10050
Global Const WSAENETUNREACH = 10051
Global Const WSAENETRESET = 10052
Global Const WSAECONNABORTED = 10053
Global Const WSAECONNRESET = 10054
Global Const WSAENOBUFS = 10055
Global Const WSAEISCONN = 10056
Global Const WSAENOTCONN = 10057
Global Const WSAESHUTDOWN = 10058
Global Const WSAETOOMANYREFS = 10059
Global Const WSAETIMEDOUT = 10060
Global Const WSAECONNREFUSED = 10061
Global Const WSAELOOP = 10062
Global Const WSAENAMETOOLONG = 10063
Global Const WSAEHOSTDOWN = 10064
Global Const WSAEHOSTUNREACH = 10065
Global Const WSAENOTEMPTY = 10066
Global Const WSAEPROCLIM = 10067
Global Const WSAEUSERS = 10068
Global Const WSAEDQUOT = 10069
Global Const WSAESTALE = 10070
Global Const WSAEREMOTE = 10071
' Extended Windows Sockets error constant definitions
Global Const WSASYSNOTREADY = 10091
Global Const WSAVERNOTSUPPORTED = 10092
Global Const WSANOTINITIALISED = 10093
Global Const WSAHOST_NOT_FOUND = 11001
Global Const WSATRY_AGAIN = 11002
Global Const WSANO_RECOVERY = 11003
Global Const WSANO_DATA = 11004
Global Const WSANO_ADDRESS = 11004
'---ioctl Constants
    Public Const FIONREAD = &H8004667F
    Public Const FIONBIO = &H8004667E
    Public Const FIOASYNC = &H8004667D

#If Win16 Then
'---Windows System functions
    Public Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
    Public Declare Sub MemCopy Lib "Kernel" Alias "hmemcpy" (Dest As Any, Src As Any, ByVal cb&)
    Public Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer
'---async notification constants
    Public Const SOL_SOCKET = &HFFFF
    Public Const SO_LINGER = &H80
    Public Const SO_RCVBUFFER = &H1002              ' Agregado por Maraxus
    Public Const SO_SNDBUFFER = &H1001              ' Agregado por Maraxus
    Public Const SO_CONDITIONAL_ACCEPT = &H3002    ' Agregado por Maraxus
    Public Const FD_READ = &H1
    Public Const FD_WRITE = &H2
    Public Const FD_OOB = &H4
    Public Const FD_ACCEPT = &H8
    Public Const FD_CONNECT = &H10
    Public Const FD_CLOSE = &H20
'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "ws2_32.DLL" (ByVal s As Integer, addr As sockaddr, AddrLen As Integer) As Integer
    Public Declare Function bind Lib "ws2_32.DLL" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Public Declare Function apiclosesocket Lib "ws2_32.DLL" Alias "closesocket" (ByVal s As Integer) As Integer
    Public Declare Function connect Lib "ws2_32.DLL" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Public Declare Function ioctlsocket Lib "ws2_32.DLL" (ByVal s As Integer, ByVal Cmd As Long, argp As Long) As Integer
    Public Declare Function getpeername Lib "ws2_32.DLL" (ByVal s As Integer, sName As sockaddr, namelen As Integer) As Integer
    Public Declare Function getsockname Lib "ws2_32.DLL" (ByVal s As Integer, sName As sockaddr, namelen As Integer) As Integer
    Public Declare Function getsockopt Lib "ws2_32.DLL" (ByVal s As Integer, ByVal level As Integer, ByVal optname As Integer, optval As Any, optlen As Integer) As Integer
    Public Declare Function htonl Lib "ws2_32.DLL" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "ws2_32.DLL" (ByVal hostshort As Integer) As Integer
    Public Declare Function inet_addr Lib "ws2_32.DLL" (ByVal cp As String) As Long
    Public Declare Function inet_ntoa Lib "ws2_32.DLL" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "ws2_32.DLL" (ByVal s As Integer, ByVal backlog As Integer) As Integer
    Public Declare Function ntohl Lib "ws2_32.DLL" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "ws2_32.DLL" (ByVal netshort As Integer) As Integer
    Public Declare Function recv Lib "ws2_32.DLL" (ByVal s As Integer, ByVal buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
    Public Declare Function recvfrom Lib "ws2_32.DLL" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer, from As sockaddr, fromlen As Integer) As Integer
    Public Declare Function ws_select Lib "ws2_32.DLL" Alias "select" (ByVal nfds As Integer, readfds As Any, writefds As Any, exceptfds As Any, timeout As timeval) As Integer
    Public Declare Function send Lib "ws2_32.DLL" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
    Public Declare Function sendto Lib "ws2_32.DLL" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer, to_addr As sockaddr, ByVal tolen As Integer) As Integer
    Public Declare Function setsockopt Lib "ws2_32.DLL" (ByVal s As Integer, ByVal level As Integer, ByVal optname As Integer, optval As Any, ByVal optlen As Integer) As Integer
    Public Declare Function ShutDown Lib "ws2_32.DLL" Alias "shutdown" (ByVal s As Integer, ByVal how As Integer) As Integer
    Public Declare Function Socket Lib "ws2_32.DLL" Alias "socket" (ByVal af As Integer, ByVal s_type As Integer, ByVal Protocol As Integer) As Integer
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "ws2_32.DLL" (addr As Long, ByVal addr_len As Integer, ByVal addr_type As Integer) As Long
    Public Declare Function gethostbyname Lib "ws2_32.DLL" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "ws2_32.DLL" (ByVal host_name As String, ByVal namelen As Integer) As Integer
    Public Declare Function getservbyport Lib "ws2_32.DLL" (ByVal Port As Integer, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "ws2_32.DLL" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "ws2_32.DLL" (ByVal proto As Integer) As Long
    Public Declare Function getprotobyname Lib "ws2_32.DLL" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "ws2_32.DLL" (ByVal wVR As Integer, lpWSAD As WSADataType) As Integer
    Public Declare Function WSACleanup Lib "ws2_32.DLL" () As Integer
    Public Declare Sub WSASetLastError Lib "ws2_32.DLL" (ByVal iError As Integer)
    Public Declare Function WSAGetLastError Lib "ws2_32.DLL" () As Integer
    Public Declare Function WSAIsBlocking Lib "ws2_32.DLL" () As Integer
    Public Declare Function WSAUnhookBlockingHook Lib "ws2_32.DLL" () As Integer
    Public Declare Function WSASetBlockingHook Lib "ws2_32.DLL" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "ws2_32.DLL" () As Integer
    Public Declare Function WSAAsyncGetServByName Lib "ws2_32.DLL" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetServByPort Lib "ws2_32.DLL" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal Port As Integer, ByVal proto As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetProtoByName Lib "ws2_32.DLL" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal proto_name As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetProtoByNumber Lib "ws2_32.DLL" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal Number As Integer, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.DLL" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal host_name As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.DLL" (ByVal hWnd As Integer, ByVal wMsg As Integer, addr As Long, ByVal addr_len As Integer, ByVal addr_type As Integer, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSACancelAsyncRequest Lib "ws2_32.DLL" (ByVal hAsyncTaskHandle As Integer) As Integer
    Public Declare Function WSAAsyncSelect Lib "ws2_32.DLL" (ByVal s As Integer, ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal lEvent As Long) As Integer
    Public Declare Function WSARecvEx Lib "ws2_32.DLL" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
'Agregado por Maraxus
    Declare Function WSAAccept Lib "ws2_32.DLL" (ByVal s As Integer, pSockAddr As sockaddr, AddrLen As Integer, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As Integer
    
    Public Const SOMAXCONN As Integer = &H7FFF            ' Agregado por Maraxus

#ElseIf Win32 Then
'---Windows System Functions
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'---async notification constants
    Public Const SOL_SOCKET = &HFFFF&
    Public Const SO_LINGER = &H80&
    Public Const SO_RCVBUFFER = &H1002&             ' Agregado por Maraxus
    Public Const SO_SNDBUFFER = &H1001&              ' Agregado por Maraxus
    'Public Const SO_RCVBUFFER = &H2800&             ' Agregado por Maraxus
    'Public Const SO_SNDBUFFER = &H2800&              ' Agregado por Maraxus
    Public Const SO_CONDITIONAL_ACCEPT = &H3002&    ' Agregado por Maraxus
    Public Const FD_READ = &H1&
    Public Const FD_WRITE = &H2&
    Public Const FD_OOB = &H4&
    Public Const FD_ACCEPT = &H8&
    Public Const FD_CONNECT = &H10&
    Public Const FD_CLOSE = &H20&
'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, AddrLen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function apiclosesocket Lib "wsock32.dll" Alias "closesocket" (ByVal s As Long) As Long
    Public Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal s As Long, ByVal Cmd As Long, argp As Long) As Long
    Public Declare Function getpeername Lib "wsock32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockname Lib "wsock32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function recvfrom Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
    Public Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Public Declare Function send Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function sendto Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
    Public Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Public Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
    Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Public Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Public Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Public Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
    Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Number As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
'Agregado por Maraxus
    Declare Function WSAAccept Lib "ws2_32.DLL" (ByVal s As Long, pSockAddr As sockaddr, AddrLen As Long, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As Long
    Public Const SOMAXCONN As Long = &H7FFFFFFF            ' Agregado por Maraxus


#End If


'SOME STUFF I ADDED
Public MySocket%
Public SockReadBuffer$
Public Const WSA_NoName = "Unknown"
Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled


'---------------------------------------------------------------------------------------
' Procedure : WSAGetSelectEvent
' DateTime  : 18/02/2007 19:45
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If
End Function



'---------------------------------------------------------------------------------------
' Procedure : WSAGetAsyncError
' DateTime  : 18/02/2007 19:45
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

'---------------------------------------------------------------------------------------
' Procedure : EndWinsock
' DateTime  : 18/02/2007 19:46
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EndWinsock()
    Dim Ret&

    If WSAIsBlocking() Then
        Ret = WSACancelBlockingCall()
    End If
    
    Ret = WSACleanup()
    WSAStartedUp = False

End Sub

#If Win16 Then
    Function GetProtocolByName(ByVal Protocol$) As Integer
    Dim tmpShort%
#ElseIf Win32 Then
    Function GetProtocolByName(ByVal Protocol$) As Long
    Dim tmpShort&
#End If
    Dim ppe&
    Dim peDestProt As protoent
    ppe = getprotobyname(Protocol)
    If ppe Then
        MemCopy peDestProt, ByVal ppe, protoent_size
        GetProtocolByName = peDestProt.p_proto
    Else
        tmpShort = val(Protocol)
        If tmpShort Then
            GetProtocolByName = htons(tmpShort)
        Else
            GetProtocolByName = SOCKET_ERROR
        End If
    End If
End Function

'this function should work on 16 and 32 bit systems
'---------------------------------------------------------------------------------------
' Procedure : GetWSAErrorString
' DateTime  : 18/02/2007 19:47
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function GetWSAErrorString(ByVal errnum&) As String

    Select Case errnum
        Case 10004: GetWSAErrorString = "Interrupted system call."
        Case 10009: GetWSAErrorString = "Bad file number."
        Case 10013: GetWSAErrorString = "Permission Denied."
        Case 10014: GetWSAErrorString = "Bad Address."
        Case 10022: GetWSAErrorString = "Invalid Argument."
        Case 10024: GetWSAErrorString = "Too many open files."
        Case 10035: GetWSAErrorString = "Operation would block."
        Case 10036: GetWSAErrorString = "Operation now in progress."
        Case 10037: GetWSAErrorString = "Operation already in progress."
        Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
        Case 10039: GetWSAErrorString = "Destination address required."
        Case 10040: GetWSAErrorString = "Message too long."
        Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
        Case 10042: GetWSAErrorString = "Protocol not available."
        Case 10043: GetWSAErrorString = "Protocol not supported."
        Case 10044: GetWSAErrorString = "Socket type not supported."
        Case 10045: GetWSAErrorString = "Operation not supported on socket."
        Case 10046: GetWSAErrorString = "Protocol family not supported."
        Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
        Case 10048: GetWSAErrorString = "Address already in use."
        Case 10049: GetWSAErrorString = "Can't assign requested address."
        Case 10050: GetWSAErrorString = "Network is down."
        Case 10051: GetWSAErrorString = "Network is unreachable."
        Case 10052: GetWSAErrorString = "Network dropped connection."
        Case 10053: GetWSAErrorString = "Software caused connection abort."
        Case 10054: GetWSAErrorString = "Connection reset by peer."
        Case 10055: GetWSAErrorString = "No buffer space available."
        Case 10056: GetWSAErrorString = "Socket is already connected."
        Case 10057: GetWSAErrorString = "Socket is not connected."
        Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
        Case 10059: GetWSAErrorString = "Too many references: can't splice."
        Case 10060: GetWSAErrorString = "Connection timed out."
        Case 10061: GetWSAErrorString = "Connection refused."
        Case 10062: GetWSAErrorString = "Too many levels of symbolic links."
        Case 10063: GetWSAErrorString = "File name too long."
        Case 10064: GetWSAErrorString = "Host is down."
        Case 10065: GetWSAErrorString = "No route to host."
        Case 10066: GetWSAErrorString = "Directory not empty."
        Case 10067: GetWSAErrorString = "Too many processes."
        Case 10068: GetWSAErrorString = "Too many users."
        Case 10069: GetWSAErrorString = "Disk quota exceeded."
        Case 10070: GetWSAErrorString = "Stale NFS file handle."
        Case 10071: GetWSAErrorString = "Too many levels of remote in path."
        Case 10091: GetWSAErrorString = "Network subsystem is unusable."
        Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
        Case 10093: GetWSAErrorString = "Winsock not initialized."
        Case 10101: GetWSAErrorString = "Disconnect."
        Case 11001: GetWSAErrorString = "Host not found."
        Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
        Case 11003: GetWSAErrorString = "Nonrecoverable error."
        Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
        Case Else:
    End Select

End Function


'this function should work on 16 and 32 bit systems
#If Win16 Then
'---------------------------------------------------------------------------------------
' Procedure : ListenForConnect
' DateTime  : 18/02/2007 19:48
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ListenForConnect(ByVal Port%, ByVal HWndToMsg%, ByVal Enlazar As String) As Integer
    Dim s%, dummy%
    Dim SelectOps%

#ElseIf Win32 Then
Public Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&, ByVal Enlazar As String) As Long
    Dim s&, dummy&
    Dim SelectOps&
#End If
    Dim sockin As sockaddr
    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    If Enlazar = "" Then
        sockin.sin_addr = htonl(INADDR_ANY)
    Else
        sockin.sin_addr = inet_addr(Enlazar)
    End If
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    s = Socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    
'Agregado por Maraxus
    'If setsockopt(s, SOL_SOCKET, SO_CONDITIONAL_ACCEPT, True, 2) Then
    '    LogApiSock ("Error seteando conditional accept")
    '    Debug.Print "Error seteando conditional accept"
    'Else
    '    LogApiSock ("Conditional accept seteado")
    '    Debug.Print "Conditional accept seteado ^^"
    'End If
    
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            dummy = apiclosesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
'    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    SelectOps = FD_READ Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
        If s > 0 Then
            dummy = apiclosesocket(s)
        End If
        ListenForConnect = SOCKET_ERROR
        Exit Function
    End If
    
    'If listen(s, 5) Then
    If listen(s, SOMAXCONN) Then
        If s > 0 Then
            dummy = apiclosesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    ListenForConnect = s

End Function

'---------------------------------------------------------------------------------------
' Procedure : StartWinsock
' DateTime  : 18/02/2007 19:48
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function StartWinsock(sDescription As String) As Boolean
    Dim StartupData As WSADataType

    If Not WSAStartedUp Then
        'If Not WSAStartup(&H101, StartupData) Then
        If Not WSAStartup(&H202, StartupData) Then  'Use sockets v2.2 instead of 1.1 (Maraxus)
            WSAStartedUp = True
'            Debug.Print "wVersion="; StartupData.wVersion, "wHighVersion="; StartupData.wHighVersion
'            Debug.Print "If wVersion == 257 then everything is kewl"
'            Debug.Print "szDescription="; StartupData.szDescription
'            Debug.Print "szSystemStatus="; StartupData.szSystemStatus
'            Debug.Print "iMaxSockets="; StartupData.iMaxSockets, "iMaxUdpDg="; StartupData.iMaxUdpDg
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp

End Function

