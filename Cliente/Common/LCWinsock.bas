Attribute VB_Name = "CDMLCWinsock"
'Date stamp: sept 1, 1996 (for version control, please don't remove)
'********************
'Modifications, improvements and additions © 2001-2004 by Luis Cantero
'Modifications: ListenForConnect, ConnectSock, SendData, StartWinsock, IsConnected, IsSocketReady
'16-Bit declarations removed.
'Additions: Ping, GetSMTPserver, MyIP, HTTPRequest, GetDNSInfo, etc.
'L.C. Enterprises - http://LCen.com
'********************

'Visual Basic 6.0 Winsock "Header"
'   Alot of the information contained inside this file was originally
'   obtained from ALT.WINSOCK.PROGRAMMING and most of it has since been
'   modified in some way.
'
'Disclaimer: This file is public domain, updated periodically by
'   Topaz, SigSegV@mail.utexas.edu, Use it at your own risk.
'   Neither myself(Topaz) or anyone related to alt.programming.winsock
'   may be held liable for its use, or misuse.
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

Public Const hostent_size = 16
Type HostEnt

    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Const servent_size = 14
Type servent

    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type

Public Const protoent_size = 10
Type protoent

    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type

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
Public Const WSAEINTR = 10004
Public Const WSAEBADF = 10009
Public Const WSAEACCES = 10013
Public Const WSAEFAULT = 10014
Public Const WSAEINVAL = 10022
Public Const WSAEMFILE = 10024
' Windows Sockets definitions of regular Berkeley error constants
Public Const WSAEWOULDBLOCK = 10035
Public Const WSAEINPROGRESS = 10036
Public Const WSAEALREADY = 10037
Public Const WSAENOTSOCK = 10038
Public Const WSAEDESTADDRREQ = 10039
Public Const WSAEMSGSIZE = 10040
Public Const WSAEPROTOTYPE = 10041
Public Const WSAENOPROTOOPT = 10042
Public Const WSAEPROTONOSUPPORT = 10043
Public Const WSAESOCKTNOSUPPORT = 10044
Public Const WSAEOPNOTSUPP = 10045
Public Const WSAEPFNOSUPPORT = 10046
Public Const WSAEAFNOSUPPORT = 10047
Public Const WSAEADDRINUSE = 10048
Public Const WSAEADDRNOTAVAIL = 10049
Public Const WSAENETDOWN = 10050
Public Const WSAENETUNREACH = 10051
Public Const WSAENETRESET = 10052
Public Const WSAECONNABORTED = 10053
Public Const WSAECONNRESET = 10054
Public Const WSAENOBUFS = 10055
Public Const WSAEISCONN = 10056
Public Const WSAENOTCONN = 10057
Public Const WSAESHUTDOWN = 10058
Public Const WSAETOOMANYREFS = 10059
Public Const WSAETIMEDOUT = 10060
Public Const WSAECONNREFUSED = 10061
Public Const WSAELOOP = 10062
Public Const WSAENAMETOOLONG = 10063
Public Const WSAEHOSTDOWN = 10064
Public Const WSAEHOSTUNREACH = 10065
Public Const WSAENOTEMPTY = 10066
Public Const WSAEPROCLIM = 10067
Public Const WSAEUSERS = 10068
Public Const WSAEDQUOT = 10069
Public Const WSAESTALE = 10070
Public Const WSAEREMOTE = 10071
' Extended Windows Sockets error constant definitions
Public Const WSASYSNOTREADY = 10091
Public Const WSAVERNOTSUPPORTED = 10092
Public Const WSANOTINITIALISED = 10093
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004
Public Const WSANO_ADDRESS = 11004
'---ioctl Constants
Public Const FIONREAD = &H8004667F
Public Const FIONBIO = &H8004667E
Public Const FIOASYNC = &H8004667D

'---Windows System Functions
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'---async notification constants
Public Const SOL_SOCKET = &HFFFF&
Public Const SO_LINGER = &H80&
Public Const FD_READ = &H1&
Public Const FD_WRITE = &H2&
Public Const FD_OOB = &H4&
Public Const FD_ACCEPT = &H8&
Public Const FD_CONNECT = &H10&
Public Const FD_CLOSE = &H20&
'---SOCKET FUNCTIONS
Public Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, addr As sockaddr, addrLen As Long) As Long
Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
Public Declare Function getsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
Public Declare Function ws_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, TimeOut As timeval) As Long
Public Declare Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function ShutDown Lib "ws2_32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'---DATABASE FUNCTIONS
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Long, ByVal proto As String) As Long
Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
Public Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Sub WSASetLastError Lib "ws2_32.dll" (ByVal iError As Long)
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSAIsBlocking Lib "ws2_32.dll" () As Long
Public Declare Function WSAUnhookBlockingHook Lib "ws2_32.dll" () As Long
Public Declare Function WSASetBlockingHook Lib "ws2_32.dll" (ByVal lpBlockFunc As Long) As Long
Public Declare Function WSACancelBlockingCall Lib "ws2_32.dll" () As Long
Public Declare Function WSAAsyncGetServByName Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetServByPort Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByName Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByNumber Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal number As Long, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSACancelAsyncRequest Lib "ws2_32.dll" (ByVal hAsyncTaskHandle As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSARecvEx Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long

'SOME STUFF I ADDED
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

Public MySocket%
Public SockReadBuffer$
Public Const WSA_NoName = "Unknown"
Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled
'Ping
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long  'formerly integer
    'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type
Private Const PING_TIMEOUT = 200
Private Declare Function IcmpCreateFile Lib "Icmp.dll" () As Long
Private Declare Function IcmpSendEcho Lib "Icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function IcmpCloseHandle Lib "Icmp.dll" (ByVal IcmpHandle As Long) As Long
'DNSInfo
Private Type IP_ADDRESS_STRING
    IpAddressString(4 * 4 - 1) As Byte
End Type

Private Type IP_MASK_STRING
    IpMaskString(4 * 4 - 1) As Byte
End Type

Private Type IP_ADDR_STRING
Next      As Long
IpAddress As IP_ADDRESS_STRING
IpMask    As IP_MASK_STRING
Context   As Long
End Type

Private Const MAX_HOSTNAME_LEN = 128
Private Const MAX_DOMAIN_NAME_LEN = 128
Private Const MAX_SCOPE_ID_LEN = 256

Private Type FIXED_INFO
    hostname(MAX_HOSTNAME_LEN + 4 - 1) As Byte
    DomainName(MAX_DOMAIN_NAME_LEN + 4 - 1) As Byte
    CurrentDnsServer As Long
    DnsServerList    As IP_ADDR_STRING
    NodeType         As Long
    ScopeId(MAX_SCOPE_ID_LEN + 4 - 1) As Byte
    EnableRouting    As Long
    EnableProxy      As Long
    EnableDns        As Long
End Type

Public Const ERROR_BUFFER_OVERFLOW = 111

Private Declare Function GetNetworkParams Lib "iphlpapi.dll" (pFixedInfo As Any, pOutBufLen As Long) As Long
'For Wait Routines
Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Function AddrToIP(ByVal AddrOrIP$) As String

    AddrToIP$ = getascip(GetHostByNameAlias(AddrOrIP$))

End Function

Function ConnectSock(ByVal Host$, ByVal Port&, ByVal HWndToMsg&) As Long

  Dim intSocket&, SelectOps&
  Dim sockin As sockaddr

    'Start Winsock
    Call StartWinsock

    SockReadBuffer$ = ""
    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Get Address and check if valid
    sockin.sin_addr = GetHostByNameAlias(Host$)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Create socket and check if OK
    intSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If intSocket < 0 Then
        ConnectSock = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Set Linger
    If SetSockLinger(intSocket, 1, 0) = SOCKET_ERROR Then
        If intSocket > 0 Then
            Call closesocket(intSocket)
        End If
        ConnectSock = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Connect
    If Connect(intSocket, sockin, sockaddr_size) <> 0 Then
        If intSocket > 0 Then
            Call closesocket(intSocket)
        End If
        ConnectSock = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Set receive Window
    SelectOps = FD_READ Or FD_CLOSE
    If WSAAsyncSelect(intSocket, HWndToMsg, ByVal &H202, ByVal SelectOps) Then '&H202 is the MouseUp Event
        If intSocket > 0 Then
            Call closesocket(intSocket)
        End If
        ConnectSock = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Return
    ConnectSock = intSocket

End Function

Public Sub EndWinsock()

    If WSAIsBlocking() Then
        Call WSACancelBlockingCall
    End If

    WSAStartedUp = False

End Sub

Public Function getascip(ByVal inn As Long) As String

  Dim nStr&
  Dim lpStr&
  Dim retString$

    retString = String$(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
        retString = Left$(retString, nStr)
        getascip = retString
      Else
        getascip = "255.255.255.255"
    End If

End Function

'Returns address of DNS server used by local machine
Function GetDNSInfo() As String

  Dim sFinalBuff              As String
  Dim lngFixedInfoNeeded      As Long
  Dim bytFixedInfoBuffer()    As Byte
  Dim udtFixedInfo            As FIXED_INFO
  Dim lngIpAddrStringPtr      As Long
  Dim udtIpAddrString         As IP_ADDR_STRING
  Dim strDnsIpAddress         As String
  Dim lngWin32apiResultCode   As Long
  Dim lngDNSPing              As Long
  Dim lngDNSPing2             As Long
  Dim arrDNS                  As Variant
  Dim Member                  As Variant

    lngWin32apiResultCode = GetNetworkParams(ByVal vbNullString, lngFixedInfoNeeded)
    If lngWin32apiResultCode = ERROR_BUFFER_OVERFLOW Then
        ReDim bytFixedInfoBuffer(lngFixedInfoNeeded)
      Else
        GoTo TerminateGetNetworkParams
    End If

    lngWin32apiResultCode = GetNetworkParams(bytFixedInfoBuffer(0), lngFixedInfoNeeded)
    MemCopy udtFixedInfo, bytFixedInfoBuffer(0), Len(udtFixedInfo)

    With udtFixedInfo
        lngIpAddrStringPtr = VarPtr(.DnsServerList)
        Do While lngIpAddrStringPtr
            MemCopy udtIpAddrString, ByVal lngIpAddrStringPtr, Len(udtIpAddrString)
            With udtIpAddrString
                strDnsIpAddress = StrConv(.IpAddress.IpAddressString, vbUnicode)
                If sFinalBuff = vbNullString Then
                    sFinalBuff = Left$(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ","
                  Else
                    If InStr(1, sFinalBuff, Left$(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ",") = 0 Then
                        sFinalBuff = sFinalBuff & Left$(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ","
                    End If
                End If
                lngIpAddrStringPtr = .Next
            End With
        Loop
    End With

    If Right$(sFinalBuff, 1) = "," Then sFinalBuff = Left$(sFinalBuff, Len(sFinalBuff) - 1)

    arrDNS = Split(sFinalBuff, ",")

    'Compare and select fastest DNS
    For Each Member In arrDNS
        lngDNSPing2 = Ping(Member, , True)
        If (lngDNSPing2 < lngDNSPing And lngDNSPing2 > -1) Or lngDNSPing2 > 0 And lngDNSPing = 0 Then sFinalBuff = Member: lngDNSPing = lngDNSPing2
    Next Member

    'Return fastest DNS
    If InStr(1, sFinalBuff, ",") = 0 Then 'Ping failed
        GetDNSInfo = sFinalBuff
      Else 'Return first DNS in array
        GetDNSInfo = arrDNS(0)
    End If

TerminateGetNetworkParams:

End Function

Public Function GetHostByAddress(ByVal addr As Long) As String

  Dim phe&
  Dim heDestHost As HostEnt
  Dim hostname$

    phe = gethostbyaddr(addr, 4, PF_INET)
    If phe Then
        MemCopy heDestHost, ByVal phe, hostent_size
        hostname = String$(256, 0)
        MemCopy ByVal hostname, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left$(hostname, InStr(hostname, Chr$(0)) - 1)
      Else
        GetHostByAddress = WSA_NoName
    End If

End Function

'returns IP as long, in network byte order
Public Function GetHostByNameAlias(ByVal hostname$) As Long

  Dim phe&
  Dim heDestHost As HostEnt
  Dim addrList&
  Dim retIP&

    retIP = inet_addr(hostname$)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname$)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
          Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP

End Function

'returns your local machines name
Public Function GetLocalHostName() As String

  Dim sName$

    sName = String$(256, 0)

    Call gethostname(sName, 256)
    If InStr(sName, Chr$(0)) Then
        sName = Left$(sName, InStr(sName, Chr$(0)) - 1)
    End If
    If sName = "" Then sName = WSA_NoName
    GetLocalHostName = sName

End Function

Public Function GetPeerAddress(ByVal intSocket&) As String

  Dim addrLen&
  Dim sa As sockaddr

    addrLen = sockaddr_size
    If getpeername(intSocket, sa, addrLen) Then
        GetPeerAddress = ""
      Else
        GetPeerAddress = SockAddressToString(sa)
    End If

End Function

Public Function GetPortFromString(ByVal PortStr$) As Long

  'sometimes users provide ports outside the range of a VB
  'integer, so this function returns an integer for a string
  'just to keep an error from happening, it converts the
  'number to a negative if needed

    If val(PortStr$) > 32767 Then
        GetPortFromString = CInt(val(PortStr$) - &H10000)
      Else
        GetPortFromString = val(PortStr$)
    End If
    If Err Then GetPortFromString = 0

End Function

Function GetProtocolByName(ByVal protocol$) As Long

  Dim tmpShort&
  Dim ppe&
  Dim peDestProt As protoent

    ppe = getprotobyname(protocol)
    If ppe Then
        MemCopy peDestProt, ByVal ppe, protoent_size
        GetProtocolByName = peDestProt.p_proto
      Else
        tmpShort = val(protocol)
        If tmpShort Then
            GetProtocolByName = htons(tmpShort)
          Else
            GetProtocolByName = SOCKET_ERROR
        End If
    End If

End Function

Function GetServiceByName(ByVal service$, ByVal protocol$) As Long

  Dim serv&
  Dim pse&
  Dim seDestServ As servent

    pse = getservbyname(service, protocol)
    If pse Then
        MemCopy seDestServ, ByVal pse, servent_size
        GetServiceByName = seDestServ.s_port
      Else
        serv = val(service)
        If serv Then
            GetServiceByName = htons(serv)
          Else
            GetServiceByName = INVALID_SOCKET
        End If
    End If

End Function

'Gets the SMTP Server for an email address (Best MX Record)
Function GetSMTPserver(strEmail As String) As String

    On Error Resume Next
      Dim SocketBuffer As sockaddr
      Dim dnsReply(2048) As Byte
      Dim Sock As Integer
      Dim strDNS As String

      Dim strSMTP As String
      Dim intSearch As Integer
      Dim intDot1 As Integer
      Dim intDot2 As Integer
      Dim sDatagram As String
        'Dim arrMessage() As Byte 'Used with sendto()

      Dim intLength As Integer
      Dim intCurrentPos As Integer

        'Get best DNS
        strDNS = GetDNSInfo

        If strDNS = "" Then
            Exit Function
        End If

        Call StartWinsock

        Sock = socket(AF_INET, SOCK_DGRAM, 0)
        If Sock = SOCKET_ERROR Then
            Call EndWinsock
            Exit Function
        End If

        With SocketBuffer
            .sin_family = AF_INET
            .sin_port = htons(53)
            .sin_addr = GetHostByNameAlias(strDNS)
            .sin_zero = String$(8, 0)
        End With

        If Connect(Sock, SocketBuffer, sockaddr_size) = SOCKET_ERROR Then
            Call EndWinsock
            Exit Function
        End If

        strSMTP = mid$(strEmail, InStr(2, strEmail, "@") + 1)

        'Convert Server: yahoo.com -> [5]yahoo[3]com
        strSMTP = Chr$(InStr(2, strSMTP, ".") - 1) & strSMTP
        intSearch = 2
        Do
            intDot1 = InStr(intSearch, strSMTP, ".")
            intDot2 = InStr(intDot1 + 1, strSMTP, ".")

            'Indicate that there are no more dots
            If intDot2 = 0 Then intDot2 = Len(strSMTP) + 1

            'Convert dot into Char(length of next part)
            Mid$(strSMTP, intDot1, 1) = Chr$(intDot2 - intDot1 - 1)

            intSearch = intDot1 + 1
        Loop Until intDot2 = Len(strSMTP) + 1

        'Form Datagram
        sDatagram = Chr$(Int(Rnd * 255)) & Chr$(Int(Rnd * 255)) & Chr$(1) & Chr$(128)
        sDatagram = sDatagram & Chr$(0) & Chr$(1) & String$(6, 0) & strSMTP & Chr$(0)
        sDatagram = sDatagram & Chr$(0) & Chr$(15) & Chr$(0) & Chr$(255)

        'Send request
        'arrMessage = StrConv(sDatagram, vbFromUnicode)
        'If sendto(Sock, arrMessage(0), UBound(arrMessage) + 1, 0, SocketBuffer, Len(SocketBuffer)) = SOCKET_ERROR Then
        If SendData(Sock, sDatagram) = SOCKET_ERROR Then
            Call EndWinsock
            Exit Function
        End If

        'Get answer and process it
        If recvfrom(Sock, dnsReply(0), 2048, 0, SocketBuffer, sockaddr_size) > 0 Then 'Process reply

            'Convert reply to a string
            strDNS = StrConv(dnsReply(), vbUnicode)

            intCurrentPos = 13

            'Step over server's name
            intLength = Asc(mid$(strDNS, intCurrentPos, 1))
            While intLength
                'Add part's length to current position
                intCurrentPos = intCurrentPos + intLength + 1

                'Get length of next part
                intLength = Asc(mid$(strDNS, intCurrentPos, 1))
            Wend

            'Step over null (2 Bytes) + (6 Bytes)
            intCurrentPos = intCurrentPos + intLength + 2 + 6

            'Check to make sure we received an MX record
            If Asc(mid$(strDNS, intCurrentPos)) = 15 Then

                'Step over the last half of the integer that specifies the record type (1 byte)
                'Step over the RR Type, RR Class, TTL (3 integers = 6 bytes)
                'Step over the MX data length specifier (1 integer = 2 bytes)
                'Step over the MX preference value (1 integer = 2 bytes)
                intCurrentPos = intCurrentPos + 1 + 6 + 2 + 2

                'Get Mail Server's name
                intLength = Asc(mid$(strDNS, intCurrentPos, 1))
                While intLength

                    'If MX Record is compressed, 0xc0 or 192 (compression char)
                    If intLength = 192 Then
                        'Go to start of next part
                        intCurrentPos = Asc(mid$(strDNS, intCurrentPos + 1, 1)) + 1
                        'Get length of next part
                        intLength = Asc(mid$(strDNS, intCurrentPos, 1))
                    End If

                    'Parse server's name
                    GetSMTPserver = GetSMTPserver & mid$(strDNS, intCurrentPos + 1, intLength) & "."

                    'Add part's length to current position
                    intCurrentPos = intCurrentPos + 1 + intLength

                    'Get length of next part
                    intLength = Asc(mid$(strDNS, intCurrentPos, 1))

                Wend

                'Trim last dot and return
                GetSMTPserver = Left$(GetSMTPserver, Len(GetSMTPserver) - 1)

            End If
        End If

        'Clean up
        Call closesocket(Sock)
        Call EndWinsock

End Function

Function GetSockAddress(ByVal intSocket&) As String

  Dim addrLen&
  Dim sa As sockaddr
  Dim szRet$

    szRet = String$(32, 0)
    addrLen = sockaddr_size
    If getsockname(intSocket, sa, addrLen) Then
        GetSockAddress = ""
      Else
        GetSockAddress = SockAddressToString(sa)
    End If

End Function

Function GetWSAErrorString(ByVal errnum&) As String

    On Error Resume Next
        Select Case errnum
          Case 10004
            GetWSAErrorString = "Interrupted system call."
          Case 10009
            GetWSAErrorString = "Bad file number."
          Case 10013
            GetWSAErrorString = "Permission Denied."
          Case 10014
            GetWSAErrorString = "Bad Address."
          Case 10022
            GetWSAErrorString = "Invalid Argument."
          Case 10024
            GetWSAErrorString = "Too many open files."
          Case 10035
            GetWSAErrorString = "Operation would block."
          Case 10036
            GetWSAErrorString = "Operation now in progress."
          Case 10037
            GetWSAErrorString = "Operation already in progress."
          Case 10038
            GetWSAErrorString = "Socket operation on nonsocket."
          Case 10039
            GetWSAErrorString = "Destination address required."
          Case 10040
            GetWSAErrorString = "Message too long."
          Case 10041
            GetWSAErrorString = "Protocol wrong type for socket."
          Case 10042
            GetWSAErrorString = "Protocol not available."
          Case 10043
            GetWSAErrorString = "Protocol not supported."
          Case 10044
            GetWSAErrorString = "Socket type not supported."
          Case 10045
            GetWSAErrorString = "Operation not supported on socket."
          Case 10046
            GetWSAErrorString = "Protocol family not supported."
          Case 10047
            GetWSAErrorString = "Address family not supported by protocol family."
          Case 10048
            GetWSAErrorString = "Address already in use."
          Case 10049
            GetWSAErrorString = "Can't assign requested address."
          Case 10050
            GetWSAErrorString = "Network is down."
          Case 10051
            GetWSAErrorString = "Network is unreachable."
          Case 10052
            GetWSAErrorString = "Network dropped connection."
          Case 10053
            GetWSAErrorString = "Software caused connection abort."
          Case 10054
            GetWSAErrorString = "Connection reset by peer."
          Case 10055
            GetWSAErrorString = "No buffer space available."
          Case 10056
            GetWSAErrorString = "Socket is already connected."
          Case 10057
            GetWSAErrorString = "Socket is not connected."
          Case 10058
            GetWSAErrorString = "Can't send after socket shutdown."
          Case 10059
            GetWSAErrorString = "Too many references: can't splice."
          Case 10060
            GetWSAErrorString = "Connection timed out."
          Case 10061
            GetWSAErrorString = "Connection refused."
          Case 10062
            GetWSAErrorString = "Too many levels of symbolic links."
          Case 10063
            GetWSAErrorString = "File name too long."
          Case 10064
            GetWSAErrorString = "Host is down."
          Case 10065
            GetWSAErrorString = "No route to host."
          Case 10066
            GetWSAErrorString = "Directory not empty."
          Case 10067
            GetWSAErrorString = "Too many processes."
          Case 10068
            GetWSAErrorString = "Too many users."
          Case 10069
            GetWSAErrorString = "Disk quota exceeded."
          Case 10070
            GetWSAErrorString = "Stale NFS file handle."
          Case 10071
            GetWSAErrorString = "Too many levels of remote in path."
          Case 10091
            GetWSAErrorString = "Network subsystem is unusable."
          Case 10092
            GetWSAErrorString = "Winsock DLL cannot support this application."
          Case 10093
            GetWSAErrorString = "Winsock not initialized."
          Case 10101
            GetWSAErrorString = "Disconnect."
          Case 11001
            GetWSAErrorString = "Host not found."
          Case 11002
            GetWSAErrorString = "Nonauthoritative host not found."
          Case 11003
            GetWSAErrorString = "Nonrecoverable error."
          Case 11004
            GetWSAErrorString = "Valid name, no data record of requested type."
          Case Else

        End Select

End Function

Function HTTPRequest(ByVal strURL As String, strMethod As String, HWndToMsg As Long, Optional strProxy As String, Optional strPostCommand As String) As Long

    On Error GoTo Problems

  Dim intSock As Integer
  Dim strHost As String
  Dim intPort As Integer
  Dim intDot As Integer
  Dim intSep As Integer
  Dim strMsg As String

    'Remove http if not using Proxy
    If StrComp(Left$(strURL, 4), "http", vbTextCompare) = 0 And strProxy = "" Then
        strURL = mid$(strURL, InStr(5, strURL, "//") + 2)
      Else 'Add http if using Proxy
        If strProxy <> "" Then
            If StrComp(Left$(strURL, 4), "http", vbTextCompare) <> 0 Then strURL = "http://" & strURL
        End If
    End If

    If strProxy <> "" Then 'There is a Proxy
        'Find out Host and Port
        strHost = mid$(strProxy, 1, InStr(1, strProxy, ":") - 1)
        intPort = mid$(strProxy, InStr(1, strProxy, ":") + 1)

        'If there's no / in URL add on at the end
        If StrComp(Left$(strURL, 4), "http", vbTextCompare) = 0 Then
            If InStr(9, strURL, "/") = 0 Then strURL = strURL & "/"
        End If
      Else 'No Proxy
        intDot = InStr(1, strURL, ":")
        If intDot > 0 And (InStr(1, strURL, "/") = 0 Or InStr(1, strURL, "/") > intDot) Then 'Port needs to be changed
            intSep = InStr(intDot + 1, strURL, "/")
            If intSep = 0 Then intSep = Len(strURL) + 1
            'get Host and Port
            strHost = mid$(strURL, 1, intDot - 1)
            intPort = mid$(strURL, intDot + 1, intSep - intDot - 1)
          Else 'Port is default: 80
            intSep = InStr(1, strURL, "/")
            If intSep = 0 Then intSep = Len(strURL) + 1
            'get Host
            strHost = mid$(strURL, 1, intSep - 1)
            intPort = 80
        End If

        'If requested file is default index add / if necessary
        strURL = mid$(strURL, intSep)
        If strURL = "" Then strURL = "/"

    End If

    'Connect
    intSock = ConnectSock(strHost, intPort, HWndToMsg)

    If intSock = SOCKET_ERROR Then
        HTTPRequest = SOCKET_ERROR
        Exit Function
    End If

    'Form request string
    strMsg = strMethod & " " & strURL & " HTTP/1.0" & vbCrLf
    strMsg = strMsg & "Accept: */*" & vbCrLf
    strMsg = strMsg & "User-Agent: " & app.Title & vbCrLf
    strMsg = strMsg & "Host: " & strHost & vbCrLf
    If strMethod = "POST" Then 'Complete POST request, URL Encoded!
        strMsg = strMsg & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
        strMsg = strMsg & "Content-Length: " & Len(strPostCommand) & vbCrLf
        strMsg = strMsg & vbCrLf & strPostCommand
      Else 'GET request
        strMsg = strMsg & vbCrLf
    End If

    'Send request
    Call SendData(intSock, strMsg)

    'Return socket number
    HTTPRequest = intSock

Exit Function

Problems:
    HTTPRequest = SOCKET_ERROR
    Call EndWinsock

End Function

Function IpToAddr(ByVal AddrOrIP$) As String

    On Error Resume Next
        IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP$))
        If Err Then IpToAddr = WSA_NoName

End Function

Function IrcGetAscIp(ByVal IPL$) As String

  'this function is IRC specific, it expects a long ip stored in Network byte order, in a string
  'the kind that would be parsed out of a DCC command string

    On Error GoTo IrcGetAscIPError

  Dim lpStr&
  Dim nStr&
  Dim retString$
  Dim inn&
    If val(IPL) > 2147483647 Then
        inn = val(IPL) - 4294967296#
      Else
        inn = val(IPL)
    End If
    inn = ntohl(inn)
    retString = String$(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        IrcGetAscIp = "0.0.0.0"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left$(retString, nStr)
    IrcGetAscIp = retString

Exit Function

IrcGetAscIPError:
    IrcGetAscIp = "0.0.0.0"

End Function

Function IrcGetLongIp(ByVal AscIp$) As String

  'this function converts an ascii ip string into a long ip in network byte order
  'and stick it in a string suitable for use in a DCC command.

    On Error GoTo IrcGetLongIpError

  Dim inn&
    inn = inet_addr(AscIp)
    inn = htonl(inn)
    If inn < 0 Then
        IrcGetLongIp = CVar(inn + 4294967296#)
        Exit Function
      Else
        IrcGetLongIp = CVar(inn)
        Exit Function
    End If

Exit Function

IrcGetLongIpError:
    IrcGetLongIp = "0"

End Function

Public Function IsConnected() As Boolean

    IsConnected = InternetGetConnectedState(0&, 0&)

End Function

'For waiting when sending
Public Function IsSocketReady(intSocketNumber As Integer, intTimeOut As Integer, Optional blnReading As Boolean, Optional blnWriting As Boolean, Optional blnExcept As Boolean) As Boolean

  Dim lngReturn As Long

  Dim tmpReadSet As fd_set
  Dim tmpWriteSet As fd_set
  Dim tmpExceptSet As fd_set
  Dim tmpTimeVal As timeval

    If blnWriting Then
        tmpWriteSet.fd_count = 1
        tmpWriteSet.fd_array(0) = intSocketNumber
      Else
        If blnReading Then
            tmpReadSet.fd_count = 1
            tmpReadSet.fd_array(0) = intSocketNumber
          Else
            tmpExceptSet.fd_count = 1
            tmpExceptSet.fd_array(0) = intSocketNumber
        End If
    End If

    tmpTimeVal.tv_sec = intTimeOut
    tmpTimeVal.tv_usec = intTimeOut

    lngReturn = ws_select(0, tmpReadSet, tmpWriteSet, tmpExceptSet, tmpTimeVal)

    If lngReturn > 0 Or lngReturn = SOCKET_ERROR Then IsSocketReady = True

End Function

Public Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&) As Long

  Dim intSocket&
  Dim SelectOps&
  Dim sockin As sockaddr

    Call StartWinsock

    'Configure our socket
    sockin = saZero     'zero out the structure
    With sockin
        .sin_family = AF_INET
        .sin_port = htons(Port)      ' Port to listen
        .sin_zero = String$(8, 0)
    End With

    If sockin.sin_port = INVALID_SOCKET Then
        ListenForConnect = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Try to create our TCP socket
    intSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    'An error occured
    If intSocket < 0 Then
        ListenForConnect = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Bind socket...
    If bind(intSocket, sockin, sockaddr_size) Then
        'An error occured while trying to bind
        If intSocket > 0 Then
            Call closesocket(intSocket)
        End If
        ListenForConnect = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    'Perform Async functions, with our form's hwnd
    SelectOps = FD_CONNECT Or FD_ACCEPT
    If WSAAsyncSelect(intSocket, HWndToMsg, ByVal &H202, ByVal SelectOps) Then
        If intSocket > 0 Then
            Call closesocket(intSocket)
        End If
        ListenForConnect = SOCKET_ERROR
        Call EndWinsock
        Exit Function
    End If

    'Start listening for connections
    If listen(intSocket, 1) Then
        If intSocket > 0 Then
            Call closesocket(intSocket)
        End If
        ListenForConnect = INVALID_SOCKET
        Call EndWinsock
        Exit Function
    End If

    ListenForConnect = intSocket

End Function

Public Function MyIP() As String

    MyIP = AddrToIP(GetLocalHostName)

End Function

Public Function Ping(ByVal hostnameOrIpaddress As String, Optional timeOutmSec As Long = PING_TIMEOUT, Optional bolReturnTime As Boolean)

  Dim echoValues As ICMP_ECHO_REPLY
  Dim hPort As Long
  Dim dwAddress As Long
  Dim sDataToSend As String

    On Error GoTo e_Trap
    If Trim$(hostnameOrIpaddress) = "" Then GoTo e_Trap

    sDataToSend = "Echo This"
    hostnameOrIpaddress = AddrToIP(hostnameOrIpaddress)
    dwAddress = GetHostByNameAlias(hostnameOrIpaddress)
    If dwAddress = -1 Then GoTo e_Trap

    hPort = IcmpCreateFile()

    'ping an ip address, passing the
    'address and the ECHO structure
    Call IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, echoValues, Len(echoValues), timeOutmSec)
    'If Ping succeeded, .Status will be 0
    '.RoundTripTime is the time in ms for the ping to complete
    '.Data is the data returned (NULL terminated)
    '.Address is the Ip address that actually replied
    '.DataSize is the size of the string in .Data

    Call IcmpCloseHandle(hPort)

    'If Ping is OK And Reply was correct
    If echoValues.Status = 0 And sDataToSend = Left$(echoValues.Data, Len(sDataToSend)) Then
        If bolReturnTime Then 'Return RoundTripTime
            Ping = echoValues.RoundTripTime
          Else 'Return True
            If echoValues.RoundTripTime <= timeOutmSec Then Ping = True
        End If
      Else 'Ping not OK
        If bolReturnTime Then 'Return RoundTripTime
            Ping = SOCKET_ERROR
          Else 'Return False
            Ping = False
        End If
    End If

Exit Function

e_Trap:
    If bolReturnTime Then
        Ping = SOCKET_ERROR
      Else
        Ping = False
    End If

    Call EndWinsock

End Function

Public Function SendData(ByVal intSocket&, vMessage As Variant) As Long

  Dim arrMessage() As Byte, strTemp As String

    Select Case VarType(vMessage)
      Case 8209   'Byte array
        strTemp = vMessage
      Case 8      'String, if we receive a string, it's assumed we are in line mode
        strTemp = StrConv(vMessage, vbFromUnicode)
      Case Else
        strTemp = StrConv(CStr(vMessage), vbFromUnicode)
    End Select

    arrMessage = strTemp

    If UBound(arrMessage) > -1 Then
        SendData = Send(intSocket, arrMessage(0), UBound(arrMessage) + 1, 0)
    End If

    If SendData = SOCKET_ERROR Then
        Call closesocket(intSocket)
        Call EndWinsock
    End If

End Function

Public Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long

  Dim Linger As LingerType

    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        Debug.Print "Error setting linger info: " & WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
      Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            Debug.Print "Error getting linger info: " & WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
          Else
            Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
            Debug.Print "Linger time if linger is on: "; Linger.l_linger
        End If
    End If

End Function

Public Function SockAddressToString(sa As sockaddr) As String

    SockAddressToString = getascip(sa.sin_addr) & ":" & ntohs(sa.sin_port)

End Function

Public Function StartWinsock(Optional sDescription As String) As Boolean

  Dim StartupData As WSADataType

    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            Debug.Print "wVersion="; StartupData.wVersion, "wHighVersion="; StartupData.wHighVersion
            Debug.Print "If wVersion = 257 then everything is OK"
            Debug.Print "szDescription="; StartupData.szDescription
            Debug.Print "szSystemStatus="; StartupData.szSystemStatus
            Debug.Print "iMaxSockets="; StartupData.iMaxSockets, "iMaxUdpDg="; StartupData.iMaxUdpDg
            sDescription = StartupData.szDescription
          Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp

End Function

Function URLEncode(strWhat As String) As String

  Dim i As Long

    strWhat = Replace$(strWhat, "%", "%25")

    For i = 0 To 32
        strWhat = Replace$(strWhat, Chr$(i), "%" & IIf(i < 16, "0" & Hex$(i), Hex$(i)))
    Next i

    strWhat = Replace$(strWhat, "&", "%26")
    strWhat = Replace$(strWhat, "+", "%2B")

    'Return
    URLEncode = strWhat

End Function

Public Function WSAGetAsyncBufLen(ByVal lParam As Long) As Long

    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetAsyncBufLen = (lParam And &HFFFF&) - &H10000
      Else
        WSAGetAsyncBufLen = lParam And &HFFFF&
    End If

End Function

Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer

    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000

End Function

Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer

    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
      Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If

End Function

Public Function WSAMakeSelectReply(TheEvent%, TheError%) As Long

    WSAMakeSelectReply = (TheError * &H10000) + (TheEvent And &HFFFF&)

End Function

':) Ulli's VB Code Formatter V2.13.6 (30.07.2005 11:24:21) 313 + 1025 = 1338 Lines
