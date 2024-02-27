Attribute VB_Name = "modMAC"
Option Explicit

' Declarations needed for GetAdaptersInfo & GetIfTable
Private Const MIB_IF_TYPE_OTHER                   As Long = 1
Private Const MIB_IF_TYPE_ETHERNET                As Long = 6
Private Const MIB_IF_TYPE_TOKENRING               As Long = 9
Private Const MIB_IF_TYPE_FDDI                    As Long = 15
Private Const MIB_IF_TYPE_PPP                     As Long = 23
Private Const MIB_IF_TYPE_LOOPBACK                As Long = 24
Private Const MIB_IF_TYPE_SLIP                    As Long = 28

Private Const MIB_IF_ADMIN_STATUS_UP              As Long = 1
Private Const MIB_IF_ADMIN_STATUS_DOWN            As Long = 2
Private Const MIB_IF_ADMIN_STATUS_TESTING         As Long = 3

Private Const MIB_IF_OPER_STATUS_NON_OPERATIONAL  As Long = 0
Private Const MIB_IF_OPER_STATUS_UNREACHABLE      As Long = 1
Private Const MIB_IF_OPER_STATUS_DISCONNECTED     As Long = 2
Private Const MIB_IF_OPER_STATUS_CONNECTING       As Long = 3
Private Const MIB_IF_OPER_STATUS_CONNECTED        As Long = 4
Private Const MIB_IF_OPER_STATUS_OPERATIONAL      As Long = 5

Private Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Private Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Private Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Private Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8
Private Const DEFAULT_MINIMUM_ENTITIES            As Long = 32
Private Const MAX_HOSTNAME_LEN                    As Long = 128
Private Const MAX_DOMAIN_NAME_LEN                 As Long = 128
Private Const MAX_SCOPE_ID_LEN                    As Long = 256

Private Const MAXLEN_IFDESCR                      As Long = 256
Private Const MAX_INTERFACE_NAME_LEN              As Long = MAXLEN_IFDESCR * 2
Private Const MAXLEN_PHYSADDR                     As Long = 8

' Information structure returned by GetIfEntry/GetIfTable
Private Type MIB_IFROW
    wszName(0 To MAX_INTERFACE_NAME_LEN - 1) As Byte    ' MSDN Docs say pointer, but it is WCHAR array
    dwIndex             As Long
    dwType              As Long
    dwMtu               As Long
    dwSpeed             As Long
    dwPhysAddrLen       As Long
    bPhysAddr(MAXLEN_PHYSADDR - 1) As Byte
    dwAdminStatus       As Long
    dwOperStatus        As Long
    dwLastChange        As Long
    dwInOctets          As Long
    dwInUcastPkts       As Long
    dwInNUcastPkts      As Long
    dwInDiscards        As Long
    dwInErrors          As Long
    dwInUnknownProtos   As Long
    dwOutOctets         As Long
    dwOutUcastPkts      As Long
    dwOutNUcastPkts     As Long
    dwOutDiscards       As Long
    dwOutErrors         As Long
    dwOutQLen           As Long
    dwDescrLen          As Long
    bDescr As String * MAXLEN_IFDESCR
End Type

Private Type TIME_t
    aTime As Long
End Type

Private Type IP_ADDRESS_STRING
    IPadrString     As String * 16
End Type

Private Type IP_ADDR_STRING
    AdrNext         As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    NTEcontext      As Long
End Type

' Information structure returned by GetIfEntry/GetIfTable
Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
    Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
    MACadrLength        As Long
    MACaddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    AdapterIndex        As Long
    AdapterType         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    DhcpEnabled         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    CurrentIpAddress    As Long
    IpAddressList       As IP_ADDR_STRING
    GatewayList         As IP_ADDR_STRING
    DhcpServer          As IP_ADDR_STRING
    HaveWins            As Long             ' MSDN Docs say "Bool", but is 4 bytes
    PrimaryWinsServer   As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained       As TIME_t
    LeaseExpires        As TIME_t
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
Private Declare Function GetNumberOfInterfaces Lib "iphlpapi.dll" (ByRef pdwNumIf As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public hamachi As Boolean
Public ClientID As Double
Public ClientIDs As String * 16
Public macaddr As String
Public hIP As String

Private Const BIGNUMBER_32 As Double = 4294967296#

Private Declare Function GetTickCount Lib "kernel32" () As Long

Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
End Function


Private Function lngSigned(ByVal dblUnsigned As Double) As Long
    If dblUnsigned <= &H7FFFFFFF Then 'If uDouble is less than or equal To 0x7FFFFFFF, just return the raw uDouble value.
        lngSigned = dblUnsigned 'Return the raw uDouble value
    Else 'If uDouble is equal To or greater than &H80000000, we must process the uDouble value.
        lngSigned = CLng(dblUnsigned - BIGNUMBER_32) 'Because Long is a 32-bit, signed value, we must subtract (2^32) from uDouble.
    End If 'End of If statement
End Function

Private Function dblUnsigned(ByVal lngSigned As Long) As Double
    If lngSigned >= 0 Then 'If sLong is equal To or greater than zero, just return the raw sLong value.
        dblUnsigned = lngSigned 'Return the raw sLong value
    Else 'If xFileLen is less than zero (eg. -1, -2, -3, -4, -..., -65536, etc), we must process the sLong value.
        dblUnsigned = BIGNUMBER_32 + lngSigned 'Because Long is a 32-bit signed value, we need To add sLong to (2 to the power of 32). Because sLong is negative, adding the two numbers is basically subtracting sLong from (2^32).
    End If 'End of If statement
End Function

Function Hex2Decimal(sHexVal As String) As Long
    Hex2Decimal = val("&H" & sHexVal & "&")
End Function

Private Function MAC2String(AdrArray() As Byte) As String
On Error GoTo enda:
    Dim aStr As String, hexStr As String, i%
    For i = 0 To 5
        If (i > UBound(AdrArray)) Then
            hexStr = "00"
        Else
            hexStr = Hex$(AdrArray(i))
        End If
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr
    Next i
    MAC2String = aStr
    If aStr = "000000000000" Then GoTo enda
Exit Function
enda:
    Dim bytearr(3) As Byte
    Dim tmp&
    tmp = GetSerialNumber("C:\")
    
    Call CopyMemory(bytearr(0), tmp, 4)
    'Marce On error resume next
    For i = 0 To 5
        If (i > 3) Then
            hexStr = "FF"
        Else
            hexStr = Hex$(bytearr(i))
        End If
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr
    Next i
    MAC2String = aStr
End Function

Public Function get_pc_id() As Double

Dim t As Long, p As Long

t = GetDWord(HKEY_LOCAL_MACHINE, "SOFTWARE\Firewall\CurrentHost", "T")
If t = 0 Then
    t = Rnd * &HFFFFFFFF Xor GetTickCount Xor Rnd * &HFF0000FF Or Rnd * &HFFFF00
    Call SaveDWord(HKEY_LOCAL_MACHINE, "SOFTWARE\Firewall\CurrentHost", "T", t)
End If

get_pc_id = dblUnsigned(t)

End Function

Public Function get_mac_address() As String
    Dim bytes() As Byte
    bytes() = mac_get()
    get_mac_address = MAC2String(bytes)
End Function

Public Function get_hamachi_active() As Boolean
    Call mac_get(get_hamachi_active, hIP)
End Function

Private Function mac_get(Optional ByRef hamachi As Boolean, Optional ByRef haip As String) As Byte()
    Dim tmpm(0 To 7) As Byte
    
    mac_get = tmpm()
    Dim AdapInfo As IP_ADAPTER_INFO, bufLen As Long, sts As Long
    Dim retStr As String, numStructs%, i%, IPinfoBuf() As Byte, srcPtr As Long
    sts = GetAdaptersInfo(AdapInfo, bufLen)
    If (bufLen = 0) Then Exit Function
    numStructs = bufLen / Len(AdapInfo)
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetAdaptersInfo(IPinfoBuf(0), bufLen)
    
    
    If (sts <> 0) Then Exit Function
    
    
    srcPtr = VarPtr(IPinfoBuf(0))
    For i = 0 To numStructs - 1
        If (srcPtr = 0) Then Exit For
        Call CopyMemory(AdapInfo, ByVal srcPtr, Len(AdapInfo))
        With AdapInfo
            If (.AdapterType = MIB_IF_TYPE_ETHERNET) Then
                If Not (.Description Like "*Hamachi*") Then
                    mac_get = .MACaddress()
                Else
                    hamachi = True
                    haip = StripTerminator(.IpAddressList.IpAddress.IPadrString)
                End If
            End If
        End With
        srcPtr = AdapInfo.Next
    Next i
End Function

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))


    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function ci2hex(ByVal id As Double) As String
Dim j(1) As Long
    CopyMemory j(0), id, 8
    ci2hex = Hex$(j(1))
    Do While Len(ci2hex) < 8
        ci2hex = "0" & ci2hex
    Loop
    ci2hex = Hex$(j(0)) & ci2hex
    Do While Len(ci2hex) < 16
        ci2hex = "0" & ci2hex
    Loop
End Function

Public Function GetMACAddress_Pom() As String
Dim i As Integer

GetMACAddress_Pom = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Firewall\CurrentHost", "de")
If GetMACAddress_Pom = "" Then
    For i = 1 To 6
        Randomize Timer
        GetMACAddress_Pom = GetMACAddress_Pom & Hex((Rnd * 97) Mod 16) & Int((Rnd() * 3717) Mod 9)
    Next
    Call SaveString(HKEY_LOCAL_MACHINE, "SOFTWARE\Firewall\CurrentHost", "de", GetMACAddress_Pom)
End If

End Function

Private Function quitarCaracteresRaros(cadena As String) As String
       
Dim i As Byte
Dim caracter As Integer
       
On Error GoTo quitarCaracteresRaros_Err

For i = 1 To Len(cadena)
    caracter = Asc(mid$(cadena, i, 1))
    If (caracter >= 65 And caracter <= 90) Or (caracter >= 97 And caracter <= 122) Or (caracter >= 48 And caracter <= 57) Then
        quitarCaracteresRaros = quitarCaracteresRaros & Chr(caracter)
    End If
Next

'<EhFooter>
Exit Function

quitarCaracteresRaros_Err:
        quitarCaracteresRaros = "ERROR"
End Function
Public Function GetIdentificacionPC() As String

Dim UserName As String
Dim UserDomain As String
UserName = Environ("USERNAME")
UserDomain = Environ("USERDOMAIN")

'Saco el WIN si lo tiene
If InStr(1, UserDomain, "WIN-") = 1 Then
    UserDomain = mid$(UserDomain, 5)
End If
    
UserName = quitarCaracteresRaros(UserName)
UserDomain = quitarCaracteresRaros(UserDomain)
    
If Len(UserDomain) + Len(UserName) > 30 Then
    If Len(UserDomain) > 15 Then
        UserDomain = Left$(UserDomain, 15)
    End If
        
    If Len(UserName) > 15 Then
        UserName = Left$(UserName, 15)
    End If
End If


GetIdentificacionPC = UserName & UserDomain

End Function
