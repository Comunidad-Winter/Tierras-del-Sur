Attribute VB_Name = "ObtenerMacaddress"
Option Explicit

Public Const NCBASTAT As Long = &H33
Public Const NCBNAMSZ As Long = 16
Public Const HEAP_ZERO_MEMORY As Long = &H8
Public Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Public Const NCBRESET As Long = &H32

Public Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte ' Reserved, must be 0
   ncb_event      As Long
End Type

Public Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
   
Public Type NAME_BUFFER
   Name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Public Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Public Declare Function Netbios Lib "netapi32.dll" _
   (pncb As NET_CONTROL_BLOCK) As Byte
     
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, ByVal _
    hpvSource As Long, ByVal _
    cbCopy As Long)
     
Public Declare Function GetProcessHeap Lib "kernel32" () As Long

Public Declare Function HeapAlloc Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, _
     ByVal dwBytes As Long) As Long
     
Public Declare Function HeapFree Lib "kernel32" _
    (ByVal hHeap As Long, _
     ByVal dwFlags As Long, _
     lpMem As Any) As Long
     
'Marche
Public Function GetMACAddress() As String

    Dim i As Integer
    
    Call fReadValue("HKLM", "SOFTWARE\Installhere", "CurrentHost", "S", "", GetMACAddress)
    If GetMACAddress = "" Then
        For i = 1 To 5
            Randomize timer
            GetMACAddress = GetMACAddress & Chr(65 + Int(Rnd() * 17) + Int(Rnd() * 10))
        Next
        GetMACAddress = Second(Time) & GetMACAddress & CLng(GetTickCount)
        For i = 1 To 5
            Randomize timer
            GetMACAddress = GetMACAddress & Chr(65 + Int(Rnd() * 17) + Int(Rnd() * 10))
        Next
        Call fWriteValue("HKLM", "SOFTWARE\InstallHere", "CurrentHost", "S", CryptStr(GetMACAddress, 0))
    Else
    GetMACAddress = DecryptStr(GetMACAddress, 0) & "A"
  
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
        UserDomain = left$(UserDomain, 15)
    End If
        
    If Len(UserName) > 15 Then
        UserName = left$(UserName, 15)
    End If
End If


GetIdentificacionPC = UserName & UserDomain

End Function



