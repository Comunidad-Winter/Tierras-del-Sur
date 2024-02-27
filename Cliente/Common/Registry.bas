Attribute VB_Name = "CDMRegistry"
'Author: Luis Cantero
'© 2002 L.C. Enterprises
'http://LCen.com


Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Const ERROR_NONE = 0

Public Const ERROR_MORE_DATA = 234

Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum RegistryLTypes
    REG_SZ = 1
    REG_BINARY = 3
    REG_DWORD = 4
End Enum

Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegQueryValueExString Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function LoadLibraryRegister Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
  
Private Declare Function CreateThreadForRegister Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
   
Private Declare Function GetProcAddressRegister Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibraryRegister Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Function CreateNewKey(MainKey As Long, SubKey As String)

  Dim hNewKey As Long
  Dim lRetVal As Long

    On Error GoTo Problems
    RegCreateKeyEx MainKey, SubKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal
    RegCloseKey (hNewKey)

Exit Function

Problems:
    MsgBox Err.Description & " (CreateNewKey)", vbExclamation, "Error number " & Err.Number

End Function

Function DeleteKey(MainKey As Long, SubKey As String)

    On Error GoTo Problems
    RegDeleteKey MainKey, SubKey

Exit Function

Problems:
    MsgBox Err.Description & " (DeleteKey)", vbExclamation, "Error number " & Err.Number

End Function

Function DeleteValue(MainKey As Long, SubKey As String, ValueName As String)

  Dim hKey As Long

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey
    RegDeleteValue hKey, ValueName
    RegCloseKey (hKey)

Exit Function

Problems:
    MsgBox Err.Description & " (DeleteValue)", vbExclamation, "Error number " & Err.Number

End Function

Function KeyCount(MainKey As Long, SubKey As String)

  Dim ft As FILETIME
  Dim hKey As Long
  Dim Res As Long
  Dim Count As Long
  Dim keyname As String, classname As String
  Dim keylen As Long, classlen As Long

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey
    Do
        keylen = 2000
        classlen = 2000
        keyname = Space$(keylen)
        classname = Space$(classlen)
        Res = RegEnumKeyEx(hKey, Count, keyname, keylen, 0, classname, classlen, ft)
        Count = Count + 1
    Loop While Res = 0
    KeyCount = Count - 1
    RegCloseKey (hKey)

Exit Function

Problems:
    MsgBox Err.Description & " (KeyCount)", vbExclamation, "Error number " & Err.Number

End Function

Function KeyExists(MainKey As Long, SubKey As String) As Boolean

  Dim hKey As Long

    On Error GoTo Problems
    If RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey) = 0 Then RegCloseKey hKey: KeyExists = True Else KeyExists = False

Exit Function

Problems:
    MsgBox Err.Description & " (KeyExists)", vbExclamation, "Error number " & Err.Number

End Function

Function QueryValue(MainKey As Long, SubKey As String, ValueName As String, lType As RegistryLTypes)

  Dim hKey As Long
  Dim vValue

    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey
    QueryValueEx hKey, ValueName, vValue, lType
    RegCloseKey (hKey)
    QueryValue = vValue

End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant, lType As RegistryLTypes) As Variant

  Dim cch As Long
  Dim lrc As Long
  Dim lValue As Long
  Dim sValue As String

    ReDim bData(0) As Byte

    On Error GoTo Problems

    Select Case lType
      Case REG_SZ

        lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
        sValue = String$(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch)
          Else
            vValue = Empty
        End If
      Case REG_DWORD

        lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If lrc = ERROR_NONE Then vValue = lValue
      Case REG_BINARY

        lrc = RegQueryValueEx(lhKey, szValueName, 0&, lType, bData(0), cch)
        If lrc = ERROR_NONE Or lrc = ERROR_MORE_DATA Then

            ReDim bData(0 To cch - 1)
            lrc = RegQueryValueEx(lhKey, szValueName, CLng(0), lType, bData(0), cch)
        End If
        vValue = bData
      Case Else
        lrc = -1
    End Select

QueryValueExExit:
    If Right$(vValue, 1) = Chr$(0) Then
        vValue = Left$(vValue, Len(vValue) - 1)
    End If

    QueryValueEx = vValue

Exit Function

Problems:
    MsgBox Err.Description & " (QueryValue)", vbExclamation, "Error number " & Err.Number
    Resume QueryValueExExit

End Function

Function SetKeyValue(MainKey As Long, SubKey As String, ValueName As String, ValueSetting As Variant, lType As RegistryLTypes)

  Dim lValue As Long
  Dim sValue As String
  Dim hKey As Long

    ReDim bData(0) As Byte

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey

    Select Case lType
      Case REG_SZ
        sValue = ValueSetting & Chr$(0)
        RegSetValueExString hKey, ValueName, 0&, lType, sValue, Len(sValue)
      Case REG_DWORD
        lValue = ValueSetting
        RegSetValueExLong hKey, ValueName, 0&, lType, lValue, 4
      Case REG_BINARY ' Free form binary

        lLength = (UBound(ValueSetting) - LBound(ValueSetting)) + 1
        ReDim bData(LBound(ValueSetting) To UBound(ValueSetting))
        For i = LBound(ValueSetting) To UBound(ValueSetting)
            bData(i) = CByte(ValueSetting(i))
        Next i
        RegSetValueEx hKey, ValueName, 0&, lType, bData(LBound(ValueSetting)), lLength

    End Select
    RegCloseKey (hKey)

Exit Function

Problems:
    MsgBox Err.Description & " (SetKeyValue)", vbExclamation, "Error number " & Err.Number

End Function

Function ValueCount(MainKey As Long, SubKey As String)

  Dim hKey As Long
  Dim Res As Long
  Dim Count As Long
  Dim lType As Long
  Dim ValueName As String, Valuelen As Long
  Dim lData As String, Datalen As Long

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey
    Do
        ValueName = Space$(255)
        Valuelen = Len(ValueName)
        lData = Space$(255)
        Datalen = Len(lData)
        Res = RegEnumValue(hKey, Count, ValueName, Valuelen, 0, lType, lData, Datalen)
        Count = Count + 1
    Loop While Res = 0
    ValueCount = Count
    RegCloseKey (hKey)

Exit Function

Problems:
    MsgBox Err.Description & " (ValueCount)", vbExclamation, "Error number " & Err.Number

End Function

Public Function RegServer(ByVal FileName As String) As Boolean
'USAGE: PASS FULL PATH OF ACTIVE .DLL OR
'OCX YOU WANT TO REGISTER
RegServer = RegSvr32(FileName, False)
End Function

Public Function UnRegServer(ByVal FileName As String) As Boolean
'USAGE: PASS FULL PATH OF ACTIVE .DLL OR
'OCX YOU WANT TO UNREGISTER
UnRegServer = RegSvr32(FileName, True)
End Function
    
Private Function RegSvr32(ByVal FileName As String, bUnReg As Boolean) As Boolean

Dim lLib As Long
Dim lProcAddress As Long
Dim lThreadID As Long
Dim lSuccess As Long
Dim lExitCode As Long
Dim lThread As Long
Dim bAns As Boolean
Dim sPurpose As String

sPurpose = IIf(bUnReg, "DllUnregisterServer", _
  "DllRegisterServer")

If Dir(FileName) = "" Then Exit Function

lLib = LoadLibraryRegister(FileName)
'could load file
If lLib = 0 Then Exit Function

lProcAddress = GetProcAddressRegister(lLib, sPurpose)

If lProcAddress = 0 Then
  'Not an ActiveX Component
   FreeLibraryRegister lLib
   Exit Function
Else
   lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
   If lThread Then
        lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
        If Not lSuccess Then
           Call GetExitCodeThread(lThread, lExitCode)
           Call ExitThread(lExitCode)
           bAns = False
           Exit Function
        Else
           bAns = True
        End If
        CloseHandle lThread
        FreeLibraryRegister lLib
   End If
End If
    RegSvr32 = bAns
End Function




Public Sub SaveKey(hKey As Long, strPath As String)
    Dim keyhand&
    Dim r&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r&
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function

Function SaveDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function

Public Sub Delstring(hKey As Long, strPath As String, sKey As String)
    Dim keyhand&
    Dim r&
    r = RegOpenKey(hKey, strPath, keyhand&)
    r = RegDeleteValue(keyhand&, sKey)
    r = RegCloseKey(keyhand&)
End Sub

Public Sub SaveSet(AppName As String, Section As String, Key As Variant, Value As Variant)
    SaveString HKEY_CURRENT_USER, "Software\" & app.CompanyName & "\" & AppName & "\" & Section, CStr(Key), CStr(Value)
End Sub

Public Function GetSet(AppName As String, Section As String, Key As Variant, Optional default As Variant) As Variant
    GetSet = GetString(HKEY_CURRENT_USER, "Software\" & app.CompanyName & "\" & AppName & "\" & Section, CStr(Key))
    If GetSet = "" Then GetSet = default
End Function

Public Function DelSet(AppName As String, Section As String, Key As Variant) As Variant
    Delstring HKEY_CURRENT_USER, "Software\" & app.CompanyName & "\" & AppName & "\" & Section, CStr(Key)
End Function

Public Function CPUdata$(Optional ByVal what As String = "Identifier")
    CPUdata = GetString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", what)
    Debug.Print "CPU"; what; CPUdata
End Function


Public Function CPUmhz() As Long
    CPUmhz = GetDWord(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "~MHz")
    Debug.Print "CPU-MHz"; CPUmhz
End Function



':) Ulli's VB Code Formatter V2.13.6 (30.07.2005 11:24:32) 56 + 222 = 278 Lines
