Attribute VB_Name = "modRegistro"
Option Explicit
'
' Written by Dave Scarmozzino  www.TheScarms.com
'
Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As Long
    bInheritHandle       As Boolean
End Type

Public Const MAX_SIZE = 2048
Public Const MAX_INISIZE = 8192

' Constants for Registry top-level keys
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CLASSES_ROOT = &H80000000

' Return values
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_MORE_DATA = 234
Public Const ERROR_NO_MORE_ITEMS = 259&

' RegCreateKeyEx options
Public Const REG_OPTION_NON_VOLATILE = 0

' RegCreateKeyEx Disposition
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

' Registry data types
Public Const REG_NONE = 0
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

' Registry security attributes
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8


Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, _
        ByVal dwOptions As Long, ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
        lpdwDisposition As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpszValueName As String, _
        ByVal lpdwReserved As Long, lpdwType As Long, _
        lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        lpData As Any, ByVal cbData As Long) As Long


Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
        

Declare Function GetPrivateProfileInt Lib "kernel32" _
        Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
        ByVal lpKeyname As String, ByVal nDefault As Long, ByVal lpFileName _
        As String) As Long
    
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long




Public Function fReadValue(ByVal sTopKeyOrFile As String, _
    ByVal sSubKeyOrSection As String, ByVal sValueName As String, _
    ByVal sValueType As String, ByVal vDefault As Variant, _
    vValue As Variant) As Long
'
' Written by Dave Scarmozzino  www.TheScarms.com
'
' Use this function to read a:
'   -   String, 16-bit binary (True|False), 32-bit integer registry value or
'   -   String or integer value from an .ini file.
'
' sTopKeyOrIniFile
'   -   A top level registry key abbreviation {"HKCU","HKLM","HKU","HKDD","HKCC","HKCR"} or
'   -   The full path of an .ini file (ex. "C:\Windows\MyFile.ini")
'
' sSubKeyOrSection
'   -   A registry subkey or
'   -   An .ini file section name
'
' sValueName
'   -   A registry entry or
'   -   An .ini file entry
'
' sValueType
'   -   "S" to read a string value or
'   -   "B" to read a 16-bit binary value (applies to registry use only) or
'   -   "D" to read a 32-bit number value (applies to registry use only).
'
' vDefault
'   -   The default value to return. It can be a string or boolean.
'
' vValue
'   -   The value read. It can be a string or boolean.
'   -   vDefault if unsuccessful (0 when reading an integer from an .ini file)
'
' Return Value
'   -   0 if successful, non-zero otherwise.
'
' Example 1   -   Read a string value from the registry.
'   lResult = fReadValue("HKCU", "Software\YourKey\LastKey\YourApp", "AppName", "S", "", sValue)
'
' Example 2   -   Read a boolean (True|False) value from the registry.
'   lResult = fReadValue("HKCU", "Software\YourKey\LastKey\YourApp", "AutoHide", "B", False, bValue)
'
' Example 3   -   Read an integer value from the registry.
'   lResult = fReadValue("C:\Windows\Myfile.ini", "SectionName", "NumApps", "D", 12345, lValue)
'
' Example 4   -   Read a string value from an .ini file.
'   lResult = fReadValue("C:\Windows\Myfile.ini", "SectionName", "AppName", "S", "", sValue)
'
' Example 5   -   Read an integer value from an .ini file.
'   lResult = fReadValue("C:\Windows\Myfile.ini", "SectionName", "NumApps", "B", "0", iValue)
'
Dim lTopKey     As Long
Dim lHandle     As Long
Dim lLenData    As Long
Dim lResult     As Long
Dim lDefault    As Long
Dim lValue      As Long
Dim svalue      As String
Dim sDefaultStr As String
Dim bValue      As Boolean

On Error GoTo fReadValueError
lResult = 99
vValue = vDefault
lTopKey = fTopKey(sTopKeyOrFile)
If lTopKey = 0 Then GoTo fReadValueError

If lTopKey = 1 Then
    '
    ' Read the .ini file value.
    '
    If UCase$(sValueType) = "S" Then
        lLenData = 255
        sDefaultStr = vDefault
        svalue = Space$(lLenData)
        lResult = getprivateprofilestring(sSubKeyOrSection, sValueName, sDefaultStr, svalue, lLenData, sTopKeyOrFile)
        vValue = left$(svalue, lResult)
    Else
        lDefault = 0
        lResult = GetPrivateProfileInt(sSubKeyOrSection, sValueName, lDefault, sTopKeyOrFile)
    End If
Else
    '
    ' Open the registry SubKey.
    '
    lResult = RegOpenKeyEx(lTopKey, sSubKeyOrSection, 0, KEY_QUERY_VALUE, lHandle)
    If lResult <> ERROR_SUCCESS Then
        fReadValue = lResult
        Exit Function
    End If
    '
    ' Get the actual value.
    '
    Select Case UCase$(sValueType)
        Case "S"
            '
            ' String value. The first query gets the string length. The second
            ' gets the string value.
            '
            lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_SZ, "", lLenData)
            If lResult = ERROR_MORE_DATA Then
                svalue = Space(lLenData)
                lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_SZ, ByVal svalue, lLenData)
            End If
            If lResult = ERROR_SUCCESS Then  'Remove null character.
                vValue = left$(svalue, lLenData - 1)
            Else
                GoTo fReadValueError
            End If
        Case "B"
            lLenData = Len(bValue)
            lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_BINARY, bValue, lLenData)
            If lResult = ERROR_SUCCESS Then
                vValue = bValue
            Else
                GoTo fReadValueError
            End If
        Case "D"
            lLenData = 32
            lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_DWORD, lValue, lLenData)
            If lResult = ERROR_SUCCESS Then
                vValue = lValue
            Else
                GoTo fReadValueError
            End If
    End Select
    '
    ' Close the key.
    '
    lResult = RegCloseKey(lHandle)
    fReadValue = lResult
End If
Exit Function
'
' Error processing.
'
fReadValueError:
    fReadValue = lResult
End Function


Private Function fTopKey(ByVal sTopKeyOrFile As String) As Long
Dim sDir   As String
'
' Written by Dave Scarmozzino  www.TheScarms.com
'
' This function returns:
'   -   the numeric value of a top level registry key or
'   -   1 if sTopKey is a valid .ini file or
'   -   0 otherwise.
'
On Error GoTo fTopKeyError
fTopKey = 0
Select Case UCase$(sTopKeyOrFile)
    Case "HKCU"
        fTopKey = HKEY_CURRENT_USER
    Case "HKLM"
        fTopKey = HKEY_LOCAL_MACHINE
    Case "HKU"
        fTopKey = HKEY_USERS
    Case "HKDD"
        fTopKey = HKEY_DYN_DATA
    Case "HKCC"
        fTopKey = HKEY_CURRENT_CONFIG
    Case "HKCR"
        fTopKey = HKEY_CLASSES_ROOT
    Case Else
        On Error Resume Next
        sDir = Dir$(sTopKeyOrFile)
        If Err.Number = 0 And sDir <> "" Then fTopKey = 1
End Select
Exit Function

fTopKeyError:
End Function

Public Function fWriteValue(ByVal sTopKeyOrFile As String, _
    ByVal sSubKeyOrSection As String, ByVal sValueName As String, _
    ByVal sValueType As String, ByVal vValue As Variant) As Long
'
' Written by Dave Scarmozzino  www.TheScarms.com
'
' Use this function to write a:
'   -   String, 16-bit binary (True|False), 32-bit integer registry value or
'   -   String value to an .ini file.
'
' sTopKeyOrIniFile
'   -   A top level registry key abbreviation {"HKCU","HKLM","HKU","HKDD","HKCC","HKCR"} or
'   -   The full path of an .ini file (ex. "C:\Windows\MyFile.ini")
'
' sSubKeyOrSection
'   -   A registry subkey or
'   -   An .ini file section name
'
' sValueName
'   -   A registry entry or
'   -   An .ini file entry
'
' sValueType
'   -   "S" to write a string value or
'   -   "B" to write a 16-bit binary value (applies to registry use only) or
'   -   "D" to write a 32-bit number value (applies to registry use only).
'
' vValue
'   -   The value to write. It can be a string, binary or integer.
'
' Return Value
'   -   0 if successful, non-zero otherwise.
'
' Example 1   -   Write a string value to the registry.
'   lResult = fWriteValue("HKCU", "Software\YourKey\LastKey\YourApp", "AppName", "S", "MyApp")
'
' Example 2   -   Write a True|False value to the registry.
'   lResult = fWriteValue("HKCU", "Software\YourKey\LastKey\YourApp", "AutoHide", "B", True)
'
' Example 3   -   Write an integer value to the registry.
'   lResult = fWriteValue("HKCU", "Software\YourKey\LastKey\YourApp", "NumOfxxx", "D", 12345)
'
' Example 4   -   Write a string value to an .ini file.
'   lResult = fWriteValue("C:\Windows\Myfile.ini", "SectionName", "AppName", "S", "MyApp")
'
' NOTE:
'   This function cannot write a non-string value to an .ini file.
'
Dim lTopKey             As Long
Dim lOptions            As Long
Dim lsamDesired         As Long
Dim lHandle             As Long
Dim lDisposition        As Long
Dim lLenData            As Long
Dim lResult             As Long
Dim lValue              As Long
Dim sClass              As String
Dim svalue              As String
Dim bValue              As Boolean
Dim tSecurityAttributes As SECURITY_ATTRIBUTES

On Error GoTo fWriteValueError
lResult = 99
lTopKey = fTopKey(sTopKeyOrFile)
If lTopKey = 0 Then GoTo fWriteValueError

If lTopKey = 1 Then
    '
    ' Read the .ini file value.
    '
    If UCase$(sValueType) = "S" Then
        svalue = vValue
        lResult = writeprivateprofilestring(sSubKeyOrSection, sValueName, svalue, sTopKeyOrFile)
    Else
        GoTo fWriteValueError
    End If
Else
    sClass = ""
    lOptions = REG_OPTION_NON_VOLATILE
    lsamDesired = KEY_CREATE_SUB_KEY Or KEY_SET_VALUE
    '
    ' Create the SubKey or open it if it exists. Return its handle.
    ' lDisposition will be REG_CREATED_NEW_KEY if the key did not exist.
    '
    lResult = RegCreateKeyEx(lTopKey, sSubKeyOrSection, 0, sClass, lOptions, _
                  lsamDesired, tSecurityAttributes, lHandle, lDisposition)
    If lResult <> ERROR_SUCCESS Then GoTo fWriteValueError
    '
    ' Set the actual value.
    '
    Select Case UCase$(sValueType)
        Case "S"
            svalue = vValue
'02/05/2002            lLenData = Len(sValue) + 1
            lLenData = Len(svalue)
            lResult = RegSetValueEx(lHandle, sValueName, 0, REG_SZ, ByVal svalue, lLenData)
        Case "B"
            bValue = vValue
            lLenData = Len(bValue)
            lResult = RegSetValueEx(lHandle, sValueName, 0, REG_BINARY, bValue, lLenData)
        Case "D"
            lValue = CInt(vValue)
            lLenData = 4
            lResult = RegSetValueEx(lHandle, sValueName, 0, REG_DWORD, lValue, lLenData)
    End Select
    '
    ' Close the key.
    '
    If lResult = ERROR_SUCCESS Then
        lResult = RegCloseKey(lHandle)
        fWriteValue = lResult
        Exit Function
    End If
End If
Exit Function
'
' Error processing.
'
fWriteValueError:
    fWriteValue = lResult
End Function

