Attribute VB_Name = "CDMFileAccess"
Option Explicit

Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const OFS_MAXPATHNAME = 128
Const OF_READ = &H0

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Public Function GetFilenameFromPath(ByVal sFilePath As String, Optional ByVal bWithExtension As Boolean = True, Optional ByVal enmCase As VbStrConv = vbLowerCase) As String

'---------------------------------------------------------------------------------------
' Author     : Ruturaaj
' Email      : ruturajvpatki@hotmail.com
' Website    : http://www.rcreations.co.nr
'=======================================================================================
' Procedure  : GetFilenameFromPath
' Type       : Function
' ReturnType : String
'=======================================================================================
' Purpose    : Extract File name from given File path.
'---------------------------------------------------------------------------------------

    On Error GoTo GetFilenameFromPath_Error

    Dim sRet As String

    sRet = mid$(sFilePath, InStrRev(sFilePath, "\") + 1)

    If bWithExtension Then
        GetFilenameFromPath = StrConv(sRet, enmCase)
    Else

        If InStr(sRet, ".") = 0 Then
            GetFilenameFromPath = StrConv(sRet, enmCase)
        Else
            GetFilenameFromPath = StrConv(mid$(sRet, 1, InStrRev(sRet, ".") - 1), enmCase)
        End If

    End If

    'This will avoid empty error window to appear.
    Exit Function

GetFilenameFromPath_Error:

    'Show the Error Message with Error Number and its Description.
    MsgBox "Error on Line " & Erl & vbCrLf & vbCrLf & Err.Description, vbCritical, "GetFilenameFromPath Function"

    'Safe Exit from GetFilenameFromPath Function
    Exit Function

End Function



Public Function GetFile(strFilePath As String, frmProgress As Frame, lblPcent As Label, blnCancelFlag As Boolean, Optional blnAsString = True)

  Dim arrFileMain() As Byte
  Dim arrFileBuffer() As Byte
  Dim lngAllBytes As Long
  Dim lngSize As Long, lngRet As Long
  Dim lngFileHandle As Long
  Dim ofData As OFSTRUCT
  Const lngMaxSizeForOneStep = 10000000

    'Prepare Arrays ==========================================================
    ReDim arrFileMain(0)
    ReDim arrFileBuffer(lngMaxSizeForOneStep)

    'Open the two files
    lngFileHandle = OpenFile(GetShortPath(strFilePath), ofData, OF_READ)

    'Get the file size
    lngSize = GetFileSize(lngFileHandle, 0)

    Do While Not UBound(arrFileMain) = lngSize - 1
        If lngSize = 0 Then Exit Function

        'Redim Array to fit a smaller file
        lngAllBytes = UBound(arrFileMain)
        If lngSize - lngAllBytes < lngMaxSizeForOneStep Then ReDim arrFileBuffer(lngSize - lngAllBytes - 2)

        'Read from the file
        ReadFile lngFileHandle, arrFileBuffer(0), UBound(arrFileBuffer) + 1, lngRet, ByVal 0&

        'Calculate Buffer's position in Main Array
        If lngAllBytes > 0 Then lngAllBytes = lngAllBytes + 1

        'Make place for the Buffer in the Main Array
        ReDim Preserve arrFileMain(lngAllBytes + UBound(arrFileBuffer))

        'Put Buffer at end of Main Array
        MemCopy arrFileMain(lngAllBytes), arrFileBuffer(0), UBound(arrFileBuffer) + 1

        frmProgress.Width = (2400 / 100) * (UBound(arrFileMain) * 50 / lngSize)
        lblPcent = Int(UBound(arrFileMain) * 50 / lngSize) & "%"
        DoEvents

        If blnCancelFlag Then
            Call CloseHandle(lngFileHandle)
            Exit Function
        End If
    Loop

    'Close the file
    Call CloseHandle(lngFileHandle)
    ReDim arrFileBuffer(0)

    'Convert Main Array to String
    If blnAsString Then 'Return as string
        GetFile = StrConv(arrFileMain(), vbUnicode)
      Else 'Return as Byte Array
        GetFile = arrFileMain()
    End If

End Function

Public Function GetFileOLD(strFilePath As String, frmProgress As Frame, lblPcent As Label, blnCancelFlag As Boolean, Optional blnAsString = True)

  Dim intFile As Integer
  Dim arrFileMain() As Byte
  Dim arrFileBuffer() As Byte
  Dim lngAllBytes As Long
  Const lngMaxSizeForOneStep = 10000000

    'Prepare Arrays ==========================================================
    ReDim arrFileMain(0)
    ReDim arrFileBuffer(lngMaxSizeForOneStep)

    intFile = FreeFile
    Open strFilePath For Binary Access Read As intFile

    Do While Not EOF(intFile)
        If LOF(intFile) = 0 Then Exit Function

        'Redim Array to fit a smaller file
        lngAllBytes = UBound(arrFileMain)
        If LOF(intFile) - lngAllBytes < lngMaxSizeForOneStep Then ReDim arrFileBuffer(LOF(intFile) - lngAllBytes - 1)

        Get intFile, , arrFileBuffer

        'Calculate Buffer's position in Main Array
        If lngAllBytes > 0 Then lngAllBytes = lngAllBytes + 1

        'Make place for the Buffer in the Main Array
        ReDim Preserve arrFileMain(lngAllBytes + UBound(arrFileBuffer))

        'Put Buffer at end of Main Array
        MemCopy arrFileMain(lngAllBytes), arrFileBuffer(0), UBound(arrFileBuffer) + 1

        frmProgress.Width = (2400 / 100) * (UBound(arrFileMain) * 50 / LOF(intFile))
        lblPcent = Int(UBound(arrFileMain) * 50 / LOF(intFile)) & "%"
        DoEvents

        If blnCancelFlag Then Close intFile: Exit Function
    Loop

    'Close file, set Buffer to 0, delete last char in Main Array
    Close intFile
    ReDim arrFileBuffer(0)
    'Delete last (empty) item
    ReDim Preserve arrFileMain(UBound(arrFileMain) - 1)

    If blnAsString Then 'Return as string
        GetFileOLD = StrConv(arrFileMain(), vbUnicode)
      Else 'Return as Byte Array
        GetFileOLD = arrFileMain()
    End If

End Function

Public Function GetFileQuick(strFilePath As String, Optional blnAsString = True)

  Dim arrFileMain() As Byte
  Dim lngSize As Long, lngRet As Long
  Dim lngFileHandle As Long
  Dim ofData As OFSTRUCT

    'Open the two files
    lngFileHandle = OpenFile(GetShortPath(strFilePath), ofData, OF_READ)

    'Get the file size
    lngSize = GetFileSize(lngFileHandle, 0)

    'Create an array of bytes
    ReDim arrFileMain(lngSize) As Byte

    'Read from the file
    ReadFile lngFileHandle, arrFileMain(0), UBound(arrFileMain), lngRet, ByVal 0&

    'Close the file
    Call CloseHandle(lngFileHandle)

    'Delete last (empty) item
    ReDim Preserve arrFileMain(UBound(arrFileMain) - 1)

    If blnAsString Then 'Return as string
        GetFileQuick = StrConv(arrFileMain(), vbUnicode)
      Else 'Return as Byte Array
        GetFileQuick = arrFileMain()
    End If

End Function

'Rounds a Byte amount and returns KB with 2 decimal places
Public Function GetRoundedKB(lngByteAmount As Long) As Double

    GetRoundedKB = Int(lngByteAmount / 1024 * 100 + 0.5) / 100

End Function

'Rounds a Byte amount and returns, acording to an elapsed time in seconds, KB/s with 2 decimal places
Public Function GetRoundedKBperS(lngByteAmount As Long, lngSecondsElapsed As Double) As Double

    GetRoundedKBperS = Int(lngByteAmount / 1024 / lngSecondsElapsed * 100 + 0.5) / 100

End Function

'Rounds a Byte amount and returns MB with 2 decimal places
Public Function GetRoundedMB(lngByteAmount As Long) As Double

    GetRoundedMB = Int(lngByteAmount / 1048576 * 100 + 0.5) / 100

End Function

Public Function GetShortPath(strCurrentPath As String) As String

  Dim lngLength As Long, strPathBuffer As String

    'Create a buffer
    strPathBuffer = String$(255, 0)

    'retrieve the short pathname
    lngLength = GetShortPathName(strCurrentPath, strPathBuffer, 255)

    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPathBuffer, lngLength)

End Function

':) Ulli's VB Code Formatter V2.13.6 (26.12.2003 14:39:58) 31 + 187 = 218 Lines
