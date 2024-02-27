Attribute VB_Name = "modConsolaDeWindows"
Option Explicit
''''''''''''''''''''''''''''''''''''''''
'Coded by:
' Joacim Andersson: Brixoft Software: http://www.brixoft.net
' William Moeur:  http://moeur.net
''''''''''''''''''''''''''''''''''''''''

' STARTUPINFO flags
Public Const STARTF_USESHOWWINDOW = &H1
Public Const STARTF_USESTDHANDLES = &H100

' ShowWindow flags
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_SHOWNORMAL = 1

' DuplicateHandle flags
Public Const DUPLICATE_CLOSE_SOURCE = &H1
Public Const DUPLICATE_SAME_ACCESS = &H2

' Error codes
Public Const ERROR_BROKEN_PIPE = 109

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
End Type

Public Declare Function CreatePipe _
 Lib "kernel32" ( _
 phReadPipe As Long, _
 phWritePipe As Long, _
 lpPipeAttributes As Any, _
 ByVal nSize As Long) As Long

Public Declare Function ReadFile _
 Lib "kernel32" ( _
 ByVal hFile As Long, _
 lpBuffer As Any, _
 ByVal nNumberOfBytesToRead As Long, _
 lpNumberOfBytesRead As Long, _
 lpOverlapped As Any) As Long

Public Declare Function CreateProcess _
 Lib "kernel32" Alias "CreateProcessA" ( _
 ByVal lpApplicationName As String, _
 ByVal lpCommandLine As String, _
 lpProcessAttributes As Any, _
 lpThreadAttributes As Any, _
 ByVal bInheritHandles As Long, _
 ByVal dwCreationFlags As Long, _
 lpEnvironment As Any, _
 ByVal lpCurrentDriectory As String, _
 lpStartupInfo As STARTUPINFO, _
 lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function GetCurrentProcess _
 Lib "kernel32" () As Long

Public Declare Function DuplicateHandle _
 Lib "kernel32" ( _
 ByVal hSourceProcessHandle As Long, _
 ByVal hSourceHandle As Long, _
 ByVal hTargetProcessHandle As Long, _
 lpTargetHandle As Long, _
 ByVal dwDesiredAccess As Long, _
 ByVal bInheritHandle As Long, _
 ByVal dwOptions As Long) As Long

Public Declare Function CloseHandle _
 Lib "kernel32" ( _
 ByVal hObject As Long) As Long

Public Declare Function OemToCharBuff _
 Lib "user32" Alias "OemToCharBuffA" ( _
 lpszSrc As Any, _
 ByVal lpszDst As String, _
 ByVal cchDstLength As Long) As Long

Public Declare Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As Long, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long _
) As Long

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long _
) As Long

Public Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long _
) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDest As Any, _
    pSrc As Any, _
    ByVal ByteLen As Long _
)

Public Sub StartTimer(hwnd As Long, TimerID As Long, uElapse As Long)
   Call SetTimer(hwnd, TimerID, uElapse, AddressOf TimerProc)
End Sub

Private Sub TimerProc( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal idEvent As Long, _
    ByVal dwTime As Long _
)
    'pointer to class instance is in timerid
    Dim RefToCLS As clsConsolaWindows
    Set RefToCLS = ObjFromPtr(idEvent)
    RefToCLS.CLSTimerProc hwnd, idEvent, dwTime
End Sub

Private Function ObjFromPtr(ByVal lpObject As Long) As Object
Dim objTemp As Object
    CopyMemory objTemp, lpObject, 4&
    Set ObjFromPtr = objTemp
    CopyMemory objTemp, 0&, 4&
End Function


