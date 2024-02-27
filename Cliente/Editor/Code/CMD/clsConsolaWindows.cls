VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsolaWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsDOS - used to run console programs and return results
'Code by:
'   William Moeur
'   http://moeur.net
'
'   Joacim Andersson, Brixoft Software
'   http://www.brixoft.net

Option Explicit

Public Event Complete(stdOut As String, stdErr As String)
Private hSTDOUT As Long, hSTDERR As Long
Private Const BUFSIZE = 1024 ' pipe buffer size
Private pararProceso As Boolean

'Run the console program specified
'returns zero on success, otherwise returns error code
Public Function RunCommand(hwnd As Long, strCommand As String) As Long
    Dim hPipeWrite1 As Long, hPipeWrite2 As Long
    Dim hCurProcess As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    ReDim baOutput(BUFSIZE - 1) As Byte

    pararProceso = False
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1    ' get inheritable pipe handles
    End With

'create stdout pipe
    If CreatePipe(hSTDOUT, hPipeWrite1, sa, BUFSIZE) = 0 Then
        RunCommand = Err.LastDllError
        Exit Function
    End If

'create stderr pipe
    If CreatePipe(hSTDERR, hPipeWrite2, sa, BUFSIZE) = 0 Then
        RunCommand = Err.LastDllError
        Call CloseHandle(hPipeWrite1)
        Exit Function
    End If
    
    hCurProcess = GetCurrentProcess()

    ' Replace our inheritable read handles with an non-inheritable.
    'Not that it seems to be necessary in this case, but the docs
    'say we should.
    Call DuplicateHandle(hCurProcess, hSTDOUT, hCurProcess, hSTDOUT, 0&, _
                         0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)

    Call DuplicateHandle(hCurProcess, hSTDERR, hCurProcess, hSTDERR, 0&, _
                         0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)
    
    'startup info
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
         ' hide the window since all output is redirected anyway
         .wShowWindow = SW_HIDE
        .hStdOutput = hPipeWrite1
        .hStdError = hPipeWrite2
    End With

    If CreateProcess(vbNullString, strCommand, ByVal 0&, ByVal 0&, 1, 0&, _
        ByVal 0&, vbNullString, si, pi) = 0 Then
        RunCommand = Err.LastDllError
        Call CloseHandle(hPipeWrite1)
        Call CloseHandle(hPipeWrite2)
        Exit Function
    End If

        ' Close thread handle - we don't need it
        Call CloseHandle(pi.hThread)
        
        ' Also close our handle(s) to the write end of the pipe. This is _
        ' important, since ReadFile will *not* return until all write _
        ' handles are closed or the buffer is full.
        Call CloseHandle(hPipeWrite1)
        Call CloseHandle(hPipeWrite2)

        RunCommand = 0
        'start our timer to poll for stdout and stderr messages
        StartTimer hwnd, ObjPtr(Me), 100

End Function

'timer event, polls for stdout and stderr messages
'once any message arrives, our process is finished so
'shut down the timer and return the results in an event
Friend Sub CLSTimerProc(ByVal hwnd As Long, ByVal TimerID As Long, _
    ByVal dwTime As Long)
    Dim baOutput(BUFSIZE - 1) As Byte
    Dim sNewOutput As String
    Dim lBytesRead As Long
    Dim TotalBytesAvail As Long
    Dim BytesLeftThisMessage As Long
    Dim KillFlag As Boolean
    Dim strSTDOUT As String
    Dim strSTDERR As String
    
    
    'see if stdout has anything for us
    Call PeekNamedPipe(hSTDOUT, baOutput(0), BUFSIZE, lBytesRead, TotalBytesAvail, _
        BytesLeftThisMessage)
    If lBytesRead > 0 Then
        strSTDOUT = ReadPipe(hSTDOUT)
        KillFlag = True
    End If
    
    'see if stderr has anything for us
    Call PeekNamedPipe(hSTDERR, baOutput(0), BUFSIZE, lBytesRead, TotalBytesAvail, _
        BytesLeftThisMessage)
    If lBytesRead > 0 Or pararProceso Then
        strSTDERR = ReadPipe(hSTDERR)
        KillFlag = True
        'maybe stdout has something for us now if it didn't before
        If strSTDOUT = "" Then
            Call PeekNamedPipe(hSTDOUT, baOutput(0), BUFSIZE, lBytesRead, _
                TotalBytesAvail, BytesLeftThisMessage)
            If lBytesRead > 0 Then strSTDOUT = ReadPipe(hSTDOUT)
        End If
    End If
    
    If KillFlag Then
    'were done here
        KillTimer hwnd, TimerID
        Call CloseHandle(hSTDOUT)
        Call CloseHandle(hSTDERR)
        RaiseEvent Complete(strSTDOUT, strSTDERR)
    End If
    
End Sub

Public Sub parar()
    pararProceso = True
End Sub

Private Function ReadPipe(hPipe As Long) As String
    Dim baOutput(BUFSIZE - 1) As Byte
    Dim sNewOutput As String
    Dim lBytesRead As Long
    Dim CommandOutput As String
    
    CommandOutput = ""
    Do While ReadFile(hPipe, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) <> 0
        ' convert from "DOS" to "Windows" characters
        sNewOutput = String$(lBytesRead, 0)
        Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
        CommandOutput = CommandOutput & sNewOutput
    Loop
    ReadPipe = CommandOutput
End Function

