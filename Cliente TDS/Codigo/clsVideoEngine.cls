VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************* Video Engine **************************
' Reproduces video and audio
' Plays .mpg; .avi; .mov; .wav; .mp2; .mp3 files
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************
'
' Juan Martín Sotuyo Dodero(juansotuyo@hotmail.com) - 1/14/2004
'   -First Release
'*****************************************************************
Option Explicit

'Time we must wait before checking again if video has ended (in milisecs.)
Private Const TIME_BETWEEN_CHECKS = 250

Dim FileName As String        'Loaded Filename
Dim length As Double          'Duration in seconds
Dim Playing As Boolean        'Flag used to trigger internal clock
Dim LoopVideo As Boolean      'Flag used to loop a video
Dim StartPos As Double        'Determines the point from which to start reproducing

Dim engine_initialized As Boolean

'The output screen handler
Dim ViewhWnd As Long

'The output screen dimensions
Dim ViewLeft As Long
Dim ViewTop As Long
Dim ViewWidth As Long
Dim ViewHeight As Long

'Last check time
Dim LastCheck As Long

Dim Audio  As IBasicAudio
Dim Video As IBasicVideo
Dim MediaEvent As IMediaEvent
Dim VideoWindow As IVideoWindow
Dim MediaPos As IMediaPosition
Dim MediaControl As IMediaControl

'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function Engine_Initialize(ByVal v_hWnd As Long, ByVal v_width_in_pixels As Long, ByVal v_height_in_pixels As Long, Optional ByVal v_left_in_pixels As Long = 0, Optional ByVal v_top_in_pixels As Long = 0) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
On Local Error GoTo ErrHandler

    ViewhWnd = v_hWnd
    ViewLeft = v_left_in_pixels
    ViewTop = v_top_in_pixels
    ViewHeight = v_height_in_pixels
    ViewWidth = v_width_in_pixels
    
    engine_initialized = True
    
    Engine_Initialize = True
Exit Function

ErrHandler:
    MsgBox "Error in Video Engine Initialization: " & Err.Number & ": " & Err.Description
    engine_initialized = False
    Engine_Initialize = False
End Function

Public Sub Engine_DeInitialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    'Call clean-up method
    CleanUp
    
    engine_initialized = False
End Sub

Public Function File_Load(ByVal file_path As String, Optional ByVal fullscreen As Boolean = False, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0, Optional ByVal loop_video As Boolean = False, Optional ByVal hide_cursor As Boolean = False) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Sets all objects to reproduce the given file and starts reproducing.
'**************************************************************
    Dim nCount As Long
    
On Local Error GoTo ErrHandler
    
    'Make sure engine was properly initialized
    If Not engine_initialized Then
        File_Load = False
        Exit Function
    End If
    
    'Clean-up to avoid errors
    CleanUp
    
    'Make sure file exists
    If Dir(file_path) = "" Then
        File_Load = False
        Exit Function
    End If
    
    FileName = file_path
    
    'Instantiate a filter graph for the requested file format
    Set MediaControl = New FilgraphManager
    MediaControl.RenderFile FileName
    
    Set Audio = MediaControl
    
    'Set volume
    If volume < -4000 Then volume = -4000
    Audio.volume = volume
    
    'Set balance
    If balance > 5000 Then balance = 5000
    If balance < -5000 Then balance = -5000
    Audio.balance = balance
    
    'Setup the IVideoWindow object
    Set VideoWindow = MediaControl
    
    VideoWindow.WindowStyle = CLng(&H6000000)
    VideoWindow.top = ViewTop
    VideoWindow.left = ViewLeft
    If fullscreen Then
        'Fullscreen mode automatically sets the height and width to the ideal
        'height and width for the video
        VideoWindow.FullScreenMode = 1
    Else
        VideoWindow.width = ViewWidth
        VideoWindow.height = ViewHeight
    End If
    'Check if we have to hide the cursor
    VideoWindow.HideCursor CLng(hide_cursor)
    
    'reset the video window owner
    VideoWindow.Owner = ViewhWnd
    
    'Setup the IMediaEvent object
    Set MediaEvent = MediaControl
    
    'Setup the IMediaPosition object so that we
    'can display the duration of the selected
    'video as well as the elapsed time
    Set MediaPos = MediaControl
    
    MediaPos.Rate = 1
    
    length = Round(MediaPos.Duration, 2)
    
    'Set loop
    LoopVideo = loop_video
    
    'Reset start position to 0
    StartPos = 0
    
    File_Play
    
    File_Load = True
Exit Function
            
ErrHandler:
    File_Load = False
End Function

Public Function File_Play() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Starts/continues reproducing the file.
'**************************************************************
    'Make sure there is a file loaded (therefore the engine was initialized)
    If FileName = "" Then
        File_Play = False
        Exit Function
    End If
    
    'Assign specified starting position dependent on state
    If CLng(MediaPos.CurrentPosition) < CLng(StartPos) Then
        MediaPos.CurrentPos = StartPos
    ElseIf CLng(MediaPos.CurrentPosition) = CLng(length) Then
        MediaPos.CurrentPos = StartPos
    End If
    
    Call MediaControl.Run
    Playing = True
    
    File_Play = True
Exit Function
            
ErrHandler:
    Playing = False
    File_Play = False
End Function

Public Function File_Pause() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Pauses file reproduction.
'**************************************************************
    If Not MediaControl Is Nothing Then
        Call MediaControl.Pause
        Dim a As Long
        a = MediaPos.CurrentPosition
        Playing = False
        'Show is over, show cursor back
        Dim cursor_hidden As Long
        VideoWindow.IsCursorHidden cursor_hidden
        If cursor_hidden Then
            VideoWindow.HideCursor 0
        End If
        File_Pause = True
        Exit Function
    End If
    
    File_Pause = False
End Function

Public Function File_Stop() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Stops reproducing the file.
'**************************************************************
    If Not MediaControl Is Nothing Then
        Call MediaControl.Stop
        'Reset to the beginning of the file
        MediaPos.CurrentPosition = 0
        Playing = False
        'Show is over, show cursor back
        Dim cursor_hidden As Long
        VideoWindow.IsCursorHidden cursor_hidden
        If cursor_hidden Then
            VideoWindow.HideCursor 0
        End If
        File_Stop = True
        Exit Function
    End If
    
    File_Stop = False
End Function

Public Function Balance_Change(ByVal balance As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    If Not MediaControl Is Nothing Then
        If balance > 5000 Then balance = 5000
        If balance < -5000 Then balance = -5000
        Audio.balance = balance
        Balance_Change = True
        Exit Function
    End If
    
    Balance_Change = False
End Function
 
 Public Function Volume_Change(ByVal volume As Long) As Boolean
 '**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    'Set the volume using the slider
    If Not MediaControl Is Nothing Then
        If volume < -4000 Then volume = -4000
        Audio.volume = volume
        Volume_Change = True
        Exit Function
    End If
    
    Volume_Change = False
End Function
            
Public Function File_Check_If_Ended() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Returns True if the file has ended (even if it loops), false otherwise.
'**************************************************************
    Dim ReturnCode As Long

On Local Error GoTo ErrHandler
    
    If Playing = True Then
        
        If LastCheck + TIME_BETWEEN_CHECKS < GetTickCount Then
            LastCheck = GetTickCount
        Else
            File_Check_If_Ended = False
            Exit Function
        End If
        
        'Obtain return code
        If Not MediaEvent Is Nothing Then
            Call MediaEvent.WaitForCompletion(100, ReturnCode)
            
            If ReturnCode = 1 Then
                'It has finished
                Playing = False
                
                'Check if we have to loop
                If LoopVideo Then
                    File_Play
                Else
                    'Show is over, show cursor back
                    Dim cursor_hidden As Long
                    VideoWindow.IsCursorHidden cursor_hidden
                    If cursor_hidden Then
                        VideoWindow.HideCursor 0
                    End If
                End If
                
                File_Check_If_Ended = True
            Else
                File_Check_If_Ended = False
            End If
        End If
        Exit Function
    End If
    
ErrHandler:
    Err.Clear
    File_Check_If_Ended = False
End Function

Public Sub Start_Pos_Set(ByVal start_pos As Double)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Sets the pos from which to start reproducing.
'**************************************************************
    StartPos = start_pos
End Sub

Public Function File_GoTo(ByVal pos As Double) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Moves to the given pos within the file.
'**************************************************************
On Local Error GoTo ErrHandler
    If length = 0 Then
        File_GoTo = False
        Exit Function
    End If
    
    If pos >= length Then pos = length
    
    If Playing Then
        'Stop playing to avoid errors
        File_Pause
        MediaPos.CurrentPosition = pos
        'Continue playing
        File_Play
    Else
        MediaPos.CurrentPosition = pos
    End If
    
    File_GoTo = True
    
Exit Function

ErrHandler:
    File_GoTo = False
End Function

Public Function File_Get_Length() As Double
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Returns the currently loaded file´s length. -1 if no file is loaded
'**************************************************************
    If length > 0 Then
        File_Get_Length = length
    Else
        File_Get_Length = -1
    End If
End Function

Public Function Volume_Get() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    If Not Audio Is Nothing Then
        Volume_Get = Audio.volume
    End If
End Function

Public Function Balance_Get() As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    If Not Audio Is Nothing Then
        Balance_Get = Audio.balance
    End If
End Function

Public Function File_Get_Elapsed_Time() As Double
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    If Not MediaPos Is Nothing Then
        File_Get_Elapsed_Time = MediaPos.CurrentPosition
    End If
End Function

Public Function File_Get_Name() As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    File_Get_Name = FileName
End Function

Public Function File_Get_Loop() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'
'**************************************************************
    File_Get_Loop = LoopVideo
End Function

Private Sub CleanUp()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Makes sure everything is ready to load a new file
'**************************************************************
On Local Error GoTo ErrHandler
    
    Playing = False
    
    'Erase the file name
    FileName = ""
    
    'Show is over, show cursor back
    Dim cursor_hidden As Long
    VideoWindow.IsCursorHidden cursor_hidden
    If cursor_hidden Then
        VideoWindow.HideCursor 0
    End If
    
    'cleanup media control
    If Not MediaControl Is Nothing Then
        MediaControl.Stop
    End If
    'clean-up video window
    If Not VideoWindow Is Nothing Then
        VideoWindow.left = Screen.width * 8
        VideoWindow.height = Screen.height * 8
        VideoWindow.Owner = 0          'sets the Owner to NULL
    End If
    
    'clean-up & dereference
    If Not Audio Is Nothing Then Set Audio = Nothing
    If Not Video Is Nothing Then Set Video = Nothing
    If Not MediaControl Is Nothing Then Set MediaControl = Nothing
    If Not VideoWindow Is Nothing Then Set VideoWindow = Nothing
    If Not MediaPos Is Nothing Then Set MediaPos = Nothing
Exit Sub

ErrHandler:
End Sub

Public Sub DestHwnd_Set(ByVal dest_hwnd As Long, ByVal v_width_in_pixels As Long, ByVal v_height_in_pixels As Long, Optional ByVal v_left_in_pixels As Long = 0, Optional ByVal v_top_in_pixels As Long = 0)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 1/15/2004
'Changes the output window
'**************************************************************
    ViewhWnd = dest_hwnd
    ViewLeft = v_left_in_pixels
    ViewTop = v_top_in_pixels
    ViewWidth = v_width_in_pixels
    ViewHeight = v_height_in_pixels
End Sub
