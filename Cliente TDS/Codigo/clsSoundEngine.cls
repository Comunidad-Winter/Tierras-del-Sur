VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************* Sound Engine **************************
' Play Sounds....
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
' Frerik Alexandersson(fredrik@oraklet.zzn.com) - 20/05/2003
'   -The Basic Engine
'***********************************
    Dim Sound_Loaded As Boolean
    Dim Music_Loaded As Boolean
    
    Dim Resource_Dir As String
    Private Const Sounds_Dir = "\Sounds"
'**** Used By MP3 Playing. *****
    Dim IMC   As IMediaControl
    Dim IBA   As IBasicAudio
    Dim IME   As IMediaEvent
    Dim IMPos As IMediaPosition
'**** Direct X 8 **********
    Dim DX     As DirectX8
'**** Direct Sound ********
    Dim DS As DirectSound8
    Dim Sound_Buffer() As DirectSoundSecondaryBuffer8
    Dim DS_Enum As DirectSoundEnum8
'**** Direct Music ********
    Dim DMus_Preformance As DirectMusicPerformance8
    Dim DMus_Loader As DirectMusicLoader8
    Dim Music_Segment() As DirectMusicSegment8

Public Function Music_MP3_Empty() As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    
    If ObjPtr(IMC) > 0 Then
        IMC.Stop
    End If
    
    Set IBA = Nothing
    Set IME = Nothing
    Set IMPos = Nothing
    Set IMC = Nothing
    
    Music_MP3_Empty = True
    Exit Function

Error_Handl:
    Music_MP3_Empty = False
    End
End Function

Public Function Music_MP3_Load(ByVal file_path As String, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean '**** Loads a MP3 *****
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    
    If Not Music_MP3_Empty() = True Then GoTo Error_Handl
    
    Set IMC = New FilgraphManager
    IMC.RenderFile file_path
    
    Set IBA = IMC
volume = -4000
    If volume < -4000 Then volume = -4000
    IBA.volume = volume
    If balance > 5000 Then balance = 5000
    If balance < -5000 Then balance = -5000
    IBA.balance = balance
    
    Set IME = IMC
    
    Set IMPos = IMC
    If ObjPtr(IMPos) Then IMPos.Rate = 1#
    IMPos.CurrentPosition = 0
    
    Audio_MP3_Load = True
    Exit Function

Error_Handl:
    Music_MP3_Load = False
End Function

Public Function Music_MP3_Play() As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    
    IMC.Run
    
    Audio_MP3_Play = True
    Exit Function

Error_Handl:
    Music_MP3_Play = False
End Function

Public Function Music_MP3_Get_Loop() As Boolean
    On Error GoTo Error_Handl
    If IMPos.StopTime = IMPos.CurrentPosition Then
        Music_MP3_Stop
        Music_MP3_Play
    End If
Error_Handl:
End Function

Public Function Music_MP3_Stop() As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    
    IMC.Stop
    IMPos.CurrentPosition = 0
    
    Music_MP3_Stop = True
    Exit Function

Error_Handl:
    Music_MP3_Stop = False
End Function

Public Function Music_MP3_Pause() As Boolean
'***************************************************
' Created by: Juan Martín Sotuyo Dodero
' Modify Date: 25/6/2004
'***************************************************
    On Error GoTo Error_Handl
    
    IMC.Pause
    
    Music_MP3_Pause = True
    Exit Function

Error_Handl:
    Music_MP3_Pause = False
End Function

Public Function Music_MP3_Volume_Set(ByVal Sound_Volume As Long) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    
    If Sound_Volume > 0 Then Sound_Volume = 0
    If Sound_Volume < -4000 Then Sound_Volume = -4000
    
    IBA.volume = Sound_Volume

    Music_MP3_Volume_Set = True
    Exit Function

Error_Handl:
    Music_MP3_Volume_Set = False
End Function

Public Function Music_Midi_Play(ByVal Music_Number As Long, Optional ByVal Do_Loop As Boolean = True) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    
    If Do_Loop Then
        Music_Segment(Music_Number).SetRepeats -1
    Else
        Music_Segment(Music_Number).SetRepeats 1
    End If
    
    DMus_Preformance.PlaySegmentEx Music_Segment(Music_Number), DMUS_SEGF_DEFAULT, 0
    
    Music_Midi_Play = True
    Exit Function
    
Error_Handl:
    Music_Midi_Play = False
End Function

Public Function Music_Midi_Stop(ByVal Music_Number As Long) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    DMus_Preformance.StopEx Music_Segment(Music_Number), 0, DMUS_SEGF_DEFAULT
    Music_Midi_Stop = True
    Exit Function
    
Error_Handl:
End Function

Public Function Music_Midi_Volume_Set(ByVal volume As Long) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    If volume > 200 Then volume = 200
    If volume < -4000 Then volume = -4000
    
    DMus_Preformance.SetMasterVolume volume
    
    Music_Midi_Volume_Set = True
    Exit Function
    
Error_Handl:
End Function

Public Function Engine_Initialize(ByVal main_hwnd As String, ByVal Set_Resource_Dir As String, Optional ByVal Play_only_on_focus As Boolean = True, Optional ByVal Initialize_Sounds As Boolean = True, Optional ByVal Initialize_Music As Boolean = True) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
    On Error GoTo Error_Handl
    
    Set DX = New DirectX8
'****** Initializes Sounds ***********
Resource_Dir = Set_Resource_Dir
If Initialize_Sounds Then
    Dim bufferDesc As DSBUFFERDESC
    Set DS_Enum = DX.GetDSEnum
    Set DS = DX.DirectSoundCreate(DS_Enum.GetGuid(1))
    
    If Play_only_on_focus Then
        DS.SetCooperativeLevel main_hwnd, DSSCL_NORMAL
    Else
        DS.SetCooperativeLevel main_hwnd, DSSCL_PRIORITY
    End If
    '******* Loads all the wav files. **********
    ReDim Sound_Buffer(1 To 10) As DirectSoundSecondaryBuffer8
    
    Dim Sound_Number As Long
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    For Sound_Number = 1 To General_Var_Get(App.Path & "\sound.ini", "WAV", "NumWavs")
        If FileExist(Resource_Dir & Sounds_Dir & "\sound" & Sound_Number & ".wav", vbNormal) Then
            Set Sound_Buffer(Sound_Number) = DS.CreateSoundBufferFromFile(Resource_Dir & Sounds_Dir & "\sound" & Sound_Number & ".wav", bufferDesc)
        End If
    Next Sound_Number
    Sound_Loaded = True
End If

'****** Initializes Music ***********
If Initialize_Music Then
    Dim DMus_Params As DMUS_AUDIOPARAMS
    Set DMus_Preformance = DX.DirectMusicPerformanceCreate
    Set DMus_Loader = DX.DirectMusicLoaderCreate
    
    DMus_Preformance.InitAudio main_hwnd, DMUS_AUDIOF_ALL, DMus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    DMus_Preformance.SetMasterAutoDownload True

    DMus_Loader.SetSearchDirectory Sound_Dir
    
    '*** Loads all the midi files. ****
    ReDim Music_Segment(1 To General_Var_Get(App.Path & "\sound.ini", "MIDI", "NumMidi")) As DirectMusicSegment8
    Dim Music_Number As Long
    For Music_Number = 1 To General_Var_Get(App.Path & "\sound.ini", "MIDI", "NumMidi")
        If FileExist(Resource_Dir & Sounds_Dir & "\music" & Music_Number & ".mid", vbNormal) Then
            Set Music_Segment(Music_Number) = DMus_Loader.LoadSegment(Resource_Dir & Sounds_Dir & "\music" & Music_Number & ".mid")
            Music_Segment(Music_Number).SetStandardMidiFile
        End If
    Next Music_Number
    
    DMus_Preformance.SetMasterTempo 1
    DMus_Preformance.SetMasterVolume 1
    
    Music_Loaded = True
End If
    
    Sound_Initialize = True
    Exit Function

Error_Handl:
    Engine_Initialize = False
End Function

Public Function Sound_Play(ByVal Sound_Number As Long, ByVal Do_Loop As Boolean) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
On Error GoTo Error_Handl
    If Do_Loop Then
        Sound_Buffer(Sound_Number).Play DSBPLAY_LOOPING
    Else
        Sound_Buffer(Sound_Number).Play DSBPLAY_DEFAULT
    End If
    Sound_Play = True
    Exit Function
    
Error_Handl:
    Sound_Play = False
End Function

Public Function Sound_Pause(ByVal Sound_Number As Long) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
On Error GoTo Error_Handl
    Sound_Buffer(Sound_Number).Stop
    
    Sound_Pause = True
    Exit Function
Error_Handl:
    Sound_Pause = False
End Function

Public Function Sound_Stop(ByVal Sound_Number As Long) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
On Error GoTo Error_Handl

    Sound_Buffer(Sound_Number).Stop
    Sound_Buffer(Sound_Number).SetCurrentPosition 0
    
    Sound_Stop = True
    Exit Function
Error_Handl:
    Sound_Stop = False
End Function

Public Function Engine_DeInitialize() As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
On Error GoTo Error_Handl
    If Sound_Loaded Then
        Dim Sound_Number As Long

        For Sound_Number = 1 To General_Var_Get(App.Path & "\sound.ini", "WAV", "NumWavs")
            Set Sound_Buffer(Sound_Number) = Nothing
        Next Sound_Number
        Set DS_Enum = Nothing
        Set DS = Nothing
    End If

    If Music_Loaded Then
        DMus_Preformance.RemoveNotificationType DMUS_NOTIFY_ON_SEGMENT
        Dim Music_Number As Long
        For Music_Number = 1 To General_Var_Get(App.Path & "\sound.ini", "MIDI", "NumMidi")
            Set Music_Segment(Music_Number) = Nothing
        Next Music_Number
        Set DMus_Loader = Nothing
        If Not (DMus_Preformance Is Nothing) Then
            DMus_Preformance.CloseDown
            Set DMus_Preformance = Nothing
        End If
    End If
    Set DX = Nothing

    Engine_DeInitialize = True
    Exit Function

Error_Handl:
    Engine_DeInitialize = False
End Function

Public Function Sound_Volume_Set(ByVal Sound_Number As Long, ByVal volume As Long) As Boolean
'**************************************
'* Created by: Fredrik Alexandersson  *
'* Modify Date: 20/5/2003             *
'**************************************
On Error GoTo Error_Handl

If volume > 0 Then volume = 0
If volume < -4000 Then volume = -4000
Sound_Buffer(Sound_Number).SetVolume volume

Error_Handl:
    Sound_Volume_Set = False
End Function
