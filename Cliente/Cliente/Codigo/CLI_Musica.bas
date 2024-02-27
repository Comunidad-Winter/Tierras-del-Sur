Attribute VB_Name = "CLI_Musica"
Option Explicit

'**** Used By MP3 Playing. *****
Public IMC   As IMediaControl
Public IBA   As IBasicAudio
Public IMPos As IMediaPosition

Public CurMidi As Integer

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub Play_Song(song_name As String)

    If SoundMute Then Exit Sub
    'on error Resume Next
    If volumen > 0 Then volumen = 0
    IBA.volume = volumen
    Set IMPos = IMC
    IMPos.CurrentPosition = 0
    IMC.Run

End Sub

Public Function Music_MP3_Load(ByVal file_path As String, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean '**** Loads a MP3 *****

    Set IMC = New FilgraphManager
    IMC.RenderFile file_path
    Set IBA = IMC
    Set IMPos = IMC

End Function
