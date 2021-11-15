VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1560
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Winsock1.Close
Me.Winsock1.LocalPort = 888
Me.Winsock1.Listen
End Sub

Private Sub Command2_Click()
Me.Winsock1.SendData "AAA2"
End Sub

Private Sub Command3_Click()
Call PlayWaveDS("C:\2.wav")
End Sub

Private Sub Form_Load()
 Set DirectSound = DirectX.DirectSoundCreate("")
    If Err Then
        MsgBox "Error iniciando DirectSound"
        End
    End If
    
    LastSoundBufferUsed = 1
    '<----------------Direct Music--------------->
    Set Perf = DirectX.DirectMusicPerformanceCreate()
    Call Perf.Init(Nothing, 0)
    Perf.SetPort -1, 80
    Call Perf.SetMasterAutoDownload(True)
    '<------------------------------------------->
    Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Buffer(LastSoundBufferUsed).Play DSBPLAY_LOOPING
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Me.Winsock1.Close
Me.Winsock1.Accept requestID
Debug.Print requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim lucia As String
Me.Winsock1.GetData lucia
Me.Show
Select Case Left(lucia, 1)
Case "A"
'Reproduce Midis
lucia = Right(lucia, Len(lucia) - 1)

Call CargarMIDI(App.Path & "\Midi\" & lucia)

Case "V"
'Volumen Midis

Case "W"
'Reproduce wavs

Case "X"
End




Case Else
'wavs
    If Val(lucia) > 0 Then
    Call PlayWaveDS(lucia, ReadField(2, lucia, 44), ReadField(3, lucia, 44))
    Else
    PlayWaveDS (lucia)
    End If
Me.Hide
End Select

End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
Dim delimiter As String
delimiter = Chr(SepASCII)
Dim I As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    
    For I = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next I
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = Mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function
