Attribute VB_Name = "Mod_Wav"
Option Explicit

Public Const SND_SYNC = &H0 ' SINCRONO
Public Const SND_ASYNC = &H1 ' ASINCRONO
Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound
Private Const PI As Single = 3.14159265358979 'Calculadora rulz
Private Const RAD As Single = PI / 180 'radiales
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?WAVS¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public Const SND_CLICK As String = "click.Wav"

Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE  As String = "cupdice.Wav"

Function LoadWavetoDSBuffer(sFile As String) As Boolean

On Local Error Resume Next

    Dim desc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX

    desc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    Set DS = DirectX.DirectSoundCreate("")
    DS.SetCooperativeLevel frmMain.hWnd, DSSCL_EXCLUSIVE
    desc.lFlags = (DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME) Or DSBCAPS_STATIC
    Set Buffer(LastSoundBufferUsed) = DS.CreateSoundBufferFromFile(App.Path & "\Wav\" & sFile, desc, waveFormat)
    Set Buffer3Ds = Buffer(LastSoundBufferUsed).GetDirectSound3DBuffer
    LoadWavetoDSBuffer = True
End Function

Sub PlayWaveDS(File As String, Optional ByRef CharIndex As Integer)
On Error GoTo jose10:

    If Fx = 1 Then Exit Sub
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    If LoadWavetoDSBuffer(File) Then
            Buffer(LastSoundBufferUsed).SetVolume -VolumeN
         If CharIndex = 0 Or CharIndex = UserCharIndex Then
            If CharIndex = UserCharIndex Then
            Buffer(LastSoundBufferUsed).SetVolume (-VolumeN - 500)
            Buffer(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
            Else
            Buffer(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
            End If
        Else
            If CharList(CharIndex).Nombre = "" Then
             Buffer(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
          Else
            Call en3d(CharIndex)
            End If
        End If
    End If
    Exit Sub
    
jose10:
End Sub

Private Sub en3d(CharIndex As Integer)

Dim Src_X As Single, Src_Y As Single, vDir As D3DVECTOR
Dim angulo As Single
Dim DiferenciaX  As Integer
Dim DiferenciaY As Integer
Dim Distancia As Integer


    DiferenciaX = CharList(CharIndex).Pos.X - CharList(UserCharIndex).Pos.X
    DiferenciaY = CharList(CharIndex).Pos.Y - CharList(UserCharIndex).Pos.Y

        If DiferenciaX = 0 And DiferenciaY > 0 Then
        angulo = 270
        ElseIf DiferenciaX = 0 And DiferenciaY < 0 Then
        angulo = 90
        ElseIf DiferenciaY = 0 And DiferenciaX > 0 Then
        angulo = 180
        ElseIf DiferenciaY = 0 And DiferenciaX < 0 Then
        angulo = 0
        Else
        angulo = Tan(DiferenciaY / DiferenciaX)
        End If

        Distancia = ((DiferenciaY ^ 2) / 2 + (DiferenciaX ^ 2) / 2) ^ 1 / 2
        Src_X = Cos(angulo * RAD) * (Distancia / 3)
        Src_Y = Sin(angulo * RAD) * (Distancia / 3)
        If frmOpciones.Check4 = 1 Then
        DiferenciaX = DiferenciaX * -1
        DiferenciaY = DiferenciaY * -1
        End If
        
       Call Buffer3Ds.SetPosition(-DiferenciaX, 0, -DiferenciaY, DS3D_IMMEDIATE)
        vDir.X = 0 - Src_X
        vDir.z = 0 - Src_Y
        If Src_X = 0 Then Src_X = 1
    
        'Configuracion 3D
        Buffer3Ds.SetConeOrientation Src_X, vDir.Y, vDir.z, DS3D_IMMEDIATE
        Buffer(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
        'Play!
End Sub

Private Sub Asinomas()
Buffer(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
End Sub
