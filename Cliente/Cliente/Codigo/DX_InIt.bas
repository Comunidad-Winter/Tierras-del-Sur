Attribute VB_Name = "Mod_DX"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Option Explicit

Public Const NumSoundBuffers = 20

Public DirectX As DirectX7
Public DirectDraw As DirectDraw7
Public DirectSound As DirectSound

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public SecundaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7



Public Perf As DirectMusicPerformance
Public Segs As DirectMusicSegment
Public Loader As DirectMusicLoader

Public oldResHeight As Long, oldResWidth As Long
Public bNoResChange As Boolean

Public Buffer(1 To NumSoundBuffers) As DirectSoundBuffer

Private Sub LiberarDirectSound()
Dim cloop As Integer
For cloop = 1 To NumSoundBuffers
    Set Buffer(cloop) = Nothing
Next cloop
Set DirectSound = Nothing
End Sub

Private Sub IniciarDXobject(DX As DirectX7)
Err.Clear
'on error Resume Next
Set DX = New DirectX7
If Err Then
    MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido por Set DX = New DirectX7"
    End
End If
End Sub

Private Sub IniciarDDobject(DD As DirectDraw7)
Err.Clear
'on error Resume Next
Set DD = DirectX.DirectDrawCreate("")
If Err Then
    MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido en Private Sub IniciarDDobject(DD As DirectDraw7)"
    End
End If
End Sub

Public Sub IniciarObjetosDirectX()
'on error Resume Next
Dim lRes As Long
Dim MidevM As typDevMODE
Dim CambiarResolucion As Boolean

Call IniciarDXobject(DirectX)

Call IniciarDDobject(DirectDraw)



Call Audio.Initialize(DirectX, frmMain.hWnd, App.Path & "\" & cDirWav & "\", App.Path & "\" & cDirMusica & "\")

lRes = EnumDisplaySettings(0, 0, MidevM)

oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
oldResHeight = Screen.Height \ Screen.TwipsPerPixelY

If NoRes Then
    CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
Else
    CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
End If
If CambiarResolucion Then
      With MidevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 16 ' bitsc
      End With
      lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
Else
    If NoRes And (oldResWidth <> 800 Or oldResHeight <> 600) Then
    frmMain.WindowState = 0
    frmMain.Width = 11910
    frmMain.Height = 9000
    frmMain.BorderStyle = 1
    'frmMain.ScaleHeight = 1
    End If
    bNoResChange = True
End If

CambiarColores (16) 'La verdad que esto no esta muy bien..

Exit Sub
End Sub

Public Sub LiberarObjetosDX()
Err.Clear
'on error GoTo fin:
Dim LoopC As Integer

Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

LiberarDirectSound

Call LiberarGraficos

Set DirectDraw = Nothing

For LoopC = 1 To NumSoundBuffers
    Set Buffer(LoopC) = Nothing
Next LoopC

Set Loader = Nothing
Set Perf = Nothing
Set Segs = Nothing
Set DirectSound = Nothing

Set DirectX = Nothing
Exit Sub
fin: LogError "Error producido en Public Sub LiberarObjetosDX()"
End Sub
