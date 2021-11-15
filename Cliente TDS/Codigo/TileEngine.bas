Attribute VB_Name = "Mod_TileEngine"
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
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    C       O       N       S      T
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize  As Byte = 1
Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    T       I      P      O      S
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Integer
    Y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type

'Lista de cuerpos
Public Type FxData
    Fx As Grh
    OffsetX As Long
    OffsetY As Long
End Type

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As Byte
    Pos As Position
    Difpos As Position
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    Fx As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Nombre As String
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    'ME Only
    changed As Byte
End Type

Public IniPath As String
Public MapPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public LastTime As Long 'Para controlar la velocidad

'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT
Public MainViewWidth As Integer
Public MainViewHeight As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh 'Animaciones publicas
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public CharList(1 To 10000) As Char
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Sonido
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'está raineando?
'[Misery_Ezequiel 10/07/05]
Public bSnow        As Boolean 'está nevando?
'[\]Misery_Ezequiel 10/07/05]
Public bRainST      As Boolean
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long
Public bNoche       As Boolean 'es de noche?

'[Misery_Ezequiel 10/07/05]
Private RNieva(7)  As RECT  'RECT de la nieve
'[\]Misery_Ezequiel 10/07/05]

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

'[Misery_Ezequiel 10/07/05]
Private LTNieva(4) As Integer
'[\]Misery_Ezequiel 10/07/05]

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
    Public Enum PlayLoop
        plNone = 0
        plLluviain = 1
        plLluviaout = 2
        plFogata = 3
    End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

''int VBABDX_API
''BltAlphaFast( int lpDDSDest, int lpDDSSource, int iWidth, int iHeight,
''             int pitchSrc, int pitchDst, DWORD dwMode )

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

Sub CargarCabezas()
On Error Resume Next
Dim N As Integer, I As Integer, Numheads As Integer, Index As Integer
Dim Miscabezas() As tIndiceCabeza

'<gorlok> 2005-03-28

'</gorlok> 2005-03-28
N = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #N
'cabecera
Get #N, , MiCabecera
'num de cabezas
Get #N, , Numheads
'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For I = 1 To Numheads
    Get #N, , Miscabezas(I)
    InitGrh HeadData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh HeadData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh HeadData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh HeadData(I).Head(4), Miscabezas(I).Head(4), 0
Next I
Close #N
End Sub

Sub CargarCascos()
On Error Resume Next
Dim N As Integer, I As Integer, NumCascos As Integer, Index As Integer
Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #N
'cabecera
Get #N, , MiCabecera
'num de cabezas
Get #N, , NumCascos
'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza
For I = 1 To NumCascos
    Get #N, , Miscabezas(I)
    InitGrh CascoAnimData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh CascoAnimData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh CascoAnimData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh CascoAnimData(I).Head(4), Miscabezas(I).Head(4), 0
Next I
Close #N
End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim N As Integer, I As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

N = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #N
'cabecera
Get #N, , MiCabecera
'num de cabezas
Get #N, , NumCuerpos
'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo
For I = 1 To NumCuerpos
    Get #N, , MisCuerpos(I)
    InitGrh BodyData(I).Walk(1), MisCuerpos(I).Body(1), 0
    InitGrh BodyData(I).Walk(2), MisCuerpos(I).Body(2), 0
    InitGrh BodyData(I).Walk(3), MisCuerpos(I).Body(3), 0
    InitGrh BodyData(I).Walk(4), MisCuerpos(I).Body(4), 0
    BodyData(I).HeadOffset.X = MisCuerpos(I).HeadOffsetX
    BodyData(I).HeadOffset.Y = MisCuerpos(I).HeadOffsetY
Next I
Close #N
End Sub

Sub CargarFxs()
On Error Resume Next
Dim N As Integer, I As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

N = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #N
'cabecera
Get #N, , MiCabecera
'num de cabezas
Get #N, , NumFxs
'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx
For I = 1 To NumFxs
    Get #N, , MisFxs(I)
    Call InitGrh(FxData(I).Fx, MisFxs(I).Animacion, 1)
    FxData(I).OffsetX = MisFxs(I).OffsetX
    FxData(I).OffsetY = MisFxs(I).OffsetY
Next I
Close #N
End Sub

Sub CargarTips()
On Error Resume Next
Dim N As Integer, I As Integer
Dim NumTips As Integer

N = FreeFile
Open App.Path & "\init\Tips.ayu" For Binary Access Read As #N
'cabecera
Get #N, , MiCabecera
'num de cabezas
Get #N, , NumTips
'Resize array
ReDim Tips(1 To NumTips) As String * 255
For I = 1 To NumTips
    Get #N, , Tips(I)
Next I
Close #N
End Sub

Sub CargarArrayLluvia()
On Error Resume Next
Dim N As Integer, I As Integer
Dim Nu As Integer

N = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #N
'cabecera
Get #N, , MiCabecera
'num de cabezas
Get #N, , Nu
'Resize array
'ReDim bLluvia(1 To Nu) As Byte
'For I = 1 To Nu
'    Get #N, , bLluvia(I)
'Next I
Close #N

ReDim bLluvia(1 To Nu) As Byte
bLluvia(1) = 1
bLluvia(2) = 1
bLluvia(3) = 1
bLluvia(4) = 1
bLluvia(5) = 1
bLluvia(6) = 1
bLluvia(7) = 1
bLluvia(8) = 1
bLluvia(9) = 1
bLluvia(10) = 1
bLluvia(11) = 1
bLluvia(12) = 1
bLluvia(13) = 1
bLluvia(14) = 1
bLluvia(15) = 1
bLluvia(16) = 1
bLluvia(17) = 1
bLluvia(18) = 1
bLluvia(19) = 1
bLluvia(20) = 1
bLluvia(21) = 1
bLluvia(22) = 1
bLluvia(23) = 1
bLluvia(24) = 1
bLluvia(25) = 1
bLluvia(26) = 1
bLluvia(27) = 1
bLluvia(28) = 1
bLluvia(29) = 1
bLluvia(30) = 1
bLluvia(31) = 1
bLluvia(32) = 1
bLluvia(34) = 1
bLluvia(35) = 1
bLluvia(36) = 1
bLluvia(38) = 1
bLluvia(39) = 1
bLluvia(46) = 1
bLluvia(47) = 1
bLluvia(53) = 1
bLluvia(54) = 1
bLluvia(55) = 1
bLluvia(56) = 1
bLluvia(57) = 1
bLluvia(58) = 1
bLluvia(59) = 1
bLluvia(60) = 1
bLluvia(61) = 1
bLluvia(62) = 1
bLluvia(63) = 1
bLluvia(64) = 1
bLluvia(65) = 1
bLluvia(66) = 1
bLluvia(67) = 1
bLluvia(68) = 1
bLluvia(69) = 1
bLluvia(70) = 1
bLluvia(71) = 1
bLluvia(72) = 1
bLluvia(73) = 1
bLluvia(74) = 1
bLluvia(75) = 1
bLluvia(76) = 1
bLluvia(77) = 0
bLluvia(78) = 1
bLluvia(79) = 1
bLluvia(80) = 1
bLluvia(81) = 1
bLluvia(82) = 1
bLluvia(83) = 1
bLluvia(84) = 1
bLluvia(85) = 1
bLluvia(86) = 1
bLluvia(87) = 1
bLluvia(88) = 1
bLluvia(89) = 1
bLluvia(90) = 1
bLluvia(91) = 1
bLluvia(92) = 1
bLluvia(93) = 1
bLluvia(94) = 1
bLluvia(95) = 1
bLluvia(96) = 1
bLluvia(97) = 1
bLluvia(98) = 1
bLluvia(99) = 1
bLluvia(100) = 1
bLluvia(101) = 1
bLluvia(102) = 1
bLluvia(103) = 1
bLluvia(104) = 1
bLluvia(105) = 1
bLluvia(106) = 1
bLluvia(107) = 1
bLluvia(108) = 1
bLluvia(109) = 1
bLluvia(110) = 1
bLluvia(111) = 1
bLluvia(112) = 1
bLluvia(113) = 1
bLluvia(114) = 1
bLluvia(117) = 1
bLluvia(118) = 1
bLluvia(119) = 1
bLluvia(120) = 1
bLluvia(121) = 1
bLluvia(122) = 1
bLluvia(123) = 1
bLluvia(124) = 1
bLluvia(125) = 1
bLluvia(126) = 1
bLluvia(127) = 1
bLluvia(128) = 1
bLluvia(129) = 1
bLluvia(130) = 1
bLluvia(131) = 1
bLluvia(131) = 1
bLluvia(132) = 1
bLluvia(133) = 1
bLluvia(134) = 1
bLluvia(135) = 1
bLluvia(136) = 1
bLluvia(137) = 1
bLluvia(138) = 1
bLluvia(139) = 1
bLluvia(147) = 1
bLluvia(148) = 1
bLluvia(149) = 1
bLluvia(150) = 1
bLluvia(151) = 1
bLluvia(152) = 1
bLluvia(153) = 1
bLluvia(154) = 1
bLluvia(155) = 1
bLluvia(156) = 1
bLluvia(157) = 1
bLluvia(158) = 1
bLluvia(159) = 1
bLluvia(160) = 1
bLluvia(161) = 1
bLluvia(162) = 1
bLluvia(173) = 1
bLluvia(177) = 1

ReDim bNieva(1 To Nu) As Byte

bNieva(170) = 1
bNieva(169) = 1
bNieva(171) = 1

End Sub


Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)
Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If
If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If
tX = UserPos.X + CX
tY = UserPos.Y + CY
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

'Apuntamos al ultimo Char
If CharIndex > LastChar Then LastChar = CharIndex

NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2
CharList(CharIndex).iHead = Head
CharList(CharIndex).iBody = Body
CharList(CharIndex).Head = HeadData(Head)
CharList(CharIndex).Body = BodyData(Body)
CharList(CharIndex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
CharList(CharIndex).Arma.WeaponAttack = 0
CharList(CharIndex).Escudo = ShieldAnimData(Escudo)
CharList(CharIndex).Casco = CascoAnimData(Casco)
CharList(CharIndex).Heading = Heading
'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0
'Update position
CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.Y = Y
'Make active
CharList(CharIndex).Active = 1
'Plot on map
MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
CharList(CharIndex).Active = 0
CharList(CharIndex).Criminal = 0
CharList(CharIndex).Fx = 0
CharList(CharIndex).FxLoopTimes = 0
CharList(CharIndex).invisible = False
CharList(CharIndex).Moving = 0
CharList(CharIndex).muerto = False
CharList(CharIndex).Nombre = ""
CharList(CharIndex).pie = False
CharList(CharIndex).Pos.X = 0
CharList(CharIndex).Pos.Y = 0
CharList(CharIndex).UsandoArma = False
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
CharList(CharIndex).Active = 0
'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)

'Update NumChars
NumChars = NumChars - 1
End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If
Grh.FrameCounter = 1
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurría debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
'[END]'
End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y
'Figure out which way to move
Select Case nHeading
    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)
CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)

End Sub

Public Sub DoFogataFx()
If Fx = 0 Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then frmMain.StopSound
    Else
        bFogata = HayFogata()
        If bFogata Then frmMain.Play "fuego.wav", True
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean
'    Dim X As Integer, Y As Integer
'    For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
'        For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
'            If MapData(X, Y).CharIndex = Index2 Then
'                EstaPCarea = True
'                Exit Function
'            End If
'        Next X
'    Next Y
'    EstaPCarea = False
    Dim uX As Integer, uY As Integer, CX As Integer, CY As Integer
    uX = UserPos.X
    uY = UserPos.Y
    CX = CharList(Index2).Pos.X
    CY = CharList(Index2).Pos.Y
    EstaPCarea = (Abs(CX - uX) <= (WindowTileWidth \ 2)) And _
        (Abs(CY - uY) <= (WindowTileHeight \ 2))
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    Static pie As Boolean
    If Not UserNavegando Then
        If Not CharList(CharIndex).muerto And EstaPCarea(CharIndex) Then
            CharList(CharIndex).pie = Not CharList(CharIndex).pie
            If CharList(CharIndex).pie Then
                Call PlayWaveDS(SND_PASOS1, CharIndex)
            Else
                Call PlayWaveDS(SND_PASOS2, CharIndex)
            End If
        End If
    Else
        Call PlayWaveDS(SND_NAVEGANDO, CharIndex)
    End If
End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If
If Sgn(addX) = -1 Then
    nHeading = WEST
End If
If Sgn(addY) = -1 Then
    nHeading = NORTH
End If
If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If
MapData(nX, nY).CharIndex = CharIndex

CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveScreen(Heading As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case Heading
    Case NORTH
        Y = -1
    Case EAST
        X = 1
    Case SOUTH
        Y = 1
    Case WEST
        X = -1
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y
'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If
End Sub

Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean
loopc = 1
Do While CharList(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(CharList))
Loop
NextOpenChar = loopc
End Function

Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************
On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim TempInt As Integer

'Resize arrays
ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1
Get #1, , MiCabecera
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
'Fill Grh List

'Get first Grh Number
Get #1, , Grh
Do Until Grh <= 0
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    If GrhData(Grh).NumFrames > 1 Then
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        Next Frame
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    Else
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        GrhData(Grh).Frames(1) = Grh
    End If
    'Get Next Grh Number
    Get #1, , Grh
Loop
'************************************************

Close #1
Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh
End Sub

Function LegalPos(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
LegalPos = True
End Function

Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If
InMapLegalBounds = True
End Function

Function InMapBounds(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If
InMapBounds = True
End Function

Sub DDrawGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If
'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'Center Grh over X,Y pos
If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
        .left = GrhData(CurrentGrh.GrhIndex).sX
        .top = GrhData(CurrentGrh.GrhIndex).sY
        .right = .left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .bottom = .top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
surface.BltFast X, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(surface As DirectDrawSurface7, Grh As Integer, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .left = X
    .top = Y
    .right = .left + GrhData(Grh).pixelWidth
    .bottom = .top + GrhData(Grh).pixelHeight
End With

surface.GetSurfaceDesc SurfaceDesc
'Draw
If destRect.left >= 0 And destRect.top >= 0 And destRect.right <= SurfaceDesc.lWidth And destRect.bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .left = GrhData(Grh).sX
        .top = GrhData(Grh).sY
        .right = .left + GrhData(Grh).pixelWidth
        .bottom = .top + GrhData(Grh).pixelHeight
    End With
    
    surface.BltFast destRect.left, destRect.top, SurfaceDB.GetBMP(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If
End Sub

'Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[CODE 000]:MatuX
    Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                CharList(KillAnim).Fx = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub
'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
If iGrhIndex = 0 Then Exit Sub
'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .left = GrhData(iGrhIndex).sX
    .top = GrhData(iGrhIndex).sY
    .right = .left + GrhData(iGrhIndex).pixelWidth
    .bottom = .top + GrhData(iGrhIndex).pixelHeight
End With

surface.BltFast X, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End Sub

#If ConAlfaB = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                CharList(KillAnim).Fx = 0
                                Exit Sub
                            End If

                        End If
                    End If
               End If
            End If
        End If
    End If
End If
If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .right = .left + GrhData(iGrhIndex).pixelWidth
    .bottom = .top + GrhData(iGrhIndex).pixelHeight
End With
'surface.BltFast X, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Dim src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set src = SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum, 0)
src.GetSurfaceDesc ddsdSrc
surface.GetSurfaceDesc ddsdDest
With rDest
    .left = X
    .top = Y
    .right = X + GrhData(iGrhIndex).pixelWidth
    .bottom = Y + GrhData(iGrhIndex).pixelHeight
    
    If .right > ddsdDest.lWidth Then
        .right = ddsdDest.lWidth
    End If
    If .bottom > ddsdDest.lHeight Then
        .bottom = ddsdDest.lHeight
    End If
End With
' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    surface.BltFast X, Y, src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False
On Local Error GoTo HayErrorAlpha

src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

surface.GetLockedArray dArray()
src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.left * 2, SourceRect.top)), rDest.right - rDest.left, rDest.bottom - rDest.top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)
surface.Unlock rDest
DstLock = False
src.Unlock SourceRect
SrcLock = False
Exit Sub

HayErrorAlpha:
If SrcLock Then src.Unlock SourceRect
If DstLock Then surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hWnd As Long, Hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
If Grh <= 0 Then Exit Sub
SecundaryClipper.SetHWnd hWnd
SurfaceDB.GetBMP(GrhData(Grh).FileNum).BltToDC Hdc, SourceRect, destRect
End Sub

Sub PlayWaveAPI(File As String)
'*****************************************************************
'Plays a Wave using windows APIs
'*****************************************************************

End Sub

Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'On Error Resume Next
On Error GoTo errh
Dim EEE As Integer
EEE = 0

If UserCiego Then Exit Sub

Dim Y        As Integer 'Keeps track of where on map we are
Dim X        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim Moved    As Byte
Dim Grh      As Grh     'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs

'Figure out Ends and Starts of screen
' Hardcodeado para speed!
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)

'Draw floor layer
ScreenY = 8 + RenderMod.iImageSize
For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 8) - RenderMod.iImageSize
    ScreenX = 8 + RenderMod.iImageSize
    For X = (minX + 8) + RenderMod.iImageSize To (maxX - 8) - RenderMod.iImageSize
        If X > 100 Or Y < 1 Then Exit For
        'Layer 1 **********************************
        With MapData(X, Y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                            .FrameCounter = 1
                    End If
                End If
            End If

            'Figure out what frame to draw (always 1 if not animated)
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With

        rSourceRect.left = GrhData(iGrhIndex).sX
        rSourceRect.top = GrhData(iGrhIndex).sY
        rSourceRect.right = rSourceRect.left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.bottom = rSourceRect.top + GrhData(iGrhIndex).pixelHeight
        'El width fue hardcodeado para speed!
        Call BackBufferSurface.BltFast( _
                ((32 * ScreenX) - 32) + PixelOffsetX, _
                ((32 * ScreenY) - 32) + PixelOffsetY, _
                SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), _
                rSourceRect, _
                DDBLTFAST_WAIT)
        '******************************************
        If Not RenderMod.bNoCostas Then
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(2), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, _
                        1)
            End If
            '******************************************
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y

'Draw Transparent Layers  (Layer 2, 3)
ScreenY = 8 + RenderMod.iImageSize
For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
    ScreenX = 5 + RenderMod.iImageSize
    For X = (minX + 5) + RenderMod.iImageSize To (maxX - 5) - RenderMod.iImageSize
        'If X > 100 Or X < -3 Then Exit For 'Gorlok: VIEJO BUG !!!
        If X > 0 And X < 101 Then
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
            'Object Layer **********************************
            If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
    '            If Y > UserPos.Y Then
    '                Call DDrawTransGrhtoSurfaceAlpha( _
    '                        BackBufferSurface, _
    '                        MapData(X, Y).ObjGrh, _
    '                        iPPx, iPPy, 1, 1)
    '            Else
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(X, Y).ObjGrh, _
                            iPPx, iPPy, 1, 1)
    '            End If
            End If
            '***********************************************
            'Char layer ************************************
            'If MapData(X, Y).CharIndex <> 0 Then
            If MapData(X, Y).CharIndex > 0 Then
                TempChar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0
                'If needed, move left and right
                If TempChar.MoveOffset.X <> 0 Then
                    TempChar.Body.Walk(TempChar.Heading).Started = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                    PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                    TempChar.MoveOffset.X = TempChar.MoveOffset.X - (8 * Sgn(TempChar.MoveOffset.X))
                    Moved = 1
                End If
                'If needed, move up and down
                If TempChar.MoveOffset.Y <> 0 Then
                    TempChar.Body.Walk(TempChar.Heading).Started = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                    PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                    TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (8 * Sgn(TempChar.MoveOffset.Y))
                    Moved = 1
                End If
                'If done moving stop animation
                If Moved = 0 And TempChar.Moving = 1 Then
                    TempChar.Moving = 0
                    TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                    TempChar.Body.Walk(TempChar.Heading).Started = 0
                    TempChar.Arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).FrameCounter = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 0
                End If
                '[ANIM ATAK]
                If TempChar.Arma.WeaponAttack > 0 Then
                    TempChar.Arma.WeaponAttack = TempChar.Arma.WeaponAttack - 1
                    If TempChar.Arma.WeaponAttack = 0 Then
                        TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                    End If
                End If
                '[/ANIM ATAK]
                
                'Dibuja solamente players
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
                If TempChar.Heading < 1 Or TempChar.Heading > 4 Then 'by Gorlok 17/11/2005
                    'Hacer huevo...
                    'Si Heading es 0, fallan varias de las lineas siguientes
                    'y se caen los FPS a pedazos...
                ElseIf TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                    If Not CharList(MapData(X, Y).CharIndex).invisible Then
                        '[CUERPO]'
#If (ConAlfaB = 1) Then
                        If TempChar.iBody = 8 Or TempChar.iBody = 145 Then
                            Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                        Else
#End If
                            Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
#If ConAlfaB = 1 Then
                        End If
#End If
                        '[END]'
                        
                        
                        '[CABEZA]'
#If ConAlfaB = 1 Then
                        If TempChar.iHead = 500 Or TempChar.iHead = 501 Then
                            Call DDrawTransGrhtoSurfaceAlpha( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.X, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
                        Else
#End If
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.X, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
#If ConAlfaB = 1 Then
                        End If
#End If
                        '[END]'
                        
                        '[Casco]'
                            If TempChar.Casco.Head(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Casco.Head(TempChar.Heading), _
                                        iPPx + TempChar.Body.HeadOffset.X, _
                                        iPPy + TempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            End If
                        '[END]'
                        
                        '[ARMA]'
                            If TempChar.Arma.WeaponWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                        '[END]'
                        
                        '[Escudo]'
                            If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                        '[END]'
                    End If
    
                    If Dialogos.CantidadDialogos > 0 Then
                        Call Dialogos.Update_Dialog_Pos( _
                                (iPPx + TempChar.Body.HeadOffset.X), _
                                (iPPy + TempChar.Body.HeadOffset.Y), _
                                MapData(X, Y).CharIndex)
                    End If
                    
    '                If Nombres Then
     '                   If TempChar.invisible = False Then
     '                       If TempChar.Nombre <> "" Then
     '                               Dim lCenter As Long
     '                               lCenter = Len(TempChar.Nombre) \ 2
     '                               If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
     '                                   Dim sClan As String: sClan = Mid(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
     '                                   If TempChar.Criminal Then
     '                                       Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(255, 0, 0))
     '                                       lCenter = Len(sClan) \ 2
     '                                       Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(255, 0, 0))
     '                                   Else
     '                                       Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(0, 128, 255))
     '                                       lCenter = Len(sClan) * 2
     '                                       Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(0, 128, 255))
     '                                   End If
     '                               Else
     '                                   If TempChar.Criminal Then
     '                                       Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(255, 0, 0))
     '                                   Else
     '                                       Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(0, 128, 255))
     '                                   End If
     '                               End If
     '                       End If
     '                   End If
     '               End If
     
                     If Nombres Then
                        If TempChar.invisible = False Then
                            If TempChar.Nombre <> "" Then
                                    Dim lCenter As Long
                                    If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                        lCenter = (frmMain.TextWidth(left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                        Dim sClan As String: sClan = Mid(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                
                                        Select Case TempChar.priv
                                        Case 0
                                            If TempChar.Criminal Then
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(255, 0, 0))
                                                lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(255, 0, 0))
                                            Else
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(0, 128, 255))
                                                lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(0, 128, 255))
                                            End If
                                        Case 1 'consejero
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(30, 150, 30))
                                            lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(30, 150, 30))
                                        Case 2 'semidios
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(30, 225, 30))
                                            lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(30, 255, 30))
                                        Case 3 ' dios
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(250, 250, 150))
                                            lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(250, 250, 150))
                                        Case 4 ' consejo de bander - consilio
                                            If TempChar.Criminal Then 'Concilio!
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(255, 50, 0))
                                                lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(255, 50, 0))
                                            Else 'Consejo:P
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(0, 195, 255))
                                                lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(0, 195, 255))
                                            End If
                                        End Select
                                    Else
                                        lCenter = (frmMain.TextWidth(TempChar.Nombre) / 2) - 16
                                        Select Case TempChar.priv
                                        Case 0
                                            If TempChar.Criminal Then
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(255, 0, 0))
                                            Else
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(0, 128, 255))
                                            End If
                                        Case 1
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(30, 150, 30))
                                        Case 2
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(30, 255, 30))
                                        Case 3
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(250, 250, 150))
                                        Case 4
                                            If TempChar.Criminal = 255 Then 'sOMBRAS
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(100, 100, 100))
                                            Else 'Banderbill
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(0, 195, 255))
                                            End If
                                        End Select
                                    End If
                            End If
                        End If
                     End If
                Else '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                    If Dialogos.CantidadDialogos > 0 Then
                        Call Dialogos.Update_Dialog_Pos( _
                                (iPPx + TempChar.Body.HeadOffset.X), _
                                (iPPy + TempChar.Body.HeadOffset.Y), _
                                MapData(X, Y).CharIndex)
                    End If
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            TempChar.Body.Walk(TempChar.Heading), _
                            iPPx, iPPy, 1, 1)
                End If '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                If X > 0 And Y > 0 Then
                    'Refresh charlist
                    CharList(MapData(X, Y).CharIndex) = TempChar
                    'BlitFX (TM)
                    If CharList(MapData(X, Y).CharIndex).Fx <> 0 Then
#If (ConAlfaB = 1) Then
                        If RenderMod.bNoAlpha Then
#End If
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    FxData(TempChar.Fx).Fx, _
                                    iPPx + FxData(TempChar.Fx).OffsetX, _
                                    iPPy + FxData(TempChar.Fx).OffsetY, _
                                    1, 1, MapData(X, Y).CharIndex)
#If (ConAlfaB = 1) Then
                        Else
                            Call DDrawTransGrhtoSurfaceAlpha( _
                                    BackBufferSurface, _
                                    FxData(TempChar.Fx).Fx, _
                                    iPPx + FxData(TempChar.Fx).OffsetX, _
                                    iPPy + FxData(TempChar.Fx).OffsetY, _
                                    1, 1, MapData(X, Y).CharIndex)
                        End If
#End If
                    End If
                End If
            End If '<-> If MapData(X, Y).CharIndex <> 0 Then
            
            '*************************************************
            'Layer 3 *****************************************
            If MapData(X, Y).Graphic(3).GrhIndex <> 0 Then
                'Draw
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(3), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
            End If
            '************************************************
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y
If Not bTecho Then
    'Draw blocked tiles and grid
    ScreenY = 5 + RenderMod.iImageSize
    For Y = (minY + 5) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
        ScreenX = 5 + RenderMod.iImageSize
        For X = (minX + 5) + RenderMod.iImageSize To (maxX - 0) - RenderMod.iImageSize
            'Check to see if in bounds
            If X < 101 And X > 0 And Y < 101 And Y > 0 Then
                If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(4), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
End If
If bLluvia(UserMap) = 1 Then
    If bRain Or bRainST Then
                'Figure out what frame to draw
                If llTick < DirectX.TickCount - 50 Then
                    iFrameIndex = iFrameIndex + 1
                    If iFrameIndex > 7 Then iFrameIndex = 0
                    llTick = DirectX.TickCount
                End If
                For Y = 0 To 4
                    For X = 0 To 4
                        Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.GetBMP(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
                    Next X
                Next Y
    End If
End If

'[Misery_Ezequiel 10/07/05]
If bNieva(UserMap) = 1 Then
    If bSnow Then
    'Figure out what frame to draw
        If llTick < DirectX.TickCount - 50 Then
            iFrameIndex = iFrameIndex + 1
        If iFrameIndex > 7 Then iFrameIndex = 0
            llTick = DirectX.TickCount
        End If
        For Y = 0 To 4
            For X = 0 To 4
                Call BackBufferSurface.BltFast(LTNieva(Y), LTNieva(X), SurfaceDB.GetBMP(5557), RNieva(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
            Next X
        Next Y
    End If
End If
'[\]Misery_Ezequiel 10/07/05]

Dim PP As RECT

PP.left = 0
PP.top = 0
PP.right = WindowTileWidth * TilePixelWidth
PP.bottom = WindowTileHeight * TilePixelHeight

'Call BackBufferSurface.BltFast(LTLluvia(0) + TilePixelWidth, LTLluvia(0) + TilePixelHeight, SurfaceDB.GetBMP(10000), PP, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
'EfectoNoche BackBufferSurface
'[USELESS]:El codigo para llamar a la noche, nublado, etc.
'            If bTecho Then
'                Dim bbarray() As Byte, nnarray() As Byte
'                Dim ddsdBB As DDSURFACEDESC2 'backbuffer
'                Dim ddsdNN As DDSURFACEDESC2 'nnublado
'                Dim r As RECT, r2 As RECT
'                Dim retVal As Long
'                '[LOCK]:BackBufferSurface
'                    BackBufferSurface.GetSurfaceDesc ddsdBB
'                    'BackBufferSurface.Lock r, ddsdBB, DDLOCK_NOSYSLOCK + DDLOCK_WRITEONLY + DDLOCK_WAIT, 0
'                    BackBufferSurface.Lock r, ddsdBB, DDLOCK_WRITEONLY + DDLOCK_WAIT, 0
'                    BackBufferSurface.GetLockedArray bbarray()
''                '[LOCK]:BBMask
''                    SurfaceXU(2).GetSurfaceDesc ddsdNN
''                    'SurfaceXU(2).Lock r2, ddsdNN, DDLOCK_READONLY + DDLOCK_NOSYSLOCK + DDLOCK_WAIT, 0
''                    SurfaceXU(2).Lock r2, ddsdNN, DDLOCK_READONLY + DDLOCK_WAIT, 0
''                    SurfaceXU(2).GetLockedArray nnarray()
'                '[BLIT]'
'                    'retVal = BlitNoche(bbarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth, 0)
'                    'retval = BlitNublar(bbarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth)
'                    'retVal = BlitNublarMMX(bbarray(0, 0), nnarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth, ddsdBB.lPitch, ddsdNN.lHeight, ddsdNN.lWidth, ddsdNN.lPitch)
'                '[UNLOCK]'
'                    BackBufferSurface.Unlock r
'                    'SurfaceXU(2).Unlock r2
'                '[END]'
'                If retVal = -1 Then MsgBox "error!"
'            End If
'[END]'

Exit Sub
errh: 'by Gorlok
    Debug.Print "[ERROR]RenderScreen: line " & EEE & " [" & Err.Number & "] " & Err.description
End Sub

Public Function RenderSounds()
'[CODE 001]:MatuX'
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> plLluviain Then
                    Call frmMain.StopSound
                    Call frmMain.Play("lluviain.wav", True)
                    frmMain.IsPlaying = plLluviain
                End If
                'Call StopSound("lluviaout.MP3")
                'Call PlaySound("lluviain.MP3", True)
            Else
                If frmMain.IsPlaying <> plLluviaout Then
                    Call frmMain.StopSound
                    Call frmMain.Play("lluviaout.wav", True)
                    frmMain.IsPlaying = plLluviaout
                End If
                'Call StopSound("lluviain.MP3")
                'Call PlaySound("lluviaout.MP3", True)
            End If
        End If
    End If
'[END]'
'[Misery_Ezequiel 10/07/05]
If bNieva(UserMap) = 1 Then
    If bSnow Then
        'If frmMain.IsPlaying <> plLluviaout Then
         '   Call frmMain.StopSound
          '  Call frmMain.Play("nieve.wav", True)
           ' frmMain.IsPlaying = plLluviaout
        'End If
    End If
End If
'[END]'
'[\]Misery_Ezequiel 10/07/05]
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
If GrhIndex > 0 Then
        HayUserAbajo = _
            CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And CharList(UserCharIndex).Pos.Y <= Y
End If
End Function

Function PixelPos(X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
PixelPos = (TilePixelWidth * X) - TilePixelWidth
End Function

Sub LoadGraphics()
        Dim loopc As Integer
        Dim SurfaceDesc As DDSURFACEDESC2
        Dim ddck As DDCOLORKEY
        Dim ddsd As DDSURFACEDESC2
        Dim iLoopUpdate As Integer

        SurfaceDB.TotalGraficos = Config_Inicio.NumeroDeBMPs + 1
        SurfaceDB.MaxEntries = 150
        SurfaceDB.lpDirectDraw7 = DirectDraw
        SurfaceDB.Path = DirGraficos
        Call SurfaceDB.Init(IIf(RenderMod.bUseVideo, True, False))
        If Not SurfaceDB.EsDinamico Then
            iLoopUpdate = 1
            For loopc = 1 To Config_Inicio.NumeroDeBMPs + 1
                SurfaceDB.CargarGrafico loopc
                If loopc > (iLoopUpdate + (Config_Inicio.NumeroDeBMPs / 80)) Then
                    AddtoRichTextBox frmCargando.Status, ".", , , , , , True
                    iLoopUpdate = loopc
                End If
            Next loopc
        End If
         If FileExist(App.Path & "\MIDI\" & "lluviatds" & ".mp3", vbNormal) Then
         Music_MP3_Load App.Path & "\MIDI\" & "lluviatds" & ".mp3"
            End If
        'Bmp de la lluvia
        
        Call GetBitmapDimensions(DirGraficos & "5556.bmp", ddsd.lWidth, ddsd.lHeight)
        RLluvia(0).top = 0:      RLluvia(1).top = 0:      RLluvia(2).top = 0:      RLluvia(3).top = 0
        RLluvia(0).left = 0:     RLluvia(1).left = 128:   RLluvia(2).left = 256:   RLluvia(3).left = 384
        RLluvia(0).right = 128:  RLluvia(1).right = 256:  RLluvia(2).right = 384:  RLluvia(3).right = 512
        RLluvia(0).bottom = 128: RLluvia(1).bottom = 128: RLluvia(2).bottom = 128: RLluvia(3).bottom = 128
        RLluvia(4).top = 128:    RLluvia(5).top = 128:    RLluvia(6).top = 128:    RLluvia(7).top = 128
        RLluvia(4).left = 0:     RLluvia(5).left = 128:   RLluvia(6).left = 256:   RLluvia(7).left = 384
        RLluvia(4).right = 128:  RLluvia(5).right = 256:  RLluvia(6).right = 384:  RLluvia(7).right = 512
        RLluvia(4).bottom = 256: RLluvia(5).bottom = 256: RLluvia(6).bottom = 256: RLluvia(7).bottom = 256
        AddtoRichTextBox frmCargando.Status, "Hecho.", , , , 1, , False
'[Misery_Ezequiel 10/07/05]
        'Bmp de la nieve
        Call GetBitmapDimensions(DirGraficos & "5557.bmp", ddsd.lWidth, ddsd.lHeight)
        RNieva(0).top = 0:      RNieva(1).top = 0:      RNieva(2).top = 0:      RNieva(3).top = 0
        RNieva(0).left = 0:     RNieva(1).left = 128:   RNieva(2).left = 256:   RNieva(3).left = 384
        RNieva(0).right = 128:  RNieva(1).right = 256:  RNieva(2).right = 384:  RNieva(3).right = 512
        RNieva(0).bottom = 128: RNieva(1).bottom = 128: RNieva(2).bottom = 128: RNieva(3).bottom = 128
        RNieva(4).top = 128:    RNieva(5).top = 128:    RNieva(6).top = 128:    RNieva(7).top = 128
        RNieva(4).left = 0:     RNieva(5).left = 128:   RNieva(6).left = 256:   RNieva(7).left = 384
        RNieva(4).right = 128:  RNieva(5).right = 256:  RNieva(6).right = 384:  RNieva(7).right = 512
        RNieva(4).bottom = 256: RNieva(5).bottom = 256: RNieva(6).bottom = 256: RNieva(7).bottom = 256
        AddtoRichTextBox frmCargando.Status, "Hecho.", , , , 1, , False
'[\]Misery_Ezequiel 10/07/05]
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************

Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"
'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder
'Fill startup variables
DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize
MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)

ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

If Musica = 0 Or Fx = 0 Then
    DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
End If

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With

Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .left = 0 + 32 * RenderMod.iImageSize
    .top = 0 + 32 * RenderMod.iImageSize
    .right = (TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
    .bottom = (TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
End With
With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If RenderMod.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.bottom
    .lWidth = BackBufferRect.right
End With
Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)
ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs

LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

'[Misery_Ezequiel 10/07/05]
LTNieva(0) = 224
LTNieva(1) = 352
LTNieva(2) = 480
LTNieva(3) = 608
LTNieva(4) = 736
'[\]Misery_Ezequiel 10/07/05]

 frmCargando.Label1 = "Cargando Gráficos...."
Call LoadGraphics
InitTileEngine = True
End Function

'Sub ShowNextFrame(DisplayFormTop As Integer, DisplayFormLeft As Integer)
Sub ShowNextFrame()

'[CODE]:MatuX'
'
'  ESTA FUNCIÓN FUE MOVIDA AL LOOP PRINCIPAL EN Mod_General
'  PARA QUE SEA INLINE. EN OTRAS PALABRAS, LO QUE ESTÁ ACÁ
'  YA NO ES LLAMADO POR NINGUNA RUTINA.
'
'[END]'
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer
    If EngineRun Then
        '  '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            'End If
            '****** Move screen Up and Down if needed ******
            'If AddtoUserPos.Y <> 0 Then
            ElseIf AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If
            '****** Update screen ******
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            'Call DoNightFX
            'Call DoLightFogata(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
            Call MostrarFlags
            Call Dialogos.MostrarTexto
            Call DibujarCartel
            Call DrawBackBufferSurface
            'Call DibujarInv(frmMain.picInv.hWnd, 0)
            FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

'[CODE 000]:MatuX
' La hice inline
Sub MostrarFlags()
If IScombate Then
    Call Dialogos.DrawText(260, 260, "MODO COMBATE", vbRed)
End If
'[END]'
End Sub

Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1) 'Anim Invent
Call CrearGrh(534, 2) 'Animacion de teleport
End Sub

Function ControlVelocidad(ByVal LastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - LastTime > 20)
End Function

Public Sub EfectoNoche(ByRef surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT
Dim DstLock As Boolean

surface.GetSurfaceDesc ddsdDest
With rRect
.left = 0
.top = 0
.right = ddsdDest.lWidth
.bottom = ddsdDest.lHeight
End With
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
Else
    Modo = 2
End If
DstLock = False

On Local Error GoTo HayErrorAlpha

surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True
surface.GetLockedArray dArray()
Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
    ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
    Modo)
HayErrorAlpha:
If DstLock = True Then
    surface.Unlock rRect
    DstLock = False
End If
End Sub



