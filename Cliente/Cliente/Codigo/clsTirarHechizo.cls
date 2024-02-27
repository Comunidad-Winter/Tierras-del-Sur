VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTirarHechizo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Integer
Dim BolaX As Single
Dim BolaY As Single
Dim StrX As Single
Dim StrY As Single
Dim Angulo As Single
Dim TipoBola As Byte
Public Activo As Boolean
Public CharIndex As Integer
Function DDRect(x As Integer, y As Integer, x2 As Long, y2 As Long) As RECT
DDRect.Left = x
DDRect.Top = y
DDRect.Bottom = y2
DDRect.Right = x2
End Function
Public Sub Render(EndX As Integer, EndY As Integer, tilex As Integer, tiley As Integer, PixelOffsetX As Single, PixelOffsetY As Single)
'On Error Resume Next
'Dim DistX As Single, DistY As Single, Vel As Single, AuX As Single, AuY As Single, Frames As Integer
'
'Dim OffSX As Single, OffSY As Single
'DistX = Abs(BolaX - EndX * 32 - 8)
'DistY = Abs(BolaY - EndY * 32 - 8)
'Vel = 20 * VelMove / 8
'Frames = Sqr(DistX ^ 2 + DistY ^ 2) / Vel
'If DistX > 0 And Frames > 0 Then AuX = DistX / Frames
'If DistY > 0 And Frames > 0 Then AuY = DistY / Frames
'If BolaX > EndX * 32 + 8 Then AuX = -AuX
'If BolaY > EndY * 32 + 8 Then AuY = -AuY
'BolaX = BolaX + AuX
'BolaY = BolaY + AuY
'If DistX = 0 And DistY = 0 Then
'    Activo = False
'End If
'OffSX = 32 * Sgn(PixelOffsetX)
'OffSY = 32 * Sgn(PixelOffsetY)
''PuntoGiro.Y = -8
'Dim RectBola As RECT
'Select Case TipoBola
'    Case 0
'        Angulo = 0
'        RectBola = DDRect(0, 0, 16, 16)
'    Case 1
'        Angulo = Angulo + 0.5
'        RectBola = DDRect(16 * Int(Angulo / 10), 0, 16 * Int(Angulo / 10) + 16, 16)
'    Case 2
'        Angulo = DameAngulo(StrX, StrY, EndX * 32, EndY * 32)
'
'        RectBola = DDRect(16 * Int(Angulo / 10), 0, 16 * Int(Angulo / 10) + 16, 16)
'End Select
'
'BackBufferSurface.BltFast BolaX + PixelOffsetX - OffSX - (tilex - 8) * 32 + 256, BolaY + PixelOffsetY - OffSY - (tiley - 6) * 32 + 256, SurfaceBola(TipoBola + 1), RectBola, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub
Function DameAngulo(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
On Error Resume Next
Dim Esquina As Integer
Dim Angulo As Single
If x1 > x2 And y1 < y2 Then
    Esquina = 0
ElseIf x1 <= x2 And y1 < y2 Then
    Esquina = 1
ElseIf x1 <= x2 And y1 >= y2 Then
    Esquina = 2
ElseIf x1 >= x2 And y1 > y2 Then
    Esquina = 3
End If
If (y1 - y2) = (x1 - x2) Then
    Dist = 0
Else
    If Esquina = 0 Or Esquina = 2 Then
        Dist = (y1 - y2) / (x1 - x2)
    Else
        Dist = (x1 - x2) / (y1 - y2)
    End If
End If
Dist = Sqr(Dist ^ 2)
Angulo = Atn(Dist) * 57.2957795130824
Angulo = Angulo + Esquina * 90 ' / 57.2957795130824
If Esquina = 2 And x1 = x2 Then Angulo = 270
If Sgn(Angulo) = -1 Then Angulo = 360 + Angulo
DameAngulo = Angulo
End Function
Public Sub Init(tX As Integer, tY As Integer, CIndex As Integer, tipo As Byte)
x = tX
y = tY
CharIndex = CIndex
TipoBola = tipo
BolaX = x * 32
BolaY = y * 32
StrX = BolaX
StrY = BolaY
Activo = True
End Sub


