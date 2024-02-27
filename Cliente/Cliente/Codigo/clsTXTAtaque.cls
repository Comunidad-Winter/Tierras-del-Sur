VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTXTAtaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Integer
Dim texto As String
Dim OffsetX As Single
Dim OffsetY As Single
Dim IntervalW As Integer
Dim red As Single
Dim Azul As Boolean
Public Activo As Boolean
Public Apuñalar As Boolean
Public Sub Render(tilex As Integer, tiley As Integer, PixelOffsetX As Single, PixelOffsetY As Single)
'
'If OffsetY <= -25 Then
'    Activo = False
'    Exit Sub
'End If
'If IntervalW > 0 Then
'    IntervalW = IntervalW - 1
'    Exit Sub
'End If
'
'OffsetY = OffsetY - 1.5 * VelMove / 8
'red = red - 10 * VelMove / 8
'If red < 0 Then red = 0
'OffsetX = 16 - Len(texto) * 4
'If Apuñalar Then
'BackBufferSurface.SetForeColor &HC0FFFF
'Else
'BackBufferSurface.SetForeColor RGB(IIf(Azul, 0, red), IIf(Azul, red, 0), 0)
'End If
''frmMain.Caption = 32 * (X + 8 - tilex) + OffSetX + PixelOffSetX
'BackBufferSurface.DrawText 32 * (x + 8 - tilex) + OffsetX + PixelOffsetX + 256, 32 * (y + 6 - tiley) + OffsetY + PixelOffsetY + IIf(Azul, 15, 0) + 256, texto, False
End Sub
Public Sub Init(SX As Integer, SY As Integer, txt As String, TiempoWait As Integer)

x = SX
y = SY
Azul = (Sgn(val(txt)) = -1)
texto = Abs(val(txt))
IntervalW = TiempoWait * 8 / VelMove
OffsetX = 0
OffsetY = 0
red = 255
Activo = True
End Sub

