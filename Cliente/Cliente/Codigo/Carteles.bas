Attribute VB_Name = "CLI_Carteles"
Option Explicit
Const XPosCartel As Integer = 100
Const YPosCartel As Integer = 100
Const MAXLONG As Integer = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Private LeyendaFormateada() As String
Private textura As Integer

Sub InitCartel(Ley As String, Grh As Integer)
    Dim i As Integer, k As Integer, anti As Integer
    If Not Cartel Then
        Leyenda = Ley
        textura = Grh
        Cartel = True
        ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
        anti = 1
        k = 0
        i = 0
        Call DarFormato(Leyenda, i, k, anti)
        i = 0
        Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
           i = i + 1
        Loop
        ReDim Preserve LeyendaFormateada(0 To i)
    Else
        Exit Sub
    End If
End Sub

Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
    If anti + i <= Len(s) + 1 Then
        If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
            LeyendaFormateada(k) = mid(s, anti, i + 1)
            k = k + 1
            anti = anti + i + 1
            i = 0
        Else
            i = i + 1
        End If
        Call DarFormato(s, i, k, anti)
    End If
End Function

Sub DibujarCartel()
    Dim j As Integer, desp As Integer
    If Not Cartel Then Exit Sub
    Dim x As Single, y As Single
    x = XPosCartel + 20
    y = YPosCartel + 60
    Engine_GrhDraw.Grh_Render_nocolor textura, XPosCartel, YPosCartel
    For j = 0 To UBound(LeyendaFormateada)
        Engine.text_render_graphic LeyendaFormateada(j), x, y + desp, mzWhite
        desp = desp + 16
    Next
End Sub
