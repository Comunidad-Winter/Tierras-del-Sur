Attribute VB_Name = "Engine_Parallax"
#If EnableParallax = 1 Then
Option Explicit

Public ParallaxOffsets(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)   As D3DVECTOR2

Public Sub GetParalllaxOffset(ByVal X!, ByVal Y!, ByVal z!, ByRef OX!, ByRef OY!)
'    OX = ((X + D3DWindow.BackBufferWidth / 2) / D3DWindow.BackBufferWidth) * z
'    OY = ((Y + D3DWindow.BackBufferWidth / 2) / D3DWindow.BackBufferWidth) * z
    OX = ((X - 512) / 1024) * (z - Screen_Desnivel_Offset) * 1
    OY = ((Y - 512) / 1024) * (z - Screen_Desnivel_Offset) * 1
End Sub

Public Sub ActualizarParallax()

End Sub

#End If
