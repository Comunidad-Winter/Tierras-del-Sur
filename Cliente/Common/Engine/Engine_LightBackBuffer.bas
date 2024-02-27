Attribute VB_Name = "Engine_LightBackBuffer"
Option Explicit

Public LightBackBufferWidth As Single
Public LightBackBufferHeight As Single

Public LightBackBufferOffsetX As Single
Public LightBackBufferOffsetY As Single

Public Sub LightBackBuffer_Init()
    LightBackBufferWidth = D3DWindow.BackBufferWidth
    LightBackBufferHeight = D3DWindow.BackBufferHeight
    
    LightBackBufferOffsetX = 0
    LightBackBufferOffsetY = 0
End Sub

Public Sub LightBackBuffer_ApplyCoordinates(Box As Box_Vertex)
    With Box
        .tu02 = (.x0 + LightBackBufferOffsetX) / LightBackBufferWidth
        .tu12 = (.x1 + LightBackBufferOffsetX) / LightBackBufferWidth
        .tu22 = (.x2 + LightBackBufferOffsetX) / LightBackBufferWidth
        .tu32 = (.x3 + LightBackBufferOffsetX) / LightBackBufferWidth
        
        .tv02 = (.y0 + LightBackBufferOffsetY) / LightBackBufferHeight
        .tv12 = (.y1 + LightBackBufferOffsetY) / LightBackBufferHeight
        .tv22 = (.y2 + LightBackBufferOffsetY) / LightBackBufferHeight
        .tv32 = (.y3 + LightBackBufferOffsetY) / LightBackBufferHeight
    End With
End Sub

