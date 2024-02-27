Attribute VB_Name = "Engine_Map_Render"
Option Explicit

Public MapBoxes_Geometry(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As Box_Vertex

Private MapBoxes_VertexBuffer As Direct3DVertexBuffer8

Public Sub MapBox_Init()
    Dim x As Long
    Dim y As Long
    
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
            With MapBoxes_Geometry(x, y)
                .x0 = x * 32
                .y1 = y * 32
                .y0 = .y1 + 32
                .x1 = .x0
                
                .x2 = .x0 + 32
                .y2 = .y1 + 32
                .x3 = .x0 + 32
                .y3 = .y1
                
                .rhw0 = 1
                .rhw1 = 1
                .rhw2 = 1
                .rhw3 = 1
            End With
        Next x
    Next y
End Sub

Public Sub MapBox_Draw()
    Dim x As Long
    Dim y As Long
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Pisos

    Dim pMaxX As Long
    Dim pMaxY As Long
    
    pMaxX = minl(screenmaxX + 1, MaxX)
    pMaxY = minl(screenmaxY + 1, MaxY)

    If Not Cachear_Tiles Then
        
        y = screenminY
    
        While y < pMaxY
            x = screenminX
            While x < pMaxX
                MapBox_DrawTileCached x, y
                x = x + 1
            Wend
            y = y + 1
        Wend
    Else
        x = screenminX
        y = screenminY
    
        While y < pMaxY
            x = screenminX
            While x < pMaxX
                MapBox_DrawTile x, y, hMapData(x, y)
                x = x + 1
            Wend
            y = y + 1
        Wend

        Cachear_Tiles = False
    End If
End Sub

Private Sub MapBox_DrawTile(ByVal map_x As Byte, ByVal map_y As Byte, ByRef altU As AUDT)
    '*********************************************
    'Author: menduz
    '*********************************************

    Dim C3%, C1%, C2%
    Dim tn As Byte
    
    Dim tex As Integer
    Dim tileset As Long
    
    tileset = mapdata(map_x, map_y).tile_texture

    If tileset = 0 Then Exit Sub
    
    tex = Tilesets(tileset).filenum

    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    
    Call Engine_TextureDB.Obtener_Texturas_Complementarias(tex, C1, C2, C3)
    
    tn = mapdata(map_x, map_y).tile_number

    'altU = hMapData(map_x, map_y)

    Colorear_TBOX MapBoxes_Geometry(map_x, map_y), map_x, map_y
          
    With MapBoxes_Geometry(map_x, map_y)
        .Z0 = altU.hs(0)
        .Z1 = altU.hs(1)
        .z2 = altU.hs(2)
        .Z3 = altU.hs(3)
        
        .tu0 = Tileset_Grh_Array(tn).tu0
        .tv0 = Tileset_Grh_Array(tn).tv0
        .tu1 = Tileset_Grh_Array(tn).tu1
        .tv1 = Tileset_Grh_Array(tn).tv1
        .tu2 = Tileset_Grh_Array(tn).tu2
        .tv2 = Tileset_Grh_Array(tn).tv2
        .tu3 = Tileset_Grh_Array(tn).tu3
        .tv3 = Tileset_Grh_Array(tn).tv3
        
        ' Mapeo la tercera coordenada de luces contra las luces
        .tu02 = .x0 / D3DWindow.BackBufferWidth
        .tu12 = .tu02
        .tu22 = .x2 / D3DWindow.BackBufferWidth
        .tu32 = .tu22
        
        .tv02 = .y2 / D3DWindow.BackBufferHeight
        .tv12 = .y1 / D3DWindow.BackBufferHeight
        .tv22 = .tv02
        .tv32 = .tv12
    End With
    
    If C3 = 0 Then C3 = LightTextureFloor
    
    Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)
    Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureHorizontal
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MapBoxes_Geometry(map_x, map_y), TL_size
End Sub


Private Sub MapBox_DrawTileCached(ByVal map_x As Long, ByVal map_y As Long)
    '*********************************************
    'Author: menduz
    '*********************************************

    Dim C3%, C1%, C2%
    Dim tn As Byte
    
    Dim tex As Integer
    Dim tileset As Long
    
    tileset = mapdata(map_x, map_y).tile_texture

    If tileset = 0 Then Exit Sub
    
    tex = Tilesets(tileset).filenum

    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    
    Call Engine_TextureDB.Obtener_Texturas_Complementarias(tex, C1, C2, C3)
    
    If C3 = 0 Then C3 = LightTextureFloor
    
    Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)
    Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureHorizontal
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MapBoxes_Geometry(map_x, map_y), TL_size
End Sub

