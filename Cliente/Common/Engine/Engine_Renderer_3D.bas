Attribute VB_Name = "Engine_Renderer_3D"
Option Explicit


Public Sub Grh_Render_Vertical(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, ByVal PixelOffsetX As Long, ByVal PixelOffsetY As Long, ByVal mirror As Long)
    '*********************************************
    'Author: menduz
    '*********************************************

    If GrhIndex = 0 Then Exit Sub
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)

    
    Dim TGRH As GrhData
    TGRH = GrhData(GrhIndex)
    
    If TGRH.hardcor = 0 Then Init_grh_tutv GrhIndex
    
    
    Colorear_TBOX tBox, map_x, map_y
    
    With tBox
        .x0 = dest_x
        .y0 = dest_y + TGRH.pixelHeight
        .x1 = .x0
        .y1 = dest_y
        .x2 = dest_x + TGRH.pixelWidth
        .y2 = .y0
        .x3 = .x2
        .y3 = .y1

        .Z0 = (map_y / 218 * 32 + PixelOffsetY) / D3DWindow.BackBufferHeight
        .Z1 = .Z0
        .z2 = .Z0
        .Z3 = .Z0

        'If alpha Then
        '    .color0 = (.color0 And &HFFFFFF) Or Alphas(alpha)
        '    .Color1 = (.Color1 And &HFFFFFF) Or Alphas(alpha)
        '    .Color2 = (.Color2 And &HFFFFFF) Or Alphas(alpha)
        '    .color3 = (.color3 And &HFFFFFF) Or Alphas(alpha)
        'End If
        
        If Not mirror Then
            .tu0 = TGRH.tu(0)
            .tv0 = TGRH.tv(0)
            .tv1 = TGRH.tv(1)
            .tu2 = TGRH.tu(2)
            .tu1 = .tu0
            .tv2 = .tv0
            .tu3 = .tu2
            .tv3 = .tv1
        Else
            .tu0 = TGRH.tu(2)
            .tv0 = TGRH.tv(2)
            .tu1 = TGRH.tu(3)
            .tv1 = TGRH.tv(3)
            .tu2 = TGRH.tu(0)
            .tv2 = TGRH.tv(0)
            .tu3 = TGRH.tu(1)
            .tv3 = TGRH.tv(1)
        End If
        

        .tu01 = .tu0
        .tu11 = .tu1
        .tu21 = .tu2
        .tu31 = .tu3
        .tv01 = .tv0
        .tv11 = .tv1
        .tv21 = .tv2
        .tv31 = .tv3
    End With
    
    
    Dim C3%
    
    Call Obtener_Texturas_Complementarias(CInt(GrhData(GrhIndex).filenum), 0, 0, C3)

    If C3 = 0 Then
        C3 = Engine_LightsTexture.LightTextureWall
    End If

    CalcularNormalColorTBOX_TexturaVertical tBox
    Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureVertical

    Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)

    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Estandar
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    
    Grh_Render_Complementario TGRH.filenum
End Sub

Public Sub Grh_Render_Tileset(ByVal tileset As Integer, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte, ByRef altU As AUDT)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!

    Dim tn As Byte
    
    Dim tex As Integer
    
    tn = mapdata(map_x, map_y).tile_number

    If tileset = 0 Then Exit Sub
    
    tex = Tilesets(tileset).filenum
    
    
    Dim C3%, C1%, C2%
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    
    Call Engine_TextureDB.Obtener_Texturas_Complementarias(tex, C1, C2, C3)
    
    If C3 = 0 Then C3 = Engine_LightsTexture.LightTextureFloor
    
    Dim TieneC1oC2 As Long
    
    TieneC1oC2 = C1 Or C2

    If (Cachear_Tiles Or mapdata(map_x, map_y).tile_render <> 255) Then
        tBottom = tTop + 32
        tRight = tLeft + 32

        Colorear_TBOX Tileset_Grh_Array(tn), map_x, map_y
        
        D3DDevice.SetRenderState D3DRS_ZENABLE, 1
        D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
        D3DDevice.SetRenderState D3DRS_ZFUNC, D3DCMP_GREATEREQUAL
        
        With Tileset_Grh_Array(tn)
            .x0 = tLeft
            .y0 = tBottom - altU.hs(0)
            .x1 = tLeft
            .y1 = tTop - altU.hs(1)
            .x2 = tRight
            .y2 = tBottom - altU.hs(2)
            .x3 = tRight
            .y3 = tTop - altU.hs(3)
            
            .Z0 = (map_y / 218 * 32) / D3DWindow.BackBufferHeight
            .Z1 = (map_y / 218 * 32) / D3DWindow.BackBufferHeight
            .z2 = .Z0
            .Z3 = .Z1
            
            ' Mapeo las coordenadas del normal map
            .tu01 = .tu0
            .tu11 = .tu1
            .tu21 = .tu2
            .tu31 = .tu3
            .tv01 = .tv0
            .tv11 = .tv1
            .tv21 = .tv2
            .tv31 = .tv3
            
            ' Mapeo la tercera coordenada de luces contra las luces
            .tu02 = .x0 / D3DWindow.BackBufferWidth
            .tu12 = .tu02
            .tu22 = .x2 / D3DWindow.BackBufferWidth
            .tu32 = .tu22
            
            .tv02 = tBottom / D3DWindow.BackBufferHeight
            .tv12 = tTop / D3DWindow.BackBufferHeight
            .tv22 = .tv02
            .tv32 = .tv12

            Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)
            Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureHorizontal
            
            Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Estandar
            
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, .x0, TL_size
        End With

        With mapdata(map_x, map_y)
            If Cachear_Tiles Or .tile_render = 0 Then
                DXCopyMemory MapBoxes(map_x, map_y), Tileset_Grh_Array(tn), BV_size
                'MapBoxes(map_x, map_y) = Tileset_Grh_Array(tn)
                If .tile_render = 0 Then
                    .tile_render = 255
                End If
            End If
        End With
        
        If TieneC1oC2 Then Grh_Render_Complementario_Tileset tex, tn
    Else
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Estandar
        Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)
        Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureHorizontal
        
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MapBoxes(map_x, map_y), TL_size

        If TieneC1oC2 Then Grh_Render_Complementario_Tileset_Cacheado tex, map_x, map_y
    End If

End Sub
