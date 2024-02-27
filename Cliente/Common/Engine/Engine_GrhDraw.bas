Attribute VB_Name = "Engine_GrhDraw"
Option Explicit

Public Sub Grh_Render(ByVal GrhIndex As Long, ByVal dest_x%, ByVal dest_y%, ByVal Color As Long)
'*********************************************
'Author: menduz
'*********************************************
    Dim TGRH As GrhData
    If GrhIndex = 0 Then Exit Sub
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)

    If GrhData(GrhIndex).hardcor = 0 Then Init_grh_tutv GrhIndex
    TGRH = GrhData(GrhIndex)

    With tBox
        .x0 = dest_x
        .y0 = dest_y + TGRH.pixelHeight
        .x1 = .x0
        .y1 = dest_y
        .x2 = dest_x + TGRH.pixelWidth
        .y2 = .y0
        .x3 = .x2
        .y3 = .y1
        .color0 = Color
        .Color1 = Color
        .Color2 = Color
        .color3 = Color
        .tu0 = TGRH.tu(0)
        .tv0 = TGRH.tv(0)
        .tu1 = TGRH.tu(1)
        .tv1 = TGRH.tv(1)
        .tu2 = TGRH.tu(2)
        .tv2 = TGRH.tv(2)
        .tu3 = TGRH.tu(3)
        .tv3 = TGRH.tv(3)
    End With
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
End Sub

Public Sub Grh_Render_Rotated(ByVal GrhIndex As Long, ByVal dest_x%, ByVal dest_y%, ByVal Color As Long, ByVal Angulo As Integer)
'http://stackoverflow.com/questions/3451061/how-to-do-correct-polygon-rotation-in-c-sharp-though-it-applies-to-anything


'*********************************************
'Author: menduz
'*********************************************
    Dim half_x!, half_y!, s!, c!
    Dim TGRH As GrhData
    
    If GrhIndex = 0 Then Exit Sub
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)

    If GrhData(GrhIndex).hardcor = 0 Then Init_grh_tutv GrhIndex

    TGRH = GrhData(GrhIndex)
        
    Angulo = (Angulo + 360) Mod 360
    
    With tBox
        If Angulo Then
            half_y = TGRH.pixelHeight / 2
            half_x = TGRH.pixelWidth / 2
            
            dest_x = dest_x + half_x
            dest_y = dest_y + half_y
            
            s = Seno(Angulo)
            c = Coseno(Angulo)

            .x0 = -half_x * c - half_y * s + dest_x
            .y0 = -half_x * s + half_y * c + dest_y

            .x1 = -half_x * c + half_y * s + dest_x
            .y1 = -half_x * s - half_y * c + dest_y

            .x2 = half_x * c - half_y * s + dest_x
            .y2 = half_x * s + half_y * c + dest_y

            .x3 = half_x * c + half_y * s + dest_x
            .y3 = half_x * s - half_y * c + dest_y
        Else
            .x0 = dest_x
            .y0 = dest_y + TGRH.pixelHeight
            .x1 = .x0
            .y1 = dest_y
            .x2 = dest_x + TGRH.pixelWidth
            .y2 = .y0
            .x3 = .x2
            .y3 = .y1
        End If

        .color0 = Color
        .Color1 = Color
        .Color2 = Color
        .color3 = Color
        .tu0 = TGRH.tu(0)
        .tv0 = TGRH.tv(0)
        .tu1 = TGRH.tu(1)
        .tv1 = TGRH.tv(1)
        .tu2 = TGRH.tu(2)
        .tv2 = TGRH.tv(2)
        .tu3 = TGRH.tu(3)
        .tv3 = TGRH.tv(3)
    End With
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
End Sub

Public Sub Grh_Render_invselslot(ByVal X!, ByVal Y!, Optional ByVal Color As Long = &HFFFFFFFF)
'*********************************************
'Author: menduz
'*********************************************
    Static Box As Box_Vertex
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing

    With Box
        .x0 = X
        .y0 = Y + 32
        .color0 = Color
        .x1 = .x0
        .y1 = Y
        .Color1 = Color
        .x2 = X + 32
        .y2 = .y0
        .Color2 = Color
        .x3 = .x2
        .y3 = .y1
        .color3 = Color
        .tu0 = 0
        .tv0 = 1
        .tu1 = 0
        .tv1 = 0
        .tu2 = 1
        .tv2 = 1
        .tu3 = 1
        .tv3 = 0
    End With
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Box, TL_size
End Sub


Public Sub Grh_Render_Tileset(ByVal tileset As Integer, ByVal map_x As Long, ByVal map_y As Long, ByRef altU As AUDT, ByRef MapBox As Box_Vertex)
'*********************************************
'Author: menduz
'*********************************************
    Dim tn As Byte
    Dim TieneC1oC2 As Long
    Dim tex As Integer
    Dim C3%, C1%, C2%
    
    tn = mapdata(map_x, map_y).tile_number

    If tileset = 0 Then Exit Sub
    
    tex = Tilesets(tileset).filenum

    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    
    Call Engine_TextureDB.Obtener_Texturas_Complementarias(tex, C1, C2, C3)
    
    If C3 = 0 Then C3 = LightTextureFloor
    
    TieneC1oC2 = C1 Or C2
    
    If Cachear_Tiles Then
        Colorear_TBOX MapBox, map_x, map_y

        With Tileset_Grh_Array(tn)
            MapBox.tu0 = .tu0
            MapBox.tv0 = .tv0
            MapBox.tu1 = .tu1
            MapBox.tv1 = .tv1
            MapBox.tu2 = .tu2
            MapBox.tv2 = .tv2
            MapBox.tu3 = .tu3
            MapBox.tv3 = .tv3
        End With
        
        MapBox.Z0 = altU.hs(0)
        MapBox.Z1 = altU.hs(1)
        MapBox.z2 = altU.hs(2)
        MapBox.Z3 = altU.hs(3)
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Pisos
    Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)
    Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureHorizontal
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MapBox, TL_size

    If TieneC1oC2 Then Grh_Render_Complementario_Tileset_Cacheado tex, map_x, map_y

End Sub

Public Sub Grh_Render_Complementario_Tileset_Cacheado(ByVal tex As Integer, ByVal MapX As Byte, ByVal MapY As Byte)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
Dim C1%, C2%

If Obtener_Texturas_Complementarias(tex, C1, C2) = True Then
    tBox = MapBoxes(MapX, MapY)
    If C1 > 0 Then
        With tBox
            .color0 = mzWhite
            .Color1 = mzWhite
            .Color2 = mzWhite
            .color3 = mzWhite
        End With
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(C1)
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    End If
    If C2 > 0 Then
        With tBox
            .color0 = base_light
            .Color1 = base_light
            .Color2 = base_light
            .color3 = base_light
        End With
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(C2)
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    End If
End If
End Sub

Public Sub Grh_Render_Complementario_Tileset(ByVal tex As Integer, ByVal tn As Integer)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
Dim C1%, C2%

If Obtener_Texturas_Complementarias(tex, C1, C2) = True Then
    If C1 > 0 Then
        With Tileset_Grh_Array(tn)
            .color0 = mzWhite
            .Color1 = mzWhite
            .Color2 = mzWhite
            .color3 = mzWhite
        End With
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(C1)
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Tileset_Grh_Array(tn), TL_size
    End If
    If C2 > 0 Then
        With Tileset_Grh_Array(tn)
            .color0 = base_light
            .Color1 = base_light
            .Color2 = base_light
            .color3 = base_light
        End With
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(C2)
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Tileset_Grh_Array(tn), TL_size
    End If
End If
End Sub
Public Sub Grh_Render_Objeto(ByVal tex As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal Color As Long, Optional ByVal alpha As Byte, Optional ByVal dw As Single)
'*********************************************
'Author: menduz
'viva la harcodeada, VIVA!
'*********************************************
    
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    Dim w!, h!
    GetTextureDimension tex, h, w
    If tLeft = -1 Then
        tLeft = 512
        tTop = 256
    End If
    tLeft = tLeft - w / 2
    tTop = tTop - h / 2
    With tBox
        .x0 = tLeft
        .y0 = tTop + h
        .color0 = Color
        .x1 = tLeft
        .y1 = tTop
        .Color1 = Color
        .x2 = tLeft + w + dw
        .y2 = tTop + h
        .Color2 = Color
        .x3 = tLeft + w + dw
        .y3 = tTop
        .color3 = Color
        .tu0 = 0
        .tv0 = 1
        .tu1 = 0
        .tv1 = 0
        .tu2 = 1
        .tv2 = 1
        .tu3 = 1
        .tv3 = 0
    End With
    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    Grh_Render_Complementario tex
    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub


Public Sub Grh_Render_Relieve_Tileset(ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte, ByVal flip As Byte)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!
    
    Dim ll As Long

    Dim altU As AUDT
    
    Dim h!, w!, sx!, sy!, texture As Integer, Number As Integer

    Number = mapdata(map_x, map_y).tile_number
    
    If Number = 0 Then Exit Sub
    
    altU = hMapData(map_x, map_y)
    texture = Tilesets(mapdata(map_x, map_y).tile_texture).filenum
    
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(texture)
    Call GetTextureDimension(texture, h, w)
    
    sx = (Number Mod (w \ 32)) * 32!
    sy = (Number \ (h \ 32)) * 32!
    
    Colorear_TBOX tBox, map_x, map_y
    
    tBottom = tTop + 32!
    tRight = tLeft + 32!
    
    With tBox
        .x0 = tLeft
        .y0 = tBottom - altU.hs(0)
        
        .x1 = tLeft
        .y1 = tTop - altU.hs(1)
        
        .x2 = tRight
        .y2 = tBottom - altU.hs(2)

        .x3 = tRight
        .y3 = tTop - altU.hs(3)
        
        If h Then
            .tu0 = sx / w
            .tv0 = (sy + 32) / h
            .tu1 = .tu0
            .tv1 = sy / h
            .tu2 = (sx + 32) / w
            .tv2 = .tv0
            .tu3 = .tu2
            .tv3 = .tv1
        End If
    End With
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size

End Sub

Public Sub Grh_Render_Relieve_Tileset_HC(ByVal tile_texture As Integer, ByVal tile_number As Integer, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte, ByVal flip As Byte)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!
    Dim altU As AUDT
    Dim h!, w!, sx!, sy! ', texture As Integer


    altU = hMapData(map_x, map_y)
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tile_texture)
    Call GetTextureDimension(tile_texture, h, w)
    
    If w = 0 Then Exit Sub
    
    sx = (tile_number Mod (w \ 32)) * 32!
    sy = (tile_number \ (h \ 32)) * 32!
    
    Colorear_TBOX tBox, map_x, map_y
    
    tBottom = tTop + 32!
    tRight = tLeft + 32!
    
    With tBox
        .x0 = tLeft
        .y0 = tBottom - altU.hs(0)
        
        .x1 = tLeft
        .y1 = tTop - altU.hs(1)
        
        .x2 = tRight
        .y2 = tBottom - altU.hs(2)

        .x3 = tRight
        .y3 = tTop - altU.hs(3)
        
        If h Then
            .tu0 = sx / w
            .tv0 = (sy + 32) / h
            .tu1 = .tu0
            .tv1 = sy / h
            .tu2 = (sx + 32) / w
            .tv2 = .tv0
            .tu3 = .tu2
            .tv3 = .tv1
        End If
    End With
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size

End Sub

Public Sub Grh_Render_Relieve_Tileset_HCD(ByVal tex As Integer, ByVal tn As Byte, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte, Color As Long)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!
    
    Dim altU As AUDT
    

    altU = hMapData(map_x, map_y)
    
    tBottom = tTop + 32
    tRight = tLeft + 32
    
    With Tileset_Grh_Array(tn)
        .x0 = tLeft
        .y0 = tBottom - altU.hs(0)
        .x1 = tLeft
        .y1 = tTop - altU.hs(1)
        .x2 = tRight
        .y2 = tBottom - altU.hs(2)
        .x3 = tRight
        .y3 = tTop - altU.hs(3)
        .color0 = Color
        .Color1 = .color0
        .Color2 = .color0
        .color3 = .color0
    End With

    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Tileset_Grh_Array(tn), TL_size
End Sub


Public Sub Grh_Render_Solid(ByVal Color As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal tWidth As Single, ByVal tHeight As Single)
'*********************************************
'Author: menduz
'*********************************************
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
    With tBox 'With tBox
            .x0 = tLeft
            .y0 = tTop + tHeight
            .color0 = Color
            .x1 = tLeft
            .y1 = tTop
            .Color1 = Color
            .x2 = tLeft + tWidth
            .y2 = .y0
            .Color2 = Color
            .x3 = .x2
            .y3 = tTop
            .color3 = Color
    End With
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
End Sub

Public Sub Grh_Render_Blocked(ByVal Color As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!
    
    
    Dim altU As AUDT
    'If GrhIndex = 0 Then Exit Sub
    
    If Not InMapBounds(map_x, map_y) Then Exit Sub
    
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
    
        altU = hMapData(map_x, map_y)
        
        tBottom = tTop + 31
        tRight = tLeft + 31
        tTop = tTop + 1
        tLeft = tLeft + 1
        With tBox 'With tBox
                .x0 = tLeft
                .y0 = tBottom - altU.hs(0)
                .color0 = Color
                .x1 = tLeft
                .y1 = tTop - altU.hs(1)
                .Color1 = Color
                .x2 = tRight
                .y2 = tBottom - altU.hs(2)
                .Color2 = Color
                .x3 = tRight
                .y3 = tTop - altU.hs(3)
                .color3 = Color
                .tu0 = 0
                .tv0 = 0
                .tu1 = 0
                .tv1 = 0
                .tu2 = 0
                .tv2 = 0
                .tu3 = 0
                .tv3 = 0
        End With
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
        'MapData(map_x, map_y).tile_render = 255
End Sub

Public Sub Grh_Render_Bloqueos(ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte)
'*********************************************
'Author: menduz
'*********************************************
    
    
    
    
    'If GrhIndex = 0 Then Exit Sub
    
    If Not InMapBounds(map_x, map_y) Then Exit Sub
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
    
    Dim tTrigger As Long
    
    tTrigger = mapdata(map_x, map_y).trigger
    
    If tTrigger Then
        If (tTrigger And TodosBordesBloqueados) = eTriggers.TodosBordesBloqueados Then
           Grh_Render_Blocked &H7FCC0000, tLeft, tTop, map_x, map_y
        Else
           If tTrigger And eTriggers.BloqueoEste Then _
                Draw_FilledBox tLeft + 30, tTop, 2, 32, &HFFFC0000, 0, 0
            If tTrigger And eTriggers.BloqueoOeste Then _
                Draw_FilledBox tLeft, tTop, 2, 32, &HFFFC0000, 0, 0
            If tTrigger And eTriggers.BloqueoNorte Then _
                Draw_FilledBox tLeft, tTop, 32, 2, &HFFFC0000, 0, 0
            If tTrigger And eTriggers.BloqueoSur Then _
                Draw_FilledBox tLeft, tTop + 30, 32, 2, &HFFFC0000, 0, 0
        End If
    End If
End Sub


Public Sub Grh_Render_new(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, Optional ByVal mirror As Byte = 0, Optional ByVal mirrorv As Byte = 0, Optional ByVal alpha As Byte = 0, Optional ByVal texturaOverwride As Integer = 0) ', Optional ByVal shadow As Byte = 1, Optional ByRef shadowoffx As Single)
'*********************************************
'Author: menduz
'*********************************************

    If GrhIndex = 0 Then Exit Sub
    
    If texturaOverwride > 0 Then
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(texturaOverwride)
    Else
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)
    End If
    
    If GrhData(GrhIndex).hardcor = 0 Then Init_grh_tutv GrhIndex
    
    Dim TGRH As GrhData
    TGRH = GrhData(GrhIndex)
    
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

        If alpha Then
            .color0 = (.color0 And &HFFFFFF) Or Alphas(alpha)
            .Color1 = (.Color1 And &HFFFFFF) Or Alphas(alpha)
            .Color2 = (.Color2 And &HFFFFFF) Or Alphas(alpha)
            .color3 = (.color3 And &HFFFFFF) Or Alphas(alpha)
        End If
        
        If mirror Then
            .tu0 = TGRH.tu(2)
            .tv0 = TGRH.tv(2)
            .tu1 = TGRH.tu(3)
            .tv1 = TGRH.tv(3)
            .tu2 = TGRH.tu(0)
            .tv2 = TGRH.tv(0)
            .tu3 = TGRH.tu(1)
            .tv3 = TGRH.tv(1)
        Else
'            .tu0 = TGRH.tu(0)
'            .tv0 = TGRH.tv(0)
'            .tu1 = TGRH.tu(1)
'            .tv1 = TGRH.tv(1)
'            .tu2 = TGRH.tu(2)
'            .tv2 = TGRH.tv(2)
'            .tu3 = TGRH.tu(3)
'            .tv3 = TGRH.tv(3)
            .tu0 = TGRH.tu(0)
            .tv0 = TGRH.tv(0)
            .tv1 = TGRH.tv(1)
            .tu2 = TGRH.tu(2)
            .tu1 = .tu0
            .tv2 = .tv0
            .tu3 = .tu2
            .tv3 = .tv1
            
        End If
        
        If mirrorv Then
            Dim bsf!
            bsf = .tv0
            .tv0 = .tv1
            .tv1 = bsf
            bsf = .tv2
            .tv2 = .tv3
            .tv3 = bsf
        End If

    End With
    
    
    Dim C3%, Vertical As Byte
    
    Vertical = 1
    
    Call Obtener_Texturas_Complementarias(CInt(GrhData(GrhIndex).filenum), 0, 0, C3)
    
    If usaBumpMapping Then
        With tBox
            .tu01 = .tu0
            .tu11 = .tu1
            .tu21 = .tu2
            .tu31 = .tu3
            .tv01 = .tv0
            .tv11 = .tv1
            .tv21 = .tv2
            .tv31 = .tv3
        End With

        If C3 = 0 Then
            C3 = LightTextureWall
            Vertical = 1
        End If

        If Vertical = 1 Then
            CalcularNormalColorTBOX_TexturaVertical tBox
            Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureVertical
        Else
            CalcularNormalColorTBOX_Textura tBox
            Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureHorizontal
        End If
        Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)
    End If

    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.estandar
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    Grh_Render_Complementario TGRH.filenum
    

    

End Sub



Public Sub Grh_Render_reflejo(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, Optional ByVal mirror As Byte = 0, Optional ByVal mirrorv As Byte = 0, Optional ByVal alpha As Byte = 0) ', Optional ByVal shadow As Byte = 1, Optional ByRef shadowoffx As Single)
'*********************************************
'Author: menduz
'*********************************************
    Dim dest_x2 As Integer, dest_y2 As Integer
    Dim TGRH As GrhData
    Dim bsf!
    Dim ta As Long
    If GrhIndex = 0 Then Exit Sub
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)

    If GrhData(GrhIndex).hardcor = 0 Then Init_grh_tutv GrhIndex
    TGRH = GrhData(GrhIndex)
    
    dest_y2 = dest_y + TGRH.pixelHeight
    dest_x2 = dest_x + TGRH.pixelWidth

    With tBox
        .x0 = dest_x
        .y0 = dest_y2 - ModSuperWaterDD(map_x, map_y).hs(0)
        .x1 = .x0
        .y1 = dest_y - ModSuperWaterDD(map_x, map_y).hs(1)
        .x2 = dest_x2
        .y2 = .y0 - ModSuperWaterDD(map_x, map_y).hs(2)
        .x3 = .x2
        .y3 = .y1 - ModSuperWaterDD(map_x, map_y).hs(3)
        
        .color0 = ResultColorArray(map_x, map_y)
        .Color1 = ResultColorArray(map_x, map_y - 1)
        .Color2 = ResultColorArray(map_x + 1, map_y)
        .color3 = ResultColorArray(map_x + 1, map_y - 1)
        
        If alpha Then
            ta = (alpha Mod 128) * &H1000000
            .color0 = (.color0 And &HFFFFFF) Or ta
            .Color1 = (.Color1 And &HFFFFFF) Or ta
            .Color2 = (.Color2 And &HFFFFFF) Or ta
            .color3 = (.color3 And &HFFFFFF) Or ta
        End If
        
        If mirror Then
            .tu0 = TGRH.tu(2)
            .tv0 = TGRH.tv(2)
            .tu1 = TGRH.tu(3)
            .tv1 = TGRH.tv(3)
            .tu2 = TGRH.tu(0)
            .tv2 = TGRH.tv(0)
            .tu3 = TGRH.tu(1)
            .tv3 = TGRH.tv(1)
        Else
            .tu0 = TGRH.tu(0)
            .tv0 = TGRH.tv(0)
            .tu1 = TGRH.tu(1)
            .tv1 = TGRH.tv(1)
            .tu2 = TGRH.tu(2)
            .tv2 = TGRH.tv(2)
            .tu3 = TGRH.tu(3)
            .tv3 = TGRH.tv(3)
        End If
        
        If mirrorv Then
            bsf = .tv0
            .tv0 = .tv1
            .tv1 = bsf
            bsf = .tv2
            .tv2 = .tv3
            .tv3 = bsf
        End If
    End With
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
End Sub



Public Sub Grh_Render_char(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, Optional ByVal mirror As Byte = 0) ', Optional ByVal shadow As Byte = 1, Optional ByRef shadowoffx As Single)
'*********************************************
'Author: menduz
'*********************************************
    Dim dest_x2 As Integer, dest_y2 As Integer
    Dim TGRH As GrhData
    If GrhIndex = 0 Then Exit Sub
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)

    If GrhData(GrhIndex).hardcor = 0 Then Init_grh_tutv GrhIndex
    TGRH = GrhData(GrhIndex)
    
    dest_y2 = dest_y + TGRH.pixelHeight
    dest_x2 = dest_x + TGRH.pixelWidth

    With tBox
        .x0 = dest_x
        .y0 = dest_y2
        .x1 = .x0
        .y1 = dest_y
        .x2 = dest_x2
        .y2 = .y0
        .x3 = .x2
        .y3 = .y1

        .color0 = ResultColorArray(map_x, map_y)
        .Color1 = ResultColorArray(map_x, map_y - 1)
        .Color2 = ResultColorArray(map_x + 1, map_y)
        .color3 = ResultColorArray(map_x + 1, map_y - 1)

        If mirror Then
            .tu0 = TGRH.tu(2)
            .tv0 = TGRH.tv(2)
            .tu1 = TGRH.tu(3)
            .tv1 = TGRH.tv(3)
            .tu2 = TGRH.tu(0)
            .tv2 = TGRH.tv(0)
            .tu3 = TGRH.tu(1)
            .tv3 = TGRH.tv(1)
        Else
            .tu0 = TGRH.tu(0)
            .tv0 = TGRH.tv(0)
            .tu1 = TGRH.tu(1)
            .tv1 = TGRH.tv(1)
            .tu2 = TGRH.tu(2)
            .tv2 = TGRH.tv(2)
            .tu3 = TGRH.tu(3)
            .tv3 = TGRH.tv(3)
        End If
    End With
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
End Sub

Public Sub Grh_Render_Simple_box(ByVal tex As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal Color As Long, ByVal size As Single, Optional ByVal alpha As Byte, Optional ByVal dw As Single)
'*********************************************
'Author: menduz
'viva la harcodeada, VIVA!
'*********************************************
    Dim ll As Long
    With tBox
        .x0 = tLeft
        .y0 = tTop + size
        .color0 = Color
        .x1 = tLeft
        .y1 = tTop
        .Color1 = Color
        .x2 = tLeft + size + dw
        .y2 = tTop + size
        .Color2 = Color
        .x3 = tLeft + size + dw
        .y3 = tTop
        .color3 = Color
        .tu0 = 0
        .tv0 = 1
        .tu1 = 0
        .tv1 = 0
        .tu2 = 1
        .tv2 = 1
        .tu3 = 1
        .tv3 = 0
    End With
    
    

    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Public Sub Grh_Render_Simple_rect(ByVal tex As Long, ByVal offsetX As Integer, ByVal offsetY As Integer, ByVal width As Integer, ByVal height As Integer, ByVal tLeft As Single, ByVal tTop As Single, ByVal Color As Long, ByVal realWith As Integer, ByVal realHeight As Integer, Optional ByVal alpha As Byte, Optional ByVal destWidth As Integer, Optional ByVal destHeight As Integer)
'*********************************************
'Author: menduz
'viva la harcodeada, VIVA!
'*********************************************
'*********************************************
'Author: menduz
'viva la harcodeada, VIVA!
'*********************************************
    Dim ll As Long
    Dim size As Integer
    size = 64
    With tBox
        .x0 = tLeft
        .y0 = tTop + destHeight
        .color0 = Color
        .x1 = tLeft
        .y1 = tTop
        .Color1 = Color
        .x2 = tLeft + destWidth
        .y2 = tTop + destHeight
        .Color2 = Color
        .x3 = tLeft + destWidth
        .y3 = tTop
        .color3 = Color
        .tu0 = offsetX / realWith
        .tv0 = (height + offsetY) / realHeight
        .tu1 = .tu0
        .tv1 = (offsetY) / realHeight
        .tu2 = (width + offsetX) / realWith
        .tv2 = height / realHeight
        .tu3 = .tu2
        .tv3 = .tv1
    End With
    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    

End Sub


Public Sub Grh_Render_Simple_box_offset(ByVal tex As Long, ByVal offsetX As Integer, ByVal offsetY As Integer, ByVal width As Integer, ByVal height As Integer, ByVal tLeft As Single, ByVal tTop As Single, ByVal Color As Long, ByVal size As Single, Optional ByVal alpha As Byte, Optional ByVal dw As Single)
'*********************************************
'Author: menduz
'viva la harcodeada, VIVA!
'*********************************************
    Dim ll As Long
    
    With tBox
        .x0 = tLeft
        .y0 = tTop + height
        .color0 = Color
        .x1 = tLeft
        .y1 = tTop
        .Color1 = Color
        .x2 = tLeft + width + dw
        .y2 = tTop + height
        .Color2 = Color
        .x3 = tLeft + width + dw
        .y3 = tTop
        .color3 = Color
        
        .tu0 = offsetX / size
        .tv0 = (offsetY + height) / size
        .tu1 = .tu0
        .tv1 = offsetY / size
        .tu2 = (offsetX + width) / size
        .tv2 = .tv0
        .tu3 = .tu2
        .tv3 = .tv1
    End With
    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(tex)
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    
    If alpha Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub



Public Sub Grh_Proyectil(ByVal GrhIndex As Long, ByVal dest_x As Single, ByVal dest_y As Single, Optional ByVal alpha As Byte = 0, Optional ByVal light_value As Long = &HFFFFFFFF, Optional ByVal Degrees As Integer)
'*********************************************
'Author: menduz
'*********************************************
    Static dest_rect As sRECT
    Static temp_verts(3) As TLVERTEX
    Dim centerX As Single
    Dim centerY As Single
    Dim Index As Integer
    Dim NewX As Single
    Dim NewY As Single
    Dim SinRad As Single
    Dim CosRad As Single

    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhIndex)

    With GrhData(GrhIndex)
        dest_rect.bottom = dest_y + .pixelHeight
        dest_rect.left = dest_x
        dest_rect.right = dest_x + .pixelWidth
        dest_rect.top = dest_y
        
        If .hardcor = 0 Then Init_grh_tutv GrhIndex

        Call cTLVertex(temp_verts(0), dest_rect.left, dest_rect.bottom, light_value, 0, 1)
        Call cTLVertex(temp_verts(1), dest_rect.left, dest_rect.top, light_value, 0, 0)
        Call cTLVertex(temp_verts(2), dest_rect.right, dest_rect.bottom, light_value, 1, 1)
        Call cTLVertex(temp_verts(3), dest_rect.right, dest_rect.top, light_value, 1, 0)

        If Degrees > 0 And Degrees < 360 Then
            'Converts the angle to rotate by into radians
            'Set the CenterX and CenterY values
            centerX = dest_x + (.pixelHeight * 0.5)
            centerY = dest_y + (.pixelWidth * 0.5)
            'Pre-calculate the cosine and sine of the radiant
            SinRad = Seno(Degrees)
            CosRad = Coseno(Degrees)
            'Loops through the passed vertex buffer
            For Index = 0 To 3
                NewX = centerX + (temp_verts(Index).v.X - centerX) * CosRad - (temp_verts(Index).v.Y - centerY) * SinRad
                NewY = centerY + (temp_verts(Index).v.Y - centerY) * CosRad + (temp_verts(Index).v.X - centerX) * SinRad
                temp_verts(Index).v.X = NewX
                temp_verts(Index).v.Y = NewY
            Next Index
        End If
    End With
    
    If alpha Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), TL_size
    If alpha Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub


Public Sub Grh_Render_size(ByVal GrhIndex As Long, ByVal dest_x As Single, ByVal dest_y As Single, Optional ByVal alpha As Byte = 0, Optional ByVal light_value As Long = &HFFFFFFFF, Optional ByVal Degrees As Integer, Optional ByVal width As Integer = 0, Optional ByVal height As Integer = 0, Optional ByVal recortar As Boolean = False, Optional ByVal offsetX As Integer = 0, Optional ByVal offsetY As Integer = 0)
'*********************************************
'Author: menduz
'*********************************************
    Static dest_rect As sRECT
    Static temp_verts(3) As TLVERTEX
    Dim centerX As Single
    Dim centerY As Single
    Dim Index As Integer
    Dim NewX As Single
    Dim NewY As Single
    Dim SinRad As Single
    Dim CosRad As Single
    Dim width_ As Integer
    Dim height_ As Integer
    
    Dim recorteX As Single
    
    If GrhIndex = 0 Then Exit Sub
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)


    With GrhData(GrhIndex)
    
        If width = 0 Then width_ = .pixelWidth Else width_ = width
        If height = 0 Then height_ = .pixelHeight Else height_ = height
         
        dest_rect.bottom = dest_y + height_
        dest_rect.left = dest_x
        dest_rect.right = dest_x + width_
        dest_rect.top = dest_y
        
        Dim h!, w!
               
        If recortar And width > 0 Then
            Call GetTextureDimension(GrhData(GrhIndex).filenum, h, w)
        
            .tu(0) = offsetX / w
            .tv(0) = (.sy + height_) / h
            .tu(1) = .tu(0)
            .tv(1) = .sy / h
            .tu(2) = (offsetX + width_) / w
            .tv(2) = .tv(0)
            .tu(3) = .tu(2)
            .tv(3) = .tv(1)
            .hardcor = 1
        Else
              If .hardcor = 0 Then Init_grh_tutv GrhIndex
        End If

        Call cTLVertex(temp_verts(0), dest_rect.left, dest_rect.bottom, light_value Or Alphas(alpha), .tu(0), .tv(0))
        Call cTLVertex(temp_verts(1), dest_rect.left, dest_rect.top, light_value Or Alphas(alpha), .tu(1), .tv(1))
        Call cTLVertex(temp_verts(2), dest_rect.right, dest_rect.bottom, light_value Or Alphas(alpha), .tu(2), .tv(2))
        Call cTLVertex(temp_verts(3), dest_rect.right, dest_rect.top, light_value Or Alphas(alpha), .tu(3), .tv(3))

        If Degrees > 0 And Degrees < 360 Then
            'Converts the angle to rotate by into radians
            'Set the CenterX and CenterY values
            centerX = dest_x + (.pixelHeight * 0.5)
            centerY = dest_y + (.pixelWidth * 0.5)
            'Pre-calculate the cosine and sine of the radiant
            SinRad = Seno(Degrees)
            CosRad = Coseno(Degrees)
            'Loops through the passed vertex buffer
            For Index = 0 To 3
                NewX = centerX + (temp_verts(Index).v.X - centerX) * CosRad - (temp_verts(Index).v.Y - centerY) * SinRad
                NewY = centerY + (temp_verts(Index).v.Y - centerY) * CosRad + (temp_verts(Index).v.X - centerX) * SinRad
                temp_verts(Index).v.X = NewX
                temp_verts(Index).v.Y = NewY
            Next Index
        End If
    End With
    
    If alpha Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), TL_size
    If alpha Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Public Sub Grh_Render_nocolor(ByVal GrhIndex As Long, ByVal dest_x As Single, ByVal dest_y As Single, Optional ByVal alpha As Byte = 0, Optional ByVal light_value As Long = &HFFFFFFFF, Optional ByVal Degrees As Integer)
'*********************************************
'Author: menduz
'*********************************************
    Static dest_rect As sRECT
    Static temp_verts(3) As TLVERTEX
    Dim centerX As Single
    Dim centerY As Single
    Dim Index As Integer
    Dim NewX As Single
    Dim NewY As Single
    Dim SinRad As Single
    Dim CosRad As Single

    If GrhIndex = 0 Then Exit Sub
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)

    With GrhData(GrhIndex)
        dest_rect.bottom = dest_y + .pixelHeight
        dest_rect.left = dest_x
        dest_rect.right = dest_x + .pixelWidth
        dest_rect.top = dest_y
        
        If .hardcor = 0 Then Init_grh_tutv GrhIndex
        
        Call cTLVertex(temp_verts(0), dest_rect.left, dest_rect.bottom, light_value, .tu(0), .tv(0))
        Call cTLVertex(temp_verts(1), dest_rect.left, dest_rect.top, light_value, .tu(1), .tv(1))
        Call cTLVertex(temp_verts(2), dest_rect.right, dest_rect.bottom, light_value, .tu(2), .tv(2))
        Call cTLVertex(temp_verts(3), dest_rect.right, dest_rect.top, light_value, .tu(3), .tv(3))

        If Degrees > 0 And Degrees < 360 Then
            'Converts the angle to rotate by into radians
            'Set the CenterX and CenterY values
            centerX = dest_x + (.pixelHeight * 0.5)
            centerY = dest_y + (.pixelWidth * 0.5)
            'Pre-calculate the cosine and sine of the radiant
            SinRad = Seno(Degrees)
            CosRad = Coseno(Degrees)
            'Loops through the passed vertex buffer
            For Index = 0 To 3
                NewX = centerX + (temp_verts(Index).v.X - centerX) * CosRad - (temp_verts(Index).v.Y - centerY) * SinRad
                NewY = centerY + (temp_verts(Index).v.Y - centerY) * CosRad + (temp_verts(Index).v.X - centerX) * SinRad
                temp_verts(Index).v.X = NewX
                temp_verts(Index).v.Y = NewY
            Next Index
        End If
    End With
    
    If alpha Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), TL_size
    If alpha Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub


Public Sub Grh_Render_Complementario(ByVal tex As Integer)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
Dim C1%, C2%

If Obtener_Texturas_Complementarias(tex, C1, C2) = True Then
    If C1 > 0 Then
        With tBox
            .color0 = mzWhite
            .Color1 = mzWhite
            .Color2 = mzWhite
            .color3 = mzWhite
        End With
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(C1)
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    End If
    If C2 > 0 Then
        With tBox
            .color0 = base_light
            .Color1 = base_light
            .Color2 = base_light
            .color3 = base_light
        End With
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(C2)
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    End If
End If
End Sub

Public Sub Grh_Render_relieve(ByVal GrhIndex As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!
    
    Dim TGRH As GrhData
    Dim altU As AUDT

    
    If GrhIndex = 0 Then Exit Sub
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(GrhIndex).filenum)
    
   
    If GrhData(GrhIndex).hardcor = 0 Then Init_grh_tutv GrhIndex
    TGRH = GrhData(GrhIndex)
    altU = hMapData(map_x, map_y)
    
    tBottom = tTop + TGRH.pixelHeight
    tRight = tLeft + TGRH.pixelWidth
    

    With tBox 'With tBox
        .x0 = tLeft
        .y0 = tBottom - altU.hs(0)
        .x1 = tLeft
        .y1 = tTop - altU.hs(1)
        .x2 = tRight
        .y2 = tBottom - altU.hs(2)
        .x3 = tRight
        .y3 = tTop - altU.hs(3)
        .tu0 = TGRH.tu(0)
        .tv0 = TGRH.tv(0)
        .tu1 = TGRH.tu(1)
        .tv1 = TGRH.tv(1)
        .tu2 = TGRH.tu(2)
        .tv2 = TGRH.tv(2)
        .tu3 = TGRH.tu(3)
        .tv3 = TGRH.tv(3)
    End With
    
    Colorear_TBOX tBox, map_x, map_y
    
    ' INICIO NORMALES
    ' +++
    Dim C3%

    Call Obtener_Texturas_Complementarias(CInt(GrhData(GrhIndex).filenum), 0, 0, C3)

    If C3 = 0 Then C3 = LightTextureFloor

    CalcularNormalColorTBOX_Textura tBox
    Engine_PixelShaders_SetTexture_Ambient Engine_LightsTexture.LightsTextureHorizontal
    Engine_PixelShaders_SetTexture_Normal PeekTexture(C3)
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.estandar
    ' ---
    'Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    ' FIN NORMALES
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    Grh_Render_Complementario GrhData(GrhIndex).filenum

End Sub

