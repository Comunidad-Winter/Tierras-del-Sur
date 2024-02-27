Attribute VB_Name = "Engine_Landscape_Water"
' ESTE ARCHIVO ESTA COMPARTIDO POR TODOS LOS PROGRAMAS.

''
' @require Engine.bas
' @require Engine_Landscape.bas




Option Explicit

'AGUA DIFUMADA
    Public OpacidadesAgua(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)         As Byte
    Public ResultColorArrayAgua(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)    As Long
    Public AguaVisiblePosicion(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)    As Long
'/AGUA DIFUMADA

'AGUA DINÁMICA(?)
    Public ModSuperWater(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)            As Byte
    Public ModSuperWaterDD(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)            As AUDT
    Public ModSuperWaterMM(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)            As AUDT
    Public kWATER As Integer
    Public TileNumberWater(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)            As Byte
    Public AguaBoxes(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)                  As Box_Vertex
    Public Water_Actualizar_Vertices                    As Boolean
'/AUGA

Function HayAgua(ByRef tile As MapBlock) As Boolean
    ' TODO-NV
    'HayAgua = (AlturaPie(x, y) + 16) < AlturaAgua
    
    ' Version Actual
    If tile.Graphic(1).GrhIndex >= 1505 And _
       tile.Graphic(1).GrhIndex <= 1520 And _
       tile.Graphic(2).GrhIndex = 0 Then
        HayAgua = True
    Else
        HayAgua = False
    End If
    
End Function

'MAP EDITOR
Public Sub recalcular_opacidades_agua()
    Dim x%, y%, h!, d!
    'Marce On error resume next
    
    For y = Y_MINIMO_VISIBLE + 1 To Y_MAXIMO_VISIBLE - 1
        For x = X_MINIMO_VISIBLE + 1 To X_MAXIMO_VISIBLE - 1
            
            mapdata(x, y).is_water = hMapData(x, y).hs(0) < mapinfo.agua_profundidad
            
            If mapdata(x, y).is_water Then
                h = Abs(hMapData(x, y).hs(0) - mapinfo.agua_profundidad)
                
                OpacidadesAgua(x, y) = CByte(mins(h * 7, 255))
                
                d = -((hMapData(x, y).hs(0) - mapinfo.agua_profundidad) / 64)
                
                d = maxs(mins(d, 1), 0.1)
            Else
                d = 0
                OpacidadesAgua(x, y) = 0
            End If
            
            ModSuperWaterMM(x, y).hs(0) = d
            ModSuperWaterMM(x, y + 1).hs(1) = d
            ModSuperWaterMM(x - 1, y + 1).hs(3) = d
            ModSuperWaterMM(x - 1, y).hs(2) = d
            
            mapdata(x, y + 1).is_water = mapdata(x, y).is_water
            mapdata(x - 1, y + 1).is_water = mapdata(x, y).is_water
            mapdata(x - 1, y).is_water = mapdata(x, y).is_water

            ModSuperWaterDD(x, y).hs(0) = 0
            ModSuperWaterDD(x, y + 1).hs(1) = 0
            ModSuperWaterDD(x - 1, y + 1).hs(3) = 0
            ModSuperWaterDD(x - 1, y).hs(2) = 0
            
        Next x
    Next y
End Sub

Public Sub recalcular_colores_agua()
    Dim x%, y%, c&
    
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            AguaVisiblePosicion(x, y) = 0
        Next y
    Next x
    
    
    
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            'If AguaVisiblePosicion(X, Y) Then
                AguaBoxes(x, y) = Tileset_Grh_Array(TileNumberWater(x, y))
                With AguaBoxes(x, y) 'With tBox
                    .rhw0 = 1
                    .rhw1 = 1
                    .rhw2 = 1
                    .rhw3 = 1
                    .color0 = 0
                    .Color1 = 0
                    .Color2 = 0
                    .color3 = 0
                    
                    .tu02 = x / 256
                    .tu12 = .tu02
                    .tu22 = (x + 1) / 256
                    .tu32 = .tu22
                    
                    .tv02 = (y + 1) / 256
                    .tv12 = y / 256
                    .tv22 = .tv02
                    .tv32 = .tv12
                    
                    If OpacidadesAgua(x, y) Then
                        .color0 = (ResultColorArray(x, y) And &HFFFFFF) Or Alphas(OpacidadesAgua(x, y)) ' ResultColorArrayAgua(X, Y)
                    End If
                    
                    If x < X_MAXIMO_VISIBLE Then
                        If OpacidadesAgua(x + 1, y) Then
                            .Color2 = (ResultColorArray(x + 1, y) And &HFFFFFF) Or Alphas(OpacidadesAgua(x + 1, y))
                        End If
                        
                        If y > Y_MINIMO_VISIBLE Then
                            If OpacidadesAgua(x + 1, y - 1) Then
                                .color3 = (ResultColorArray(x + 1, y - 1) And &HFFFFFF) Or Alphas(OpacidadesAgua(x + 1, y - 1))
                            End If
                            If OpacidadesAgua(x, y - 1) Then
                                .Color1 = (ResultColorArray(x, y - 1) And &HFFFFFF) Or Alphas(OpacidadesAgua(x, y - 1))
                            End If
                        End If
                    End If
                    
                    AguaVisiblePosicion(x, y) = .color0 <> 0 Or .Color1 <> 0 Or .Color2 <> 0 Or .color3 <> 0
                End With
                
            'End If
        Next y
    Next x
End Sub

Public Sub map_render_kwateR()
'funcion puta veo tu futuro en C muajaja
    Dim y As Integer
    Dim x As Integer
    Dim T As Single
    Dim ta As Byte
    If MinX = 0 Then Exit Sub
    ta = MinX Mod 2
    
    Static LastUpdate As Long
    Dim TgT As Long
    TgT = GetTimer
    If LastUpdate + 60 < TgT Then
        LastUpdate = TgT
        For y = MinY To MaxY - 1
            For x = MinX + 1 To MaxX
                If mapdata(x, y).is_water Then
                    If ta Then
                        If x And 1 Then
                            T = Seno(kWATER)
                        Else
                            T = -Seno(kWATER)
                        End If
                    Else
                        If x And 1 Then
                            T = -Coseno(kWATER)
                        Else
                            T = Coseno(kWATER)
                        End If
                    End If
                    T = T * 4
                    ModSuperWaterDD(x, y).hs(0) = T * ModSuperWaterMM(x, y).hs(0) + mapinfo.agua_profundidad
                    
                    ModSuperWaterDD(x, y + 1).hs(1) = T * ModSuperWaterMM(x, y + 1).hs(1) + mapinfo.agua_profundidad
                    ModSuperWaterDD(x - 1, y + 1).hs(3) = T * ModSuperWaterMM(x - 1, y + 1).hs(3) + mapinfo.agua_profundidad
                    ModSuperWaterDD(x - 1, y).hs(2) = T * ModSuperWaterMM(x - 1, y).hs(2) + mapinfo.agua_profundidad
                    ta = y Mod 2
                End If
            Next x
        Next y
        Water_Actualizar_Vertices = True
    Else
        Water_Actualizar_Vertices = False
    End If
    
End Sub

Public Sub Grh_Render_Water(ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte)
'*********************************************
'Author: menduz 324
'*********************************************
    Dim tBottom!, tRight!
    Dim altU As AUDT
        
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(Tilesets(mapinfo.agua_tileset).filenum)
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Ambient Engine_ColoresAgua.ColoresAguaTexture
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Agua
    
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
    D3DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 2, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    
    If Water_Actualizar_Vertices Or Cachear_Tiles Then
        tBottom = tTop + 32
        tRight = tLeft + 32
        
        altU = ModSuperWaterDD(map_x, map_y)
        
        With AguaBoxes(map_x, map_y) 'With tBox
            .x0 = tLeft
            .y0 = tBottom - mapinfo.agua_profundidad
            .x1 = tLeft
            .y1 = tTop - mapinfo.agua_profundidad
            .x2 = tRight
            .y2 = tBottom - mapinfo.agua_profundidad
            .x3 = tRight
            .y3 = tTop - mapinfo.agua_profundidad
            
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, .x0, TL_size
        End With
    Else
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, AguaBoxes(map_x, map_y), TL_size
    End If
    
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
    D3DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_NONE
    D3DDevice.SetTextureStageState 2, D3DTSS_MAGFILTER, D3DTEXF_NONE
End Sub

Public Sub RemakeWaterTilenumbers(ByVal x1 As Byte, ByVal y1 As Byte, ByVal x2 As Byte, ByVal y2 As Byte, Optional SaveVal As Boolean = False)
Dim x As Integer
Dim y As Integer

Dim alto As Integer
Dim ancho As Integer
alto = y2 - y1
ancho = x2 - x1

If alto = 0 Or ancho = 0 Then Exit Sub

If SaveVal Then
    mapinfo.agua_rect.top = y1
    mapinfo.agua_rect.bottom = y2
    mapinfo.agua_rect.left = x1
    mapinfo.agua_rect.right = x2
End If

For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        TileNumberWater(x, y) = 16 * y1 + x1 + ((x - 1) Mod ancho) + (((y - 1) Mod alto) * 16)
    Next y
Next x

Water_Actualizar_Vertices = True

End Sub

