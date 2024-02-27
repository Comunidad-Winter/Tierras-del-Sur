Attribute VB_Name = "Engine_Montana"
Option Explicit

Public Sub ActualizarArraysAlturasMapas()
    'If MapInfo.UsaAguatierra Then
        Dim x As Integer, y As Integer
        
        For y = Y_MINIMO_VISIBLE + 1 To Y_MAXIMO_VISIBLE - 1
            For x = X_MINIMO_VISIBLE + 1 To X_MAXIMO_VISIBLE - 1
                'If GrhData(GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(1)).FileNum Then
                '    PreLoadTexture GrhData(GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(1)).FileNum
                'End If
                
                hMapData(x, y + 1).hs(1) = hMapData(x, y).hs(0)
                hMapData(x - 1, y + 1).hs(3) = hMapData(x, y).hs(0)
                hMapData(x - 1, y).hs(2) = hMapData(x, y).hs(0)
                
                mapdata(x, y).is_water = CBool(hMapData(x, y).hs(0) < mapinfo.agua_profundidad)
                
                If Not mapdata(x, y).is_water Then
                    ModSuperWaterMM(x, y + 1).hs(3) = 0
                    ModSuperWaterMM(x, y + 1).hs(1) = 0
    
                    ModSuperWaterMM(x, y - 1).hs(2) = 0
                    ModSuperWaterMM(x, y - 1).hs(0) = 0
    
                    ModSuperWaterMM(x + 1, y).hs(0) = 0
                    ModSuperWaterMM(x + 1, y).hs(1) = 0
    
                    ModSuperWaterMM(x - 1, y).hs(2) = 0
                    ModSuperWaterMM(x - 1, y).hs(3) = 0
    
                    ModSuperWaterMM(x + 1, y - 1).hs(0) = 0
                    ModSuperWaterMM(x + 1, y + 1).hs(1) = 0
                    ModSuperWaterMM(x - 1, y - 1).hs(2) = 0
                    ModSuperWaterMM(x - 1, y + 1).hs(3) = 0
                Else
                    mapdata(x, y + 1).is_water = True
                    mapdata(x - 1, y + 1).is_water = True
                    mapdata(x - 1, y).is_water = True
                End If
    
                ModSuperWaterMM(x, y + 1).hs(1) = ModSuperWaterMM(x, y).hs(0)
                ModSuperWaterMM(x - 1, y + 1).hs(3) = ModSuperWaterMM(x, y).hs(0)
                ModSuperWaterMM(x - 1, y).hs(2) = ModSuperWaterMM(x, y).hs(0)
    
                ModSuperWaterDD(x, y).hs(0) = 0
                ModSuperWaterDD(x, y + 1).hs(1) = 0
                ModSuperWaterDD(x - 1, y + 1).hs(3) = 0
                ModSuperWaterDD(x - 1, y).hs(2) = 0
            Next x
        Next y
        
        Call Engine_Landscape_Water.RemakeWaterTilenumbers(mapinfo.agua_rect.left, mapinfo.agua_rect.top, mapinfo.agua_rect.right, mapinfo.agua_rect.bottom)
        
        'Call Engine_Landscape_Water.recalcular_colores_agua
        'Call Engine_Landscape_Water.recalcular_opacidades_agua
    'End If
End Sub

Public Sub Compute_Mountain()
    
    CalcularNormales
    
    Engine_ColoresAgua.ColoresAgua_Redraw

    Engine_Landscape.Light_Update_Map = True
    Light_Update_Sombras = True

End Sub

