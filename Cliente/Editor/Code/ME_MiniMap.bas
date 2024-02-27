Attribute VB_Name = "ME_MiniMap"
Option Explicit

Private miniMapTexture As Direct3DTexture8

Public Enum eMiniMapaTipo
    emmBloqueos = 1
    emmLuces = 2
    emmNPC = 4
    emmTriggers = 8
    emmAcciones = 16
    emmColores = 32
    emmPiso = 64
End Enum

Private miniMapSurfaceDesc As D3DSURFACE_DESC

Public miniMapaTipo As eMiniMapaTipo
Private miniMapVisible_ As Boolean

Public MiniMapNeedToBeRedrawed As Boolean

Private vwMiniMapa As vwMiniMapa

Public Sub miniMapInit()
    Set miniMapTexture = D3DX.CreateTexture(D3DDevice, 256, 256, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
    Set vwMiniMapa = New vwMiniMapa
    
    
    miniMapTexture.GetLevelDesc 0, miniMapSurfaceDesc
    
    'Debug.Print "tamaño textura minimapa: "; miniMapSurfaceDesc.size; "bytes"
End Sub

Public Sub miniMap_Redraw()
    Dim ColorBorde As BGRACOLOR_DLL
    Dim tick As Long
    Dim lockedRect As D3DLOCKED_RECT
    Dim imagen(0 To 255, 0 To 255) As BGRACOLOR_DLL
    Dim x%, y%

    If Not miniMapVisible Then Exit Sub
    
    If Engine_Escene_Abierta = False Then
        MiniMapNeedToBeRedrawed = True
        Exit Sub
    End If
    
'    If miniMapaTipo And emmLuces Then
'        For x = 1 To 218
'            DXCopyMemory imagen(x, 1), ResultColorArray(x, 1), 217& * 4&
'        Next
'    End If
    
    tick = GetTimer
    
    ColorBorde.a = 128
    ColorBorde.r = 100 + Abs(Sin(tick / 920) * 50)
    ColorBorde.g = 100 + Abs(Sin(tick / 480) * 50)
    ColorBorde.b = 150 + Abs(Cos(tick / 700) * 100)
    
    ' Dibujamos el borde de la pantalla
    ' Dos for para que quede mas claro
    For x = SV_Constantes.X_MINIMO_VISIBLE - 1 To SV_Constantes.X_MAXIMO_VISIBLE + 1
        imagen(x, SV_Constantes.Y_MINIMO_VISIBLE - 1) = ColorBorde
        imagen(x, SV_Constantes.Y_MAXIMO_VISIBLE + 1) = ColorBorde
    Next
    
    For y = SV_Constantes.X_MINIMO_VISIBLE - 1 To SV_Constantes.X_MAXIMO_VISIBLE + 1
        imagen(SV_Constantes.X_MINIMO_VISIBLE - 1, y) = ColorBorde
        imagen(SV_Constantes.X_MAXIMO_VISIBLE + 1, y) = ColorBorde
    Next
    
    ' Establecemos por cada punto el color que va dependiendo lo que haya en ese tile
    
    For x = SV_Constantes.X_MINIMO_VISIBLE To SV_Constantes.X_MAXIMO_VISIBLE
        For y = SV_Constantes.Y_MINIMO_VISIBLE To SV_Constantes.Y_MAXIMO_VISIBLE
        
            imagen(x, y).a = 50
        
            If miniMapaTipo And emmColores Then
                imagen(x, y) = OriginalMapColor(x, y)
            ElseIf miniMapaTipo And emmLuces Then
                DXCopyMemory imagen(x, y), ResultColorArray(x, y), 4&
            ElseIf miniMapaTipo Then
                
            End If
 
            If (miniMapaTipo And emmBloqueos) And ((mapdata(x, y).Trigger And eTriggers.TodosBordesBloqueados) > 0) Then
                imagen(x, y).r = 255
                imagen(x, y).a = 255
            ElseIf (miniMapaTipo And emmNPC) And (mapdata(x, y).NpcIndex <> 0) Then
                imagen(x, y).g = 255
                imagen(x, y).b = 255
                imagen(x, y).a = 255
            ElseIf (miniMapaTipo And emmAcciones) And (Not mapdata(x, y).accion Is Nothing) Then
                imagen(x, y).b = 255
                imagen(x, y).r = 255
                imagen(x, y).a = 255
            ElseIf (miniMapaTipo And emmTriggers) And (mapdata(x, y).Trigger > 0) Then
                imagen(x, y).g = 255
                imagen(x, y).b = 255
                imagen(x, y).r = 255
                imagen(x, y).a = 255
            ElseIf (miniMapaTipo And emmPiso) And (mapdata(x, y).tile_texture > 0) Then
                imagen(x, y).g = 100
                imagen(x, y).b = 100
                imagen(x, y).r = 100
                imagen(x, y).a = 100
            End If

        Next
    Next
    
    miniMapTexture.LockRect 0, lockedRect, ByVal 0, 0
    
    DXCopyMemory ByVal lockedRect.pBits, imagen(0, 0), miniMapSurfaceDesc.Size

    miniMapTexture.UnlockRect 0
    
    MiniMapNeedToBeRedrawed = False
    
    'MiniMapNeedToBeRedrawed = True
End Sub

Public Sub miniMap_Render(ByVal tLeft As Single, ByVal tTop As Single, Optional ByVal MostrarNegro As Boolean)
'*********************************************
'Author: menduz
'*********************************************
    If Not miniMapVisible Then Exit Sub
    
    If MiniMapNeedToBeRedrawed Then miniMap_Redraw
    
    Dim tBottom!, tRight! ', tTop!, tLeft!
    
    Dim tBox As Box_Vertex
    Dim Color As Long
    Color = -1
    
    
    
    tBottom = tTop + 256
    tRight = tLeft + 256
    
    
    With tBox 'With tBox
            .x0 = tLeft
            .y0 = tBottom
            .color0 = Color
            .x1 = tLeft
            .y1 = tTop
            .Color1 = Color
            .X2 = tRight
            .Y2 = tBottom
            .Color2 = Color
            .x3 = tRight
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
            .rhw0 = 1
            .rhw1 = 1
            .rhw2 = 1
            .rhw3 = 1
    End With
    
    If MostrarNegro Then
        With tBox 'With tBox
            .y0 = tTop + 220
            .color0 = mzBlack
            .Color1 = mzBlack
            .X2 = tLeft + 220
            .Y2 = tTop + 220
            .Color2 = mzBlack
            .x3 = tLeft + 220
            .color3 = mzBlack
        End With
    
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
        
        With tBox 'With tBox
            .y0 = tBottom
            .color0 = Color
            .Color1 = Color
            .X2 = tRight
            .Y2 = tBottom
            .Color2 = Color
            .x3 = tRight
            .color3 = Color
        End With
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse miniMapTexture
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
    
    Color = &H44FFFFFF
    
    tTop = tTop + UserPos.y - HalfWindowTileHeight
    tLeft = tLeft + UserPos.x - HalfWindowTileWidth
    
    tBottom = tTop + WindowTileHeight
    tRight = tLeft + WindowTileWidth
    
    
    With tBox 'With tBox
            .x0 = tLeft
            .y0 = tBottom
            .color0 = Color
            .x1 = tLeft
            .y1 = tTop
            .Color1 = Color
            .X2 = tRight
            .Y2 = tBottom
            .Color2 = Color
            .x3 = tRight
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
            .rhw0 = 1
            .rhw1 = 1
            .rhw2 = 1
            .rhw3 = 1
    End With

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size



End Sub

Public Property Get miniMapVisible() As Boolean
    miniMapVisible = miniMapVisible_
End Property

Public Property Let miniMapVisible(ByVal vNewValue As Boolean)
    miniMapVisible_ = vNewValue
    If miniMapVisible_ Then
        miniMap_Redraw
        GUI_Load vwMiniMapa
    Else
        GUI_Quitar vwMiniMapa
    End If
End Property
