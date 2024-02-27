Attribute VB_Name = "Engine_Map_ColorsIO"
Option Explicit

Public Function Engine_MapColorsIO_GetHeightMapPath() As String
    'Engine_MapColorsIO_GetHeightMapPath = app.Path & "\Datos\Mapas\HeightMaps\" & frmMain.zonaActual

    If Not FileExist(Engine_MapColorsIO_GetHeightMapPath, vbDirectory) Then MkDir Engine_MapColorsIO_GetHeightMapPath
    
    Engine_MapColorsIO_GetHeightMapPath = Engine_MapColorsIO_GetHeightMapPath & "\" & mapinfo.numero & ".bmp"
End Function

Public Function Engine_MapColorsIO_GetLightMapPath() As String
   ' Engine_MapColorsIO_GetLightMapPath = app.Path & "\Datos\Mapas\LightMaps\" & frmMain.zonaActual

    If Not FileExist(Engine_MapColorsIO_GetLightMapPath, vbDirectory) Then MkDir Engine_MapColorsIO_GetHeightMapPath
    
    Engine_MapColorsIO_GetLightMapPath = Engine_MapColorsIO_GetLightMapPath & "\" & mapinfo.numero
End Function


Public Sub Engine_MapColorsIO_SaveHeightMap()
    Dim Path As String
    Dim MyTexture As Direct3DTexture8
    Dim lockedRect As D3DLOCKED_RECT
    
    Dim ColoresHeightMap(255, 255) As BGRACOLOR_DLL
    
    Set MyTexture = D3DX.CreateTexture(D3DDevice, 256, 256, 1, 0, D3DFMT_X8R8G8B8, D3DPOOL_SYSTEMMEM)

    Path = Engine_MapColorsIO_GetHeightMapPath
    
    Dim x As Long
    Dim y As Long
    
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            ColoresHeightMap(x, y).r = minl(maxl(hMapData(x, y).hs(1) / 4 + 127, 0), 255)
            ColoresHeightMap(x, y).g = ColoresHeightMap(x, y).r
            ColoresHeightMap(x, y).b = ColoresHeightMap(x, y).r
            ColoresHeightMap(x, y).a = 255
        Next y
    Next x
    
    MyTexture.LockRect 0, lockedRect, ByVal 0, 0

    DXCopyMemory ByVal lockedRect.pBits, ColoresHeightMap(0, 0), 262144 ' 256 * 256 * 4

    MyTexture.UnlockRect 0
    Dim a As PALETTEENTRY
    a.red = 255
    a.green = 255
    a.blue = 255
    a.flags = 255
    D3DX.SaveTextureToFile Path, D3DXIFF_BMP, MyTexture, a
    
    LogDebug "Heightmap guardado!: " + Path
End Sub

Public Sub Engine_MapColorsIO_LoadHeightMap()
    Dim Path As String
    Dim MyTexture As Direct3DTexture8
    Dim lockedRect As D3DLOCKED_RECT
    
    Dim ColoresHeightMap(255, 255) As BGRACOLOR_DLL
    
    Path = Engine_MapColorsIO_GetHeightMapPath
    
    Set MyTexture = D3DX.CreateTextureFromFileEx(D3DDevice, Path, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFFFF00, ByVal 0, ByVal 0)

    Path = Engine_MapColorsIO_GetHeightMapPath
    
    
    MyTexture.LockRect 0, lockedRect, ByVal 0, 0

    DXCopyMemory ColoresHeightMap(0, 0), ByVal lockedRect.pBits, 262144 ' 256 * 256 * 4

    MyTexture.UnlockRect 0
    
    
    
    Dim x As Long
    Dim y As Long
    
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            Dim altura As Integer
            altura = (ColoresHeightMap(x, y).r - 127) * 4
            Alturas(x, y) = altura
            hMapData(x, y).hs(0) = altura
            hMapData(x, y).h = altura
            If InMapBounds(x, y + 1) Then
                hMapData(x, y + 1).hs(1) = altura
                If InMapBounds(x - 1, y) Then hMapData(x - 1, y + 1).hs(3) = altura
            End If
            If InMapBounds(x - 1, y) Then hMapData(x - 1, y).hs(2) = altura
        Next y
    Next x
    
    LogDebug "Heightmap cargado!: " + Path
    
    CalcularNormales

End Sub

Public Sub Engine_MapColorsIO_SaveLightMap()
    Dim Path As String
    Dim MyTexture As Direct3DTexture8
    Dim lockedRect As D3DLOCKED_RECT
    
    Dim ColoresHeightMap(255, 255) As BGRACOLOR_DLL
    
    Set MyTexture = D3DX.CreateTexture(D3DDevice, 256, 256, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_SYSTEMMEM)

    Path = Engine_MapColorsIO_GetLightMapPath & ".dds"
    
    Dim x As Long
    Dim y As Long
    
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            ColoresHeightMap(x, y).r = OriginalMapColor(x, y).r
            ColoresHeightMap(x, y).g = OriginalMapColor(x, y).g
            ColoresHeightMap(x, y).b = OriginalMapColor(x, y).b
            ColoresHeightMap(x, y).a = Intensidad_Del_Terreno(x, y)
        Next y
    Next x
    
    MyTexture.LockRect 0, lockedRect, ByVal 0, 0

    DXCopyMemory ByVal lockedRect.pBits, ColoresHeightMap(0, 0), 262144 ' 256 * 256 * 4

    MyTexture.UnlockRect 0
    
    Dim a As PALETTEENTRY
    a.red = 255
    a.green = 255
    a.blue = 255
    a.flags = 255
    D3DX.SaveTextureToFile Path, D3DXIFF_DDS, MyTexture, a
    
    LogDebug "Heightmap guardado!: " + Path
End Sub

Public Sub Engine_MapColorsIO_LoadLightMap()
    Dim Path As String
    Dim MyTexture As Direct3DTexture8
    Dim lockedRect As D3DLOCKED_RECT
    
    Dim ColoresHeightMap(255, 255) As BGRACOLOR_DLL
    
    Path = Engine_MapColorsIO_GetLightMapPath & ".png"
    
    Set MyTexture = D3DX.CreateTextureFromFileEx(D3DDevice, Path, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFFFF00, ByVal 0, ByVal 0)

    MyTexture.LockRect 0, lockedRect, ByVal 0, 0

    DXCopyMemory ColoresHeightMap(0, 0), ByVal lockedRect.pBits, 262144 ' 256 * 256 * 4

    MyTexture.UnlockRect 0


    Dim x As Long
    Dim y As Long
    
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            OriginalMapColor(x, y).r = ColoresHeightMap(x, y).r
            OriginalMapColor(x, y).g = ColoresHeightMap(x, y).g
            OriginalMapColor(x, y).b = ColoresHeightMap(x, y).b
            Intensidad_Del_Terreno(x, y) = ColoresHeightMap(x, y).a
        Next y
    Next x
    
    LogDebug "LightMap cargado!: " + Path
    
    Engine_Landscape.Light_Update_Map = True
    
Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
        Compute_Mountain
    'If cron_tiempo = False Then
        cron_tiempo
        Light_Update_Map = True
        Light_Update_Sombras = True
        
        
    CalcularNormales
End Sub

Public Sub Engine_MapColorsIO_SaveMiniMap()

End Sub

