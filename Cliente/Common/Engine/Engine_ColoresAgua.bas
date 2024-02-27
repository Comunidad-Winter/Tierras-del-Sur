Attribute VB_Name = "Engine_ColoresAgua"
Option Explicit

Public ColoresAguaTexture As Direct3DTexture8

Private Priv_ImagenData(0 To 255, 0 To 255) As BGRACOLOR_DLL



Public Sub ColoresAguaInit()
    Set ColoresAguaTexture = D3DX.CreateTexture(D3DDevice, 256, 256, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
    Dim x As Long
    Dim y As Long
     
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            With AguaBoxes(x, y)
                .tu02 = x / 256
                .tu12 = .tu02
                .tu22 = (x + 1) / 256
                .tu32 = .tu22
                
                .tv02 = (y + 1) / 256
                .tv12 = y / 256
                .tv22 = .tv02
                .tv32 = .tv12
                
                .rhw0 = 1
                .rhw1 = 1
                .rhw2 = 1
                .rhw3 = 1
            End With
            
            Priv_ImagenData(x, y).b = 255
            Priv_ImagenData(x, y).r = 255
            Priv_ImagenData(x, y).g = 255
        Next y
    Next x
End Sub

Public Sub ColoresAgua_Redraw()
    Dim lockedRect As D3DLOCKED_RECT

    Dim x As Byte, y As Byte

    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
            With Tileset_Grh_Array(TileNumberWater(x, y))
                AguaBoxes(x, y).tu0 = .tu0
                AguaBoxes(x, y).tv0 = .tv0
                AguaBoxes(x, y).tu1 = .tu1
                AguaBoxes(x, y).tv1 = .tv1
                AguaBoxes(x, y).tu2 = .tu2
                AguaBoxes(x, y).tv2 = .tv2
                AguaBoxes(x, y).tu3 = .tu3
                AguaBoxes(x, y).tv3 = .tv3
            End With
            If hMapData(x, y).hs(0) < mapinfo.agua_profundidad Or hMapData(x, y).hs(1) < mapinfo.agua_profundidad Or hMapData(x, y).hs(2) < mapinfo.agua_profundidad Or hMapData(x, y).hs(3) < mapinfo.agua_profundidad Then
                Priv_ImagenData(x, y).a = minl(-minl(hMapData(x, y).hs(0), minl(hMapData(x, y).hs(1), minl(hMapData(x, y).hs(2), hMapData(x, y).hs(3)))) * 4, 255)
                mapdata(x, y).is_water = 255
                Engine_Landscape_Water.AguaVisiblePosicion(x, y) = 1
            Else
                Priv_ImagenData(x, y).a = 0
                mapdata(x, y).is_water = 0
                Engine_Landscape_Water.AguaVisiblePosicion(x, y) = 0
            End If
        Next
    Next
    
    ColoresAguaTexture.LockRect 0, lockedRect, ByVal 0, 0

    DXCopyMemory ByVal lockedRect.pBits, Priv_ImagenData(0, 0), 262144 ' 256 * 256 * 4

    ColoresAguaTexture.UnlockRect 0
End Sub

