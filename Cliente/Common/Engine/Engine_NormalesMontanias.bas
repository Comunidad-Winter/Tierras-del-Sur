Attribute VB_Name = "Engine_NormalesMontanias"
Option Explicit

Public NormalMontaniasTexture As Direct3DTexture8

Private Priv_ImagenData(0 To 255, 0 To 255) As BGRACOLOR_DLL
Private Priv_Mat_transformacion As D3DMATRIX


Public Sub NormalMontaniasInit()
    Set NormalMontaniasTexture = D3DX.CreateTexture(D3DDevice, 256, 256, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
    
    D3DXMATH_MATRIX.D3DXMatrixIdentity Priv_Mat_transformacion
End Sub

Public Sub NormalMontanias_Redraw()
    Dim lockedRect As D3DLOCKED_RECT
    
    Dim vec As D3DVECTOR
    
    D3DXMATH_MATRIX.D3DXMatrixRotationZ Priv_Mat_transformacion, -(HoraDelDia / 24) * Pi2

    Dim x As Byte, y As Byte

    For x = SV_Constantes.X_MINIMO_VISIBLE To SV_Constantes.X_MAXIMO_VISIBLE
        For y = SV_Constantes.Y_MINIMO_VISIBLE To SV_Constantes.Y_MAXIMO_VISIBLE
        
            If hMapData(x, y).h > 0 Then
                D3DXMATH_VECTOR3.D3DXVec3TransformNormal vec, NormalData(x, y), Priv_Mat_transformacion
            
                Priv_ImagenData(x, y).b = CByte(vec.y * 127 + 127)
                Priv_ImagenData(x, y).r = CByte(vec.x * 127 + 127)
                Priv_ImagenData(x, y).g = CByte(vec.z * 127 + 127)
                Priv_ImagenData(x, y).a = minl(hMapData(x, y).h * 3, 255)
            End If
        Next
    Next
    
    NormalMontaniasTexture.LockRect 0, lockedRect, ByVal 0, 0

    DXCopyMemory ByVal lockedRect.pBits, Priv_ImagenData(0, 0), 262144 ' 256 * 256 * 4

    NormalMontaniasTexture.UnlockRect 0
End Sub
