Attribute VB_Name = "Engine_Camera"
Option Explicit

Private UsedProjectionMatrix    As D3DMATRIX ' c0 c1 c2 c3
Private ScreenSize              As D3DVECTOR4 ' c4
Private HalfScreenSize          As D3DVECTOR4 ' c5
Private CameraPos               As D3DVECTOR4 ' c6
Private MultiplicarPorDos       As D3DVECTOR4 ' c7
Private MultiplicarPorYInv      As D3DVECTOR4 ' c8
Private Parallax                As D3DVECTOR4 ' c9

Public ProjectionMatrix As D3DMATRIX
Public ViewMatrix As D3DMATRIX

Public Sub InitCamera()
    ScreenSize.x = 1 / D3DWindow.BackBufferWidth
    ScreenSize.y = -1 / D3DWindow.BackBufferHeight
    ScreenSize.z = 1 / 100
    ScreenSize.w = 1
    
    D3DXMATH_VECTOR4.D3DXVec4Scale HalfScreenSize, ScreenSize, 0.5
    
    MultiplicarPorDos.x = 2
    MultiplicarPorDos.y = 2
    MultiplicarPorDos.z = 2
    MultiplicarPorDos.w = 2
    
    D3DXMATH_VECTOR4.D3DXVec4Scale Parallax, HalfScreenSize, 1
    
    D3DXMatrixOrthoLH ProjectionMatrix, D3DWindow.BackBufferWidth, -D3DWindow.BackBufferHeight, -1000, 1000
    
    Call SetCameraPixelPos(0, 0)
End Sub

Public Sub SetParallaxIntensity(ByVal Intensity As Single)
    Parallax.w = Intensity
    
    SetTransforms
End Sub

Public Sub SetCameraPixelPos(ByVal PixelPosX As Single, ByVal PixelPosY As Single)
    CameraPos.x = -PixelPosX
    CameraPos.y = -PixelPosY
    
    D3DXMatrixIdentity ViewMatrix
    D3DXMatrixTranslation ViewMatrix, -1 + PixelPosX / D3DWindow.BackBufferWidth * 2, 1 - PixelPosY / D3DWindow.BackBufferHeight * 2, 0
    
    SetTransforms
End Sub

Public Sub SetTransforms()
    ' Proyeccion * Vista
    
    MultiplicarPorYInv.x = 1
    MultiplicarPorYInv.y = -1
    MultiplicarPorYInv.z = 1
    MultiplicarPorYInv.w = 1
    
    MultiplicarPorDos.x = 0.5
    MultiplicarPorDos.y = -0.5
    
    D3DXMatrixMultiply UsedProjectionMatrix, ProjectionMatrix, ViewMatrix
    D3DXMatrixTranspose UsedProjectionMatrix, UsedProjectionMatrix
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    ' Le mandamos cumbia al Device
    D3DDevice.SetVertexShaderConstant 0, UsedProjectionMatrix, 4
    
    D3DDevice.SetVertexShaderConstant 4, ScreenSize, 1
    D3DDevice.SetVertexShaderConstant 5, HalfScreenSize, 1
    D3DDevice.SetVertexShaderConstant 6, CameraPos, 1
    D3DDevice.SetVertexShaderConstant 7, MultiplicarPorDos, 1
    D3DDevice.SetVertexShaderConstant 8, MultiplicarPorYInv, 1
    D3DDevice.SetVertexShaderConstant 9, Parallax, 1
End Sub
