Attribute VB_Name = "Engine_VertexShader"
Option Explicit

Private Camera_Projection As D3DMATRIX
Private Camera_World As D3DMATRIX
Private Camera_View As D3DMATRIX
Private MatrixIdentity As D3DMATRIX

Public Sub Engine_VS_Init()
    Call D3DXMatrixOrthoOffCenterLH(Camera_Projection, 0, D3DWindow.BackBufferWidth, 0, D3DWindow.BackBufferHeight, -1, 1)
    D3DXMatrixIdentity Camera_Projection
    D3DXMatrixIdentity Camera_World
    D3DXMatrixIdentity Camera_View
End Sub

Public Sub Engine_VS_ApplyTransforms()
    Engine_VS_Init
    D3DDevice.SetTransform D3DTS_PROJECTION, Camera_Projection
    D3DDevice.SetTransform D3DTS_WORLD, Camera_World
    D3DDevice.SetTransform D3DTS_VIEW, Camera_View
End Sub
