Attribute VB_Name = "Engine_LightsTexture"
Option Explicit

' Luces
Private LightsSurface As Direct3DSurface8
Private LightsTexture As Direct3DTexture8

Private LightTBOX As Engine.Box_Vertex

Public Const LightTexture As Integer = 19222
Public Const LightBackbufferSize As Integer = 512

Private oldSurface As Direct3DSurface8

Private tmpDesc As D3DSURFACE_DESC

Public Function Engine_LightsTexture_Init() As Boolean
On Error GoTo errh
    Set oldSurface = D3DDevice.GetDepthStencilSurface

    oldSurface.GetDesc tmpDesc ' get the current depth stencil's surface description
    
    Set LightsSurface = D3DDevice.CreateDepthStencilSurface(LightBackbufferSize, LightBackbufferSize, tmpDesc.format, D3DMULTISAMPLE_NONE) ' use that description to create a new depth surface
    Set LightsTexture = D3DDevice.CreateTexture(LightBackbufferSize, LightBackbufferSize, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)

    Set LightsSurface = LightsTexture.GetSurfaceLevel(0)
    
    With LightTBOX
        .color0 = -1
        .Color1 = -1
        .Color2 = -1
        .color3 = -1
        .rhw0 = 1
        .rhw1 = 1
        .rhw2 = 1
        .rhw3 = 1
        
        ' Coordenadas de la primer textura
        .tu0 = 0
        .tv0 = 1
        
        .tu1 = 0
        .tv0 = 0
        
        .tu2 = 1
        .tv2 = 1
        
        .tu3 = 1
        .tv3 = 0
        
        ' Coordenadas de la segunda textura
        .tu01 = 0
        .tv01 = 1
        
        .tu11 = 0
        .tv01 = 0
        
        .tu21 = 1
        .tv21 = 1
        
        .tu31 = 1
        .tv31 = 0
        
        .y0 = LightBackbufferSize
        .X2 = LightBackbufferSize
        .Y2 = LightBackbufferSize
        .x3 = LightBackbufferSize
    End With
Exit Function
errh:
LogError "Engine_LightsTexture_Init: " & D3DX.GetErrorString(Err.Number)

End Function

Public Sub Engine_LightsTexture_Render()

On Error GoTo errh
    'Render de prueba.
    
    Dim s As Direct3DSurface8
    
    Set s = LightsTexture.GetSurfaceLevel(0)
    
    D3DDevice.SetRenderTarget s, LightsSurface, ByVal 0
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF6969FF, 1#, 0
    D3DDevice.BeginScene
    
    Engine_LightsTexture_RenderLights

    D3DDevice.EndScene
    D3DDevice.SetRenderTarget D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), oldSurface, ByVal 0
    
    RenderLights_Render = True
    
Exit Sub
errh:
LogError "No se puede hacer Engine_LightsTexture_Render"
End Sub

Public Sub Engine_LightsTexture_RenderLights()
    Call GetTexture(LightTexture)

    ' for luz in luces
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LightTBOX, TL_size
End Sub


