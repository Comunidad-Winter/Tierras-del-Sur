Attribute VB_Name = "Engine_Render2texture"
Option Explicit

Private OffScreenSurf As Direct3DSurface8
Public texture As Direct3DTexture8

Public ZBuffer As Direct3DSurface8

Public BackBufferSurf As Direct3DSurface8

Private Size As Single




    Private tmpSurface As Direct3DSurface8
    Private oldSurface As Direct3DSurface8
    Private tarSurface As Direct3DSurface8
    Private nResult As Long
    Private tmpDesc As D3DSURFACE_DESC
    Private TextureTarget As Direct3DTexture8
    Private pTextureSize As Long
    
    
    Public PuedeRenderToTexture As Boolean
    
Public Function StartRenderToTexture(Optional ByRef texture As Direct3DTexture8 = Nothing)



    If PuedeRenderToTexture Then
        If texture Is Nothing Then
            Set TextureTarget = D3DDevice.CreateTexture(pTextureSize, pTextureSize, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
        Else
            Set TextureTarget = texture
        End If
        
        Dim s As Direct3DSurface8
        
        Set s = TextureTarget.GetSurfaceLevel(0)
    
        D3DDevice.SetRenderTarget s, tmpSurface, ByVal 0
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0&, 1!, 0
        D3DDevice.BeginScene
    End If
End Function

Public Function EndRenderToTexture()
    If PuedeRenderToTexture Then
        D3DDevice.EndScene
        D3DDevice.SetRenderTarget D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), oldSurface, ByVal 0
        
    End If
End Function

Public Function InitRenderToTexture(Optional Siz As Single = 512) As Boolean

'todo ARREGKLAR ESTO
On Error GoTo errh
    Set oldSurface = D3DDevice.GetDepthStencilSurface
    pTextureSize = Siz
    oldSurface.GetDesc tmpDesc ' get the current depth stencil's surface description
    Set tmpSurface = D3DDevice.CreateDepthStencilSurface(Siz, Siz, tmpDesc.format, D3DMULTISAMPLE_NONE) ' use that description to create a new depth surface
    Set TextureTarget = D3DDevice.CreateTexture(pTextureSize, pTextureSize, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
    
    Dim s As Direct3DSurface8
    
    Set s = TextureTarget.GetSurfaceLevel(0)
    
    'Render de prueba.
    D3DDevice.SetRenderTarget s, tmpSurface, ByVal 0
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0&, 1#, 0
    D3DDevice.BeginScene
    Engine.Draw_FilledBox 0, 0, 512, 512, 0, &HFFFFFFF
    D3DDevice.EndScene
    D3DDevice.SetRenderTarget D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), oldSurface, ByVal 0
    
    InitRenderToTexture = True
    PuedeRenderToTexture = True
    act_caps.Cando_RenderSurface = True
Exit Function
errh:
LogError "No se puede hacer RenderToTarget"
PuedeRenderToTexture = False
InitRenderToTexture = False
End Function


Public Sub RenderRenderToTexture(ByVal dest_x!, ByVal dest_y!, ByVal Color As Long)
''*********************************************
''Author: menduz
''*********************************************
'    Dim dest_x2!, dest_y2!
'
'    dest_y2 = dest_y + Size
'    dest_x2 = dest_x + Size
'
'    Dim tBox As Box_Vertex
'
'    With tBox
'        .x0 = dest_x
'        .y0 = dest_y2
'        .x1 = dest_x
'        .y1 = dest_y
'        .x2 = dest_x2
'        .y2 = dest_y2
'        .x3 = dest_x2
'        .y3 = dest_y
'        .color0 = color
'        .color1 = color
'        .color2 = color
'        .color3 = color
'        .tu0 = 0
'        .tv0 = 1
'        .tu1 = 0
'        .tv1 = 0
'        .tu2 = 1
'        .tv2 = 1
'        .tu3 = 1
'        .tv3 = 0
'    End With
'    D3DDevice.SetVertexShader FVF
'    D3DDevice.SetTexture 0, Texture
'    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
End Sub

Sub renderheads()
Dim c&, x&, y&, T&

    'Marce On local error resume next
    c = &HFF000000
    'Engine.Draw_FilledBox 0, 0, 512, 512, 0, c
    For x = 0 To 31
    For y = 0 To 31
    T = T + 1
    c = c Xor &H111111
    'Engine.Draw_FilledBox x * 16, y * 16, 16, 16, 0, c
    Grh_Render HeadData(T).Head(SOUTH).GrhIndex, x * 16, y * 16, &HFF7F7F7F

    Next y
    Next x
    'Marce On local error goto 0
End Sub

Public Sub rendergun(ByVal gun As Integer)
Dim c&, x&, y&, T&, f&, ff&, jj%
Dim g As GrhData
Dim XX!, YY!



   'On Error GoTo rendergun_Error

    'Marce On local error resume next
    c = &HFF000000
    'Engine.Draw_FilledBox 0, 0, 512, 512, 0, c
    For y = 0 To 7
        T = 0
        For x = 0 To 5
        
            ff = (y Mod 4) + 1
            f = 0
            If ff = E_Heading.SOUTH Then
                ff = E_Heading.NORTH
                f = 1
            End If
            If ff = E_Heading.NORTH And f = 0 Then
                ff = E_Heading.SOUTH
                f = 1
            End If
            If ff = E_Heading.WEST And f = 0 Then
                ff = E_Heading.EAST
                f = 1
            End If
            If ff = E_Heading.EAST And f = 0 Then
                ff = E_Heading.WEST
                f = 1
            End If
    
            f = 0
            g = GrhData(WeaponAnimData(gun).WeaponWalk(ff).GrhIndex)
            
            f = g.NumFrames
            
            XX = x * 64 + 16
            YY = y * 64 + 32
            T = T + 1

            If (y + x) Mod 3 Then
                c = c Xor &H222222
            End If
            
            jj = g.frames(T Mod (f + 1))
            g = GrhData(jj)
            
            'If g.TileWidth <> 1 Then _
           '     XX = XX - Int(g.TileWidth * 16) + 16
            'If g.TileHeight <> 1 Then _
           '     YY = YY - Int(g.TileHeight * 32) + 32
            
            'If x <= 5 Then 'T <= f And
                Grh_Render jj, XX, YY, &HFF7F7F7F
            'End If
        
        
        Next x
    Next y
    'Marce 'Marce On local error goto 0

   'Marce On error goto 0
   Exit Sub
'
'rendergun_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rendergun of Módulo Engine_Render2texture"
End Sub

