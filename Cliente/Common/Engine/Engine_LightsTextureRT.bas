Attribute VB_Name = "Engine_LightsTexture"
Option Explicit

' Luces
Public LightsTextureHorizontal As Direct3DTexture8
Public LightsTextureVertical As Direct3DTexture8
Public LightsTextureSombras As Direct3DTexture8
Public LightsTextureLightMap As Direct3DTexture8

Private LightTBOX As Engine.Box_Vertex

Private FrontBufferTBOX As Engine.Box_Vertex
Private TBOX_Montanias As Engine.Box_Vertex



Private RenderTargetHorizontal As New Engine_RenderTarget
Private RenderTargetVertical As New Engine_RenderTarget
Private RenderTargetSombras As New Engine_RenderTarget
Private RenderTargetLightMap As New Engine_RenderTarget


Private LucesEnPantalla(LucesEnPantallaMax) As Integer
Private LucesEnPantallaCount As Integer

Public PosicionLucesFactorX As Single
Public PosicionLucesFactorY As Single

Public AnguloLuz As Single



' Esta constante dice cuanto tiene que crecer el ancho de las luces
Public Const LightRadioAug As Single = 1.5 * 32

Public NormalesMapaNececitanActualizar As Boolean

Private ShadowTBOX(6) As TLVERTEX

Private Type ShadeableObject
    x As Long
    y As Long
    Radius As Long
End Type

Private Const MAX_SHADEABLE_OBJECTS = 100

Private ShadeableObjects(MAX_SHADEABLE_OBJECTS) As ShadeableObject
Private ShadeableObjectsCount As Long

Public Sub Engine_Shadows_Clear()
    ShadeableObjectsCount = 0
End Sub

Public Sub Engine_Shadows_Add(ByVal PixelPosX As Long, ByVal PixelPosY As Long, Optional ByVal Radius As Long = 16)
    With ShadeableObjects(ShadeableObjectsCount)
        .x = PixelPosX
        .y = PixelPosY
        .Radius = Radius
    End With
    
    ShadeableObjectsCount = (ShadeableObjectsCount Mod MAX_SHADEABLE_OBJECTS) + 1
End Sub



Public Sub Engine_LightsTexture_Clear()
    LucesEnPantallaCount = 0
End Sub

Public Sub Engine_LightsTexture_Push(ByVal i As Integer)
    LucesEnPantalla(LucesEnPantallaCount) = i
    LucesEnPantallaCount = LucesEnPantallaCount + 1 Mod LucesEnPantallaMax
End Sub

Public Function Engine_LightsTexture_Init() As Boolean
On Error GoTo errh
    
    RenderTargetHorizontal.rtCreate LightBackbufferSize, LightBackbufferSize
    Set LightsTextureHorizontal = RenderTargetHorizontal.objTexture
    
    RenderTargetVertical.rtCreate LightBackbufferSize, LightBackbufferSize
    Set LightsTextureVertical = RenderTargetVertical.objTexture
    
    RenderTargetSombras.rtCreate LightBackbufferSize, LightBackbufferSize
    Set LightsTextureSombras = RenderTargetSombras.objTexture
    
    RenderTargetLightMap.rtCreate LightBackbufferSize, LightBackbufferSize
    Set LightsTextureLightMap = RenderTargetLightMap.objTexture
    
    PosicionLucesFactorX = LightBackbufferSize / Engine.D3DWindow.BackBufferWidth
    PosicionLucesFactorY = LightBackbufferSize / Engine.D3DWindow.BackBufferHeight
    
    Dim i As Integer
    For i = 0 To 5
        With ShadowTBOX(i)
            .Color = -1
            .rhw = 1
        End With
    Next i
'    With ShadowTBOX(0)
'        .color0 = -1
'        .Color1 = -1
'        .Color2 = -1
'        .color3 = -1
'        .rhw0 = 1
'        .rhw1 = 1
'        .rhw2 = 1
'        .rhw3 = 1
'
'        ' Coordenadas de la primer textura
'        .tu0 = 0
'        .tv0 = 1
'
'        .tu1 = 0
'        .tv1 = 0
'
'        .tu2 = 1
'        .tv2 = 1
'
'        .tu3 = 1
'        .tv3 = 0
'
'        ' Coordenadas de la segunda textura
'        .tu01 = 0
'        .tv01 = 1
'
'        .tu11 = 0
'        .tv11 = 0
'
'        .tu21 = 1
'        .tv21 = 1
'
'        .tu31 = 1
'        .tv31 = 0
'
'        .y0 = LightBackbufferSize
'        .x2 = LightBackbufferSize
'        .y2 = LightBackbufferSize
'        .x3 = LightBackbufferSize
'    End With
    
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
        .tv1 = 0
        
        .tu2 = 1
        .tv2 = 1
        
        .tu3 = 1
        .tv3 = 0
        
        ' Coordenadas de la segunda textura
        .tu01 = 0
        .tv01 = 1
        
        .tu11 = 0
        .tv11 = 0
        
        .tu21 = 1
        .tv21 = 1
        
        .tu31 = 1
        .tv31 = 0
        
        .y0 = LightBackbufferSize
        .x2 = LightBackbufferSize
        .y2 = LightBackbufferSize
        .x3 = LightBackbufferSize
    End With
    
    With FrontBufferTBOX
        .color0 = &HFFFFFFFF
        .Color1 = &HFFFFFFFF
        .Color2 = &H7FFFFFFF
        .color3 = &H7FFFFFFF
        .rhw0 = 1
        .rhw1 = 1
        .rhw2 = 1
        .rhw3 = 1
        
        ' Coordenadas de la primer textura
        .tu0 = 0
        .tv0 = 1
        
        .tu1 = 0
        .tv1 = 0
        
        .tu2 = 1
        .tv2 = 1
        
        .tu3 = 1
        .tv3 = 0
        
        ' Coordenadas de la segunda textura
        .tu01 = 0
        .tv01 = 1
        
        .tu11 = 0
        .tv11 = 0
        
        .tu21 = 1
        .tv21 = 1
        
        .tu31 = 1
        .tv31 = 0
        
        .y0 = Engine.D3DWindow.BackBufferHeight
        .x2 = Engine.D3DWindow.BackBufferWidth
        .y2 = Engine.D3DWindow.BackBufferHeight
        .x3 = Engine.D3DWindow.BackBufferWidth
    End With
    
    TBOX_Montanias = FrontBufferTBOX
    
     With TBOX_Montanias
        .y0 = LightBackbufferSize
        .x2 = LightBackbufferSize
        .y2 = LightBackbufferSize
        .x3 = LightBackbufferSize
        
        .color0 = -1
        .Color1 = -1
        .Color2 = -1
        .color3 = -1
    End With
    
    
    If Engine_General.NoUsarSombras = False Then
        Engine_NormalesMontanias.NormalMontaniasInit
    End If
Exit Function
errh:
LogError "Engine_LightsTexture_Init: " & D3DX.GetErrorString(Err.Number)

End Function

Public Sub Engine_LightsTexture_Render()

#If medir = 1 Then
    If mostrarTiempos Then
        TiempoLucesLightmaps = GetElapsedTimeME
    End If
#End If



    If NormalesMapaNececitanActualizar And Engine_General.NoUsarSombras = False Then
        Engine_NormalesMontanias.NormalMontanias_Redraw
        NormalesMapaNececitanActualizar = False
    End If

    Dim timer As Single
    timer = (HoraDelDia / 24#) * Pi2 'timer + Engine.timerElapsedTime * 0.008
    
    Dim colorClear As Long
    
    If mapinfo.ColorPropio = False Then
        colorClear = D3DColorMake(Sin(timer) / 2 + 0.5, 0.5 + Cos(timer) / 2, 0.8, 0.5)
    Else
        colorClear = &H6969FF
    End If
    
    
    'Render de prueba.
    Call RenderTargetHorizontal.rtAquire
    Call RenderTargetHorizontal.rtEnable(True)
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, colorClear, 0, 0
    
    Engine_LightsTexture_RenderLights 0


    Call RenderTargetHorizontal.rtEnable(False)
    
    Call RenderTargetVertical.rtAquire
    Call RenderTargetVertical.rtEnable(True)
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, colorClear, 0, 0
    
    Engine_LightsTexture_RenderLights 1

    Call RenderTargetVertical.rtEnable(False)
    
#If medir = 1 Then
    If mostrarTiempos Then
        TiempoLucesLightmaps = GetElapsedTimeME - TiempoLucesLightmaps
    End If
#End If
    
End Sub

Public Sub Engine_LightsTexture_RenderLights(ByVal Vertical As Byte)
    
    
    Dim i As Integer
    
    
    
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim Color As RGBCOLOR
    Dim range As Byte
    Dim brillo As Byte
    Dim tipo As Integer
    Dim map_x As Byte
    Dim map_y As Byte
    Dim PixelPosX As Single
    Dim PixelPosY As Single
    Dim RadioTransformadoX As Single
    Dim RadioTransformadoY As Single
    Dim pixel_pos_x As Integer
    Dim pixel_pos_y As Integer
    
    Dim ViejoFiltroEscalado As Long
    
    ViejoFiltroEscalado = D3DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER)
        
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        
        
    If Engine_General.NoUsarSombras = False Then
        D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Engine_NormalesMontanias.NormalMontaniasTexture

        With TBOX_Montanias
            .tu0 = (screenminX - offset_screen.x / 32) / 255
            .tv0 = (Engine_Map.screenmaxY - offset_screen.y / 32) / 255
            
            .tu1 = TBOX_Montanias.tu0
            .tv1 = (Engine_Map.screenminY - offset_screen.y / 32) / 255
            
            .tu2 = (screenmaxX - offset_screen.x / 32) / 255
            .tv2 = TBOX_Montanias.tv0
            
            .tu3 = TBOX_Montanias.tu2
            .tv3 = TBOX_Montanias.tv1
                
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, .x0, TL_size
        End With
    End If
    
    If LucesEnPantallaCount = 0 Then
        D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, ViejoFiltroEscalado
        Exit Sub
    End If
    Dim LightTexture As Direct3DTexture8
    
    
    If Vertical = 0 Then
        Set LightTexture = PeekTexture(LightTextureHorizontal)
    Else
        Set LightTexture = PeekTexture(LightTextureVertical)
    End If
    
    'Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse LightTexture

    For i = 0 To LucesEnPantallaCount - 1
        Dim half_y&, half_x&, s&, c&
        If DLL_Luces.Get_Light_Ext(LucesEnPantalla(i), map_x, map_y, r, g, b, range, brillo, tipo, 1, 1, pixel_pos_x, pixel_pos_y) Then
            ' Si vos tenes una posicion en el mapa en pixeles (ej: X=13922 Y=10122)
            ' y la queres mostrar en la pantalla tenes que sumarle este vector offset_map_part
            
            PixelPosX = (CSng(pixel_pos_x) + offset_map_part.x) * PosicionLucesFactorX
            PixelPosY = (CSng(pixel_pos_y) + offset_map_part.y + 32) * PosicionLucesFactorY
            
            RadioTransformadoX = CSng(range) * LightRadioAug * PosicionLucesFactorX
            RadioTransformadoY = CSng(range) * LightRadioAug * PosicionLucesFactorY
 
            With LightTBOX
                .x0 = PixelPosX - RadioTransformadoX
                .x1 = .x0
                .x2 = PixelPosX + RadioTransformadoX
                .x3 = .x2
                
                .y0 = PixelPosY + RadioTransformadoY
                .y1 = PixelPosY - RadioTransformadoY
                .y2 = .y0
                .y3 = .y1
                
                If tipo And 1 Then ' la luz tiene brillo
                    .color0 = Not Alphas(brillo)
                Else
                    .color0 = -1
                End If
                
                .Color1 = .color0
                .Color2 = .Color1
                .color3 = .Color1
            End With
            
            Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse LightTexture

            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LightTBOX, TL_size
        
            If SombrasHQ Then
                Dim ShadowIndex As Integer
        
                For ShadowIndex = 0 To ShadeableObjectsCount - 1
                    If GenerateLightBox(ShadowTBOX, ShadeableObjects(ShadowIndex).x + offset_map_part.x, ShadeableObjects(ShadowIndex).y + offset_map_part.y + 32, _
                        CSng(pixel_pos_x) + offset_map_part.x, CSng(pixel_pos_y) + offset_map_part.y + 32, _
                        range * 32, ShadeableObjects(ShadowIndex).Radius) Then
                    
                        Dim vc As Integer
                        
                        For vc = 0 To 6
                            ShadowTBOX(vc).v.x = ShadowTBOX(vc).v.x * PosicionLucesFactorX
                            ShadowTBOX(vc).v.y = ShadowTBOX(vc).v.y * PosicionLucesFactorY
                        Next vc
                        
                        'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
                        
                        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(TexturaSombra)
                        
                        Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
                       ' Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SUBTRACT)
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 5, ShadowTBOX(0), TL_size
                        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
                    
                        
                    End If
                Next ShadowIndex
            End If
        End If
    Next i
    
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, ViejoFiltroEscalado
Exit Sub
    If Vertical Then Exit Sub
    
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        
    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SELECTARG1)
        
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse LightsTextureSombras
    
    LightTBOX.color0 = -1
    LightTBOX.Color1 = -1
    LightTBOX.Color2 = -1
    LightTBOX.color3 = -1
    
    LightTBOX.x0 = 0
    LightTBOX.x1 = LightTBOX.x0
    LightTBOX.x2 = LightBackbufferSize
    LightTBOX.x3 = LightTBOX.x2
    
    LightTBOX.y0 = LightBackbufferSize
    LightTBOX.y1 = 0
    LightTBOX.y2 = LightTBOX.y0
    LightTBOX.y3 = LightTBOX.y1
    
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LightTBOX, TL_size
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
    
    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        
    
End Sub






Public Sub Engine_LightsTexture_RenderShadows()
    Dim i As Integer
    
    
    
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim Color As RGBCOLOR
    Dim range As Byte
    Dim brillo As Byte
    Dim tipo As Integer
    Dim map_x As Byte
    Dim map_y As Byte
    Dim PixelPosX As Single
    Dim PixelPosY As Single
    Dim RadioTransformadoX As Single
    Dim RadioTransformadoY As Single
    Dim pixel_pos_x As Integer
    Dim pixel_pos_y As Integer
    

    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        
    If LucesEnPantallaCount = 0 Then
        Exit Sub
    End If
    
    If ShadeableObjectsCount = 0 Then
        Exit Sub
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing ' PeekTexture(5398)

Dim ViejoFiltroEscalado As Long
    
    ViejoFiltroEscalado = D3DDevice.GetTextureStageState(0, D3DTSS_MINFILTER)
        
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        
        
        Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_SUBTRACT)
        
        
        
        D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        


    Dim ShadowIndex As Integer
    
    For ShadowIndex = 0 To ShadeableObjectsCount - 1
        For i = 0 To LucesEnPantallaCount - 1
            Dim half_y&, half_x&, s&, c&
            If DLL_Luces.Get_Light_Ext(LucesEnPantalla(i), map_x, map_y, r, g, b, range, brillo, tipo, 1, 1, pixel_pos_x, pixel_pos_y) Then
                ' Si vos tenes una posicion en el mapa en pixeles (ej: X=13922 Y=10122)
                ' y la queres mostrar en la pantalla tenes que sumarle este vector offset_map_part
                
                If GenerateLightBox(ShadowTBOX, ShadeableObjects(ShadowIndex).x + offset_map_part.x, ShadeableObjects(ShadowIndex).y + offset_map_part.y + 32, _
                    CSng(pixel_pos_x) + offset_map_part.x, CSng(pixel_pos_y) + offset_map_part.y + 32, _
                    range * 32, ShadeableObjects(ShadowIndex).Radius) Then
                    Dim vc As Integer
                    
                    For vc = 0 To 6
                        ShadowTBOX(vc).v.x = ShadowTBOX(vc).v.x * RadioTransformadoX
                        ShadowTBOX(vc).v.y = ShadowTBOX(vc).v.y * RadioTransformadoY
                    Next vc
                    
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 5, ShadowTBOX(0), TL_size
                
                End If

            End If
        Next i
    Next ShadowIndex
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, ViejoFiltroEscalado
End Sub

Public Function GenerateLightBox(Vertices() As TLVERTEX, ByVal x As Long, ByVal y As Long, ByVal LightX As Long, ByVal LightY As Long, ByVal LightRange As Long, ByVal ObjectRadius As Long) As Boolean
' todos los parametros de esta funcion son en pixels
    Dim Angulo As Single
    Dim Distancia As Single
                
    Distancia = Sqr((x - LightX) * (x - LightX) + (y - LightY) * (y - LightY))

    Dim AltoObjeto As Single
    
    AltoObjeto = 128

    If Distancia = 0 Or Distancia > LightRange Then Exit Function

    Angulo = ATAN_2(LightX - x, LightY - y)


    Dim Intensidad As Long
    
    AltoObjeto = AltoObjeto * Distancia / LightRange
    
    Dim DistanciaSombra As Single
    
    DistanciaSombra = Distancia / LightRange * 2
    
    Intensidad = 200 - (Distancia / LightRange * 200)

    Vertices(1).Color = Alphas(Intensidad)
    Vertices(2).Color = Alphas(Intensidad)
    Vertices(0).Color = Alphas(Intensidad)
    
    
    Vertices(3).Color = 0
    Vertices(4).Color = 0
    Vertices(5).Color = 0
    
    Vertices(0).rhw = 1
    Vertices(1).rhw = 1
    Vertices(2).rhw = 1
    Vertices(3).rhw = 1
    Vertices(4).rhw = 1
    Vertices(5).rhw = 1
    Vertices(6).rhw = 1
    
    Vertices(2).v.y = y + ObjectRadius / 2 * Cos(Angulo)
    Vertices(1).v.y = y - ObjectRadius / 2 * Cos(Angulo)

    Vertices(2).v.x = x - ObjectRadius / 2 * Sin(Angulo)
    Vertices(1).v.x = x + ObjectRadius / 2 * Sin(Angulo)

    Vertices(5).v.x = Vertices(1).v.x - (LightX - Vertices(1).v.x) * DistanciaSombra
    Vertices(5).v.y = Vertices(1).v.y - (LightY - Vertices(1).v.y) * DistanciaSombra
    Vertices(3).v.x = Vertices(2).v.x - (LightX - Vertices(2).v.x) * DistanciaSombra
    Vertices(3).v.y = Vertices(2).v.y - (LightY - Vertices(2).v.y) * DistanciaSombra

    Vertices(0).v.x = x + AltoObjeto * Sin(Angulo - pi / 2)
    Vertices(0).v.y = y - AltoObjeto * Cos(Angulo - pi / 2)
    
    Vertices(4).v.x = x + AltoObjeto * Sin(Angulo - pi / 2) * 2
    Vertices(4).v.y = y - AltoObjeto * Cos(Angulo - pi / 2) * 2

    Vertices(0).tu = 0.5
    Vertices(0).tv = 0.125

    Vertices(1).tu = 1
    Vertices(1).tv = 0.75


    Vertices(2).tu = 0
    Vertices(2).tv = 0.75



    Vertices(3).tu = 0
    Vertices(3).tv = 0


    Vertices(4).tu = 0.5
    Vertices(4).tv = 0
    
    
    Vertices(5).tu = 1
    Vertices(5).tv = 0

    Vertices(6) = Vertices(1)
    
    GenerateLightBox = True
End Function






Public Sub Engine_LightsTexture_RenderBackbuffer()
Exit Sub
NormalMontanias_Redraw
D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse NormalMontaniasTexture 'LightsTextureHorizontal 'LightsTextureSombras
FrontBufferTBOX.color0 = -1
FrontBufferTBOX.Color1 = -1
FrontBufferTBOX.Color2 = -1
FrontBufferTBOX.color3 = -1
    ' for luz in luces
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, FrontBufferTBOX, TL_size
D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
End Sub


