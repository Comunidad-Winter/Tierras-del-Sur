Attribute VB_Name = "Engine"
'ARCHIVO COMPARTIDOOO!


''
' @require Engine_ErrorLOG.bas


'                  ____________________________________________
'                 /_____/  http://www.arduz.com.ar/ao/   \_____\
'                //            ____   ____   _    _ _____      \\
'               //       /\   |  __ \|  __ \| |  | |___  /      \\
'              //       /  \  | |__) | |  | | |  | |  / /        \\
'             //       / /\ \ |  _  /| |  | | |  | | / /   II     \\
'            //       / ____ \| | \ \| |__| | |__| |/ /__          \\
'           / \_____ /_/    \_\_|  \_\_____/ \____//_____|_________/ \
'           \________________________________________________________/
'           MOTOR GRÁFICO ESCRITO POR MENDUZ@NOICODER.COM



Option Explicit

'Public bRunning             As Boolean

Public DX                   As DirectX8
Public D3D                  As Direct3D8
Public D3DDevice            As Direct3DDevice8
Public D3DX                 As D3DX8
Public D3DWindow            As D3DPRESENT_PARAMETERS

Public nScreenBPP           As Long

Public FPS                  As Integer
Public puedo_deslimitar     As Boolean

Public FramesPerSecCounter As Long

Public timerElapsedTime     As Double

Public timerTicksPerFrame   As Double

Public engineBaseSpeed      As Single

Public Epsilon              As Single '= 0.0000001192093


Public Const TL_size        As Long = 28 + 8 + 8
Public Const BV_size        As Long = TL_size * 4
Public Const Part_size      As Long = 32

Public Const particleFVF    As Long = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_PSIZE Or D3DFVF_DIFFUSE)
Public Const FVF            As Long = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE)

Private font_list()         As D3DXFont

Public Type RGBCOLOR
    r As Byte
    g As Byte
    b As Byte
End Type

Public Type BGRCOLOR_DLL
    b As Byte
    g As Byte
    r As Byte
End Type


'Direcciones
Public Enum E_Heading
    None = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum


Public Type BGRACOLOR_DLL
    b As Byte
    g As Byte
    r As Byte
    a As Byte
End Type

Public Type TLVERTEX    'NO TOCAR POR NADA EN EL MUNDO
    v As D3DVECTOR
    rhw As Single       'NO TOCAR POR NADA EN EL MUNDO
    'normal As D3DVECTOR
    Color As Long       'NO TOCAR POR NADA EN EL MUNDO
    tu As Single        'NO TOCAR POR NADA EN EL MUNDO
    tv As Single        'NO TOCAR POR NADA EN EL MUNDO
    Offset As Long
    offset2 As Long
    Offset2_0 As Long
    offset2_2 As Long
End Type                'NO TOCAR POR NADA EN EL MUNDO


Public Type Box_Vertex
    x0 As Single
    y0 As Single
    Z0 As Single
    rhw0 As Single
    
    color0 As Long
    tu0 As Single
    tv0 As Single
    tu01 As Single
    tv01 As Single
    tu02 As Single
    tv02 As Single
    
    x1 As Single
    y1 As Single
    Z1 As Single
    rhw1 As Single
    Color1 As Long
    tu1 As Single
    tv1 As Single
    tu11 As Single
    tv11 As Single
    tu12 As Single
    tv12 As Single
    
    x2 As Single
    y2 As Single
    z2 As Single
    rhw2 As Single
    Color2 As Long
    tu2 As Single
    tv2 As Single
    tu21 As Single
    tv21 As Single
    tu22 As Single
    tv22 As Single
    
    x3 As Single
    y3 As Single
    Z3 As Single
    rhw3 As Single
    color3 As Long
    tu3 As Single
    tv3 As Single
    tu31 As Single
    tv31 As Single
    tu32 As Single
    tv32 As Single
End Type


Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private font_count As Long
Private font_last As Long

Private lFrameTimer As Long

' Tamaño de la pantalla (en realidad el buffer de dibujo)
Public MainViewWidth As Integer
Public MainViewHeight As Integer

' Tiles que se muestran en pantalla
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

' La mitad del tamaño en tiles de la pantalla
Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'TODO Esto no deberia estar aca
Public MouseTileX As Integer
Public MouseTileY As Integer

Private limit_fps As Boolean



'#########################################################################################################

'#########################################################################################################


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)




Public Type sRECT
    left As Single
    right As Single
    top As Single
    bottom As Single
End Type

Public offset_screen As D3DVECTOR

Public offset_map As D3DVECTOR2

Public offset_map_part As mzVECTOR2

Public offset_mapO As D3DVECTOR2








Public ZooMlevel!

Public new_text As Boolean


Public Type CAPABILITIES
    Filter_Bilinear As Boolean
    Filter_Trilinear As Boolean
    Filter_Anisotropic As Boolean
    Filter_GaussianCubic As Boolean
    Filetr_FlatCubic As Boolean

    CanDo_MultiTexture As Boolean
    CanDo_CubeMapping As Boolean
    CanDo_Dot3 As Boolean
    CanDo_VolumeTexture As Boolean
    CanDo_ProjectedTexture As Boolean
    CanDo_TextureMipMapping As Boolean
    CanDo_PureDevice As Boolean
    CanDo_PointSprite As Boolean

    Cando_RenderSurface As Boolean
    CandDo_3StagesTextureBlending As Boolean

    Cando_PixelShader As Boolean
    Cando_VertexShader As Boolean

    CanDoTableFog        As Boolean
    CanDoVertexFog       As Boolean
    CanDoWFog            As Boolean

    TandL_Device As Boolean
    CanDo_BumpMapping As Boolean

    Wbuffer_OK As Boolean
    Max_ActiveLights As Long
    Max_TextureStages As Long
    Max_AnisotropY As Long

    Pixel_ShaderVERSIOn As String

    Vertex_ShaderVERSION As String
    
    pxs_min As Long
    pxs_max As Long
End Type

Public act_caps As CAPABILITIES

Private tVerts(3) As TLVERTEX
Public tBox As Box_Vertex

Public copy_tile_now As ConstantesRecacheo

Public Cachear_Tiles As Boolean

Public WeatherFogX1 As Single       'Fog 1 position
Public WeatherFogY1 As Single       'Fog 1 position
Public WeatherFogX2 As Single       'Fog 2 position
Public WeatherFogY2 As Single       'Fog 2 position
Public WeatherDoFog As Byte         'Are we using fog? >1 = Yes, 0 = No
Public WeatherFogCount As Byte      'How many fog effects there are
Public Weatherfogalpha As Byte
Public Weatherfogalphau As Byte

Public Sandstorm_X1 As Single       'Fog 1 position
Public Sandstorm_Y1 As Single       'Fog 1 position
Public Sandstorm_X2 As Single       'Fog 2 position
Public Sandstorm_Y2 As Single       'Fog 2 position
Public Sandstorm_do As Byte         'Are we using fog? >1 = Yes, 0 = No
Public Sandstorm_Count As Byte      'How many fog effects there are

Public Lightbeam_a1 As Byte       'Fog 1 position
Public Lightbeam_a2 As Byte       'Fog 1 position
Public Lightbeam_a3 As Byte       'Fog 2 position
Public Lightbeam_Y2 As Single       'Fog 2 position
Public Lightbeam_do As Byte         'Are we using fog? >1 = Yes, 0 = No
Public Lightbeam_Count As Byte      'How many fog effects there are

Public Render_Radio_Luz As Byte      'How many fog effects there are

Private zNFrames      As Long               'No of Frames played since Last Reset

Public actual_blend_mode As Long

Public Enum render_set_states
    rsADD = 1
    rsONE = 2
End Enum

Public RENDERCRC As Long

Public FPS_LIMITER As clsPerformanceTimer
Public FRAME_TIMER As clsPerformanceTimer

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Private ultimo_error As Long
Private ultimo_error_desc As String


Public Enum ConstantesRecacheo
    mzCacheado = 0
    mzCachearEnMovimiento = 1
    mzCachear = 2
End Enum

Public pIB As Direct3DIndexBuffer8
Public IndexBufferEnabled As Boolean
Public Const INDEX_BUFFER_SIZE As Long = 4000
Public StaticIndexBuffer(INDEX_BUFFER_SIZE * 6) As Integer

Public Fps_Label As clsGUIText

Public Engine_Escene_Abierta As Boolean

Private pVertexShader As Long

Public IndiceSeteado As Boolean

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    Loops As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sx As Integer ' Posicion X donde se encuentra en la Imagen
    sy As Integer ' Posicion Y donde se encuentra en la Imagen
    
    filenum As Long ' Archivo de Imagen.
    
    pixelWidth As Integer ' Pixeles de Ancho
    pixelHeight As Integer ' Pixeles de alto
    
    NumFrames As Integer ' 1 Si es un gráfico simple, más de 1 si es una animación
    frames() As Long ' De 1 a cantidad de frames
    
    Speed As Single   ' Ahora es Single. Pero creo que no tiene sentido.
 
 
    offsetX As Integer ' Corrimiento en pixeles en X que aplica a todos las Capas
    offsetY As Integer ' Corrimient en pixeles en Y que aplica a todas las Capas
    
 
 
    tu(3) As Single
    tv(3) As Single
    hardcor As Byte
        

    
    SombrasSize As Byte  ' Tamaño de la sombra que genera el gráfico
    SombraOffsetX As Integer ' X e Y establecen el centro del gráfico que genera la luz
    SombraOffsetY As Integer

    'Informacion que utiliza el MapEditor
    #If esMe = 1 Then
        ' Tiene sentido si esta en la Capa 1 o 2.
        EfectoPisada As Integer
        
        id As String ' Identificador unico del grafico
        nombreGrafico As String ' Nombre que identifica (no univocamente) al grafico
        perteneceAunaAnimacion As Boolean ' ¿Es parte de una animación?
        esInsertableEnMapa As Boolean ' Se puede insertar en el mapa o es un npc, objeto?
        Capa(1 To CANTIDAD_CAPAS) As Boolean ' Capa en la cual se puede insertar
    #End If
    
End Type

Public GrhData() As GrhData 'Guarda todos los grh
Public grhCount As Long

Public prgRun As Boolean

Public PS_Handle As Long
Public PS_Constants(3, 2) As Single

Public PixelShaderBump As Long
'Public Const ShaderSombra As String = _
'"ps.1.0;                                " & vbNewLine & _
'"tex t0;                                " & vbNewLine & _
'"tex t1;                                " & vbNewLine & _
'"dp3_sat r0, t1_bx2, c0_bx2; // DOT3    " & vbNewLine & _
'"mov r0.a, t0;                          " & vbNewLine & _
'"mul r0, r0, t0;                        " & vbNewLine & _
'"mul r0, r0, v0;                        "

'Public Const ShaderSombra As String = _
'"ps.1.0;                    " & vbNewLine & _
'"tex t0;                    " & vbNewLine & _
'"tex t1;                    " & vbNewLine & _
'"tex t2;                    " & vbNewLine & _
'"mov r1, t2;                " & vbNewLine & _
'"dp3 r0, t1_bx2, r1_bx2;" & vbNewLine & _
'"add r0.rgb, r0, v0;        " & vbNewLine & _
'"+mov r0.a, t0;             " & vbNewLine & _
'"mul r0, r0, t0;            "

'
'Public Const ShaderSombra As String = _
'"ps.1.0;" & vbNewLine & _
'"tex t0;" & vbNewLine & _
'"tex t1;" & vbNewLine & _
'"tex t2;" & vbNewLine & _
'"mov r1, t2;" & vbNewLine & _
'"dp3_sat r0, t1_bx2, r1_bx2;" & vbNewLine & _
'"mul_x2 r0, r0, r1.a;" & vbNewLine & _
'"mul r0, v0, r0;" & vbNewLine & _
'"add r0.rgb, r0, c1;" & vbNewLine & _
'"+mov r0.a, t0;" & vbNewLine & _
'"mul r0, r0, t0;"

' El ultimo del codigo
Public Const ShaderSombra As String = _
"ps.1.0;" & vbNewLine & _
"tex t0;" & vbNewLine & _
"tex t1;" & vbNewLine & _
"tex t2;" & vbNewLine & _
"mov r1, t2;" & vbNewLine & _
"dp3_sat r0, t1_bx2, r1_bx2;" & vbNewLine & _
"mul_x2 r0, r0, r1.a;" & vbNewLine & _
"mad r0.rgb, v0, r0, c1;" & vbNewLine & _
"mul r0.rgb, r0, t0;" & vbNewLine & _
"+mov r0.a, t0;" & vbNewLine & _
"mul r0.a, r0.a, v0.a;"


Public SceneBegin As Long

' Convierte la posición x,y de la pantalla a un Tile
' viewPortX, viewPortY:  Coordenada de la pantalla
' tx, ty: OUTPUT. Tile X e Y donde se encuentra el mouse
' tamTileX, tamTileY: Tamaño, en pixeles, qye ocupa la representación de cada tile

Public Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByVal tamTileX As Byte, ByVal tamTileY As Byte, ByRef tx As Integer, ByRef ty As Integer)
' Obtengo el tile relativo al mouse TileX = MousePosX \ 32. TileY = MousePosY
' La posición relativa al mapa es RelativaX = UserPos.X - La cantidad de tiles que el personaje tiene a su derecha
' RelativaY = UserPos.Y - La cantidad de tiles que el personaje tiene para arriba.
' La posición final es RelativaX + TileX, RelativaY + RelativaY

' X = (MousePosX \ 32) + (UserPosX - CantidadTilesX)
' Y = (MousePosY \ 32) + (UserPosY - CantidadTilesX)
tx = UserPos.X + (viewPortX \ tamTileX) - HalfWindowTileWidth
ty = UserPos.Y + (viewPortY \ tamTileY) - HalfWindowTileHeight

If tx < X_MINIMO_VISIBLE Then
    tx = X_MINIMO_VISIBLE
ElseIf tx > X_MAXIMO_VISIBLE Then
    tx = X_MAXIMO_VISIBLE
End If

If ty < Y_MINIMO_VISIBLE Then
    ty = Y_MINIMO_VISIBLE
ElseIf ty > Y_MAXIMO_VISIBLE Then
    ty = Y_MAXIMO_VISIBLE
End If

End Sub

Public Sub IniciarIndexBuffers()
'On Error GoTo errh
IndexBufferEnabled = False

    Dim puntero As Long
    
    If ObjPtr(pIB) Then
        Set pIB = Nothing
    End If
    
    Set pIB = D3DDevice.CreateIndexBuffer(INDEX_BUFFER_SIZE * 12, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_MANAGED) '12=6(vertices por cuadrado) * 2(Bytes de un integer)
    Call pIB.Lock(0, 0, puntero, D3DLOCK_DISCARD)
    
    Dim j As Long, i As Long, k&
    j = 0
    i = 0
    For k = 0 To INDEX_BUFFER_SIZE - 1
        StaticIndexBuffer(j) = i
        j = j + 1
        StaticIndexBuffer(j) = i + 1
        j = j + 1
        StaticIndexBuffer(j) = i + 2
        j = j + 1
        StaticIndexBuffer(j) = i + 2
        j = j + 1
        StaticIndexBuffer(j) = i + 1
        j = j + 1
        StaticIndexBuffer(j) = i + 3
        j = j + 1
        i = i + 4
    Next k
    CopyMemory ByVal puntero, StaticIndexBuffer(0), INDEX_BUFFER_SIZE * 12
    
    'IniciarIndexBuffer INDEX_BUFFER_SIZE, puntero
    
    Call pIB.Unlock
    D3DDevice.SetIndices pIB, 0
    IndexBufferEnabled = True
    LogDebug "      IndexBuffer: True"
Exit Sub
errh:
    LogError "No se pudo iniciar el IndexBuffer"
    LogDebug "      IndexBuffer: False"
End Sub

Public Sub Set_Blend_Mode(Optional ByVal blend_mode As render_set_states)
    If ((actual_blend_mode And render_set_states.rsONE) <> (blend_mode And render_set_states.rsONE)) Then
        If (blend_mode And render_set_states.rsONE) Then
            Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)
        Else
            Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        End If
    End If
    If ((actual_blend_mode And render_set_states.rsADD) <> (blend_mode And render_set_states.rsADD)) Then
        If (blend_mode And render_set_states.rsADD) Then
            Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_ADD)
        Else
            Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
        End If
    End If
    actual_blend_mode = blend_mode
End Sub

Public Function Get_Blend_Mode() As Long
    Get_Blend_Mode = actual_blend_mode
End Function



Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub Text_Render(ByVal Font As Integer, ByRef text As String, ByVal top!, ByVal left!, _
                                ByVal width As Long, ByVal Height As Long, ByVal Color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If Not new_text Then
        Dim TextRect As RECT
        Dim ShadowRect As RECT
        
    
        
        
        TextRect.top = top
        TextRect.left = left
        TextRect.bottom = top + Height
        TextRect.right = left + width
    '    If TextRect.left < 0 Then
    '        TextRect.left = 0
    '        TextRect.Right = Width
    '        format = DT_LEFT
    '    ElseIf TextRect.left > 544 Then
    '        TextRect.left = 544 - Width
    '        TextRect.Right = 544
    '        format = DT_RIGHT
    '    End If
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        If shadow Then
            ShadowRect.top = top + 1
            ShadowRect.left = left + 1
            ShadowRect.bottom = top + Height + 1
            ShadowRect.right = left + width + 1
            D3DX.DrawText font_list(Font), &HFF000000, text, ShadowRect, format
        End If
        
        D3DX.DrawText font_list(Font), Color, text, TextRect, format
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    Else
        If format And DT_CENTER Then
            left = left - Engine.Engine_GetTextWidth(text) / 2
        End If
        text_render_graphic text, left, top, Color
    End If
End Sub

Public Sub Text_Render_alpha(ByRef text As String, ByVal top!, ByVal left!, ByVal Color As Long, ByVal format As Long, Optional ByVal alpha As Byte = 128)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
Dim color_s&

color_s = (Color And &HFFFFFF) Or Alphas(alpha) 'color - (&HFF000000 - color_s)

If format Then
    left = left - Engine.Engine_GetTextWidth(text) / 2
End If

text_render_graphic text, left, top, color_s, 1

End Sub

Public Sub Text_Render_ext(ByRef text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal width As Long, ByVal Height As Long, ByVal Color As Long, Optional ByVal alpha As Boolean, Optional ByVal center As Boolean)
Dim Alphas As Byte
Alphas = 255
If alpha = True Then _
    Alphas = 128
    
    If center = True Then
        Call Text_Render_alpha(text, top, left, Color, 1, Alphas)
    Else
        Call Text_Render_alpha(text, top, left, Color, 0, Alphas)
    End If
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.Name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, font_desc.hFont)
End Sub


Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean

'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
  MakeVector.X = X
  MakeVector.Y = Y
  MakeVector.z = z
End Function

Public Sub Device_Reset()
On Error GoTo errHandler
    D3DDevice.reset D3DWindow
    With D3DDevice
        Call Engine.SetVertexShader(FVF)
        Call .SetRenderState(D3DRS_LIGHTING, 0)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        Call .SetRenderState(D3DRS_POINTSIZE, Engine_FToDW(32))
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        Call .SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
        Call .SetRenderState(D3DRS_POINTSPRITE_ENABLE, 1)
        Call .SetRenderState(D3DRS_POINTSCALE_ENABLE, 0)
    End With
Exit Sub
errHandler:
End Sub

Public Sub Long_Color_Set_Alpha(ByRef Color As Long, ByVal alpha As Byte)
    'Dim barr(3) As Byte
    'DXCopyMemory barr(0), Color, 4
    'barr(0) = alpha
    Color = Color And &HFFFFFF Or Alphas(alpha) 'DXCopyMemory Color, barr(0), 4
End Sub


Public Sub Engine_set_max_fps(ByVal Limit As Boolean, Optional ByVal max_fps As Integer = 100)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
On Error GoTo errh
    limit_fps = True 'Limit
    'min_ms_between_render = CByte(1000 / max_fps)
Exit Sub
errh:
Debug.Print "No se limitaron las FPS, por error en ""Engine.engine_set_max_fps"" -> " & Err.Description
limit_fps = False
End Sub

Public Sub Engine_Toggle_fps_limit(Optional ByVal bool As Integer = -3)
If puedo_deslimitar Then
    If bool = -3 Then
        limit_fps = Not limit_fps
    Else
        limit_fps = CBool(bool)
    End If
Else
    limit_fps = True
End If
End Sub

Private Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS, adapter As CONST_D3DDEVTYPE, ByVal hWnd As Long) As Boolean
    Dim DispMode As D3DDISPLAYMODE
    On Error GoTo errOut

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    'If PDepth = 32 Then
    '    DispMode.format = D3DFMT_X8R8G8B8
    'ElseIf PDepth = 16 Then
    '    DispMode.format = D3DFMT_R5G6B5
    'End If
            
    With D3DWindow
        .Windowed = 2
        
        If UsarVSync Then
            .SwapEffect = CONST_D3DSWAPEFFECT.D3DSWAPEFFECT_COPY_VSYNC
        Else
            .SwapEffect = CONST_D3DSWAPEFFECT.D3DSWAPEFFECT_COPY
        End If
        .BackBufferFormat = DispMode.format
        .hDeviceWindow = hWnd
        .BackBufferWidth = Engine_Resolution.pixelesAncho
        .BackBufferHeight = Engine_Resolution.pixelesAlto
        '.EnableAutoDepthStencil = True
        '.AutoDepthStencilFormat = D3DFMT_D16
        
        .BackBufferCount = 1
    End With
    
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, adapter, hWnd, D3DCREATEFLAGS, D3DWindow)
    
    Call Device_Reset
    
    
    Engine_Init_D3DDevice = True

    LogDebug "   Device iniciado: " & D3DCREATEFLAGS

    Dim Jo As D3DCAPS8
    Call D3D.GetDeviceCaps(D3DADAPTER_DEFAULT, adapter, Jo)

    If Jo.MaxPointSize > 1 Then
        InitParticles Jo.MaxPointSize
    Else
        InitParticles 256
    End If

    
    If Jo.MaxPointSize = 0 Then LogError "El dispositivo no soporta PointSprites"
    
    DoEvents
    
Exit Function

errOut:
    Set D3DDevice = Nothing
    Engine_Init_D3DDevice = False
    ultimo_error = Err.Number
    ultimo_error_desc = D3DX.GetErrorString(Err.Number)
    
    
End Function

' hoja: identificador del objeto en donde se va a dibujar
' tilesAncho, tilesAlto: Cantidad de tiles que debe dibujar el motor grafico
Public Sub Engine_Init(hoja As Long, tilesAncho As Long, tilesAlto As Long, Optional ByVal max_fps As Integer = 100)

    Dim flags As Long
    
    Dim tamaño_textures As Long
    
    Dim ValidFormat As Boolean
    
    Dim ColorKeyVal As Long
    
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    Dim Caps8 As D3DCAPS8
    Dim DevType As CONST_D3DDEVTYPE
    Dim resultado As VbMsgBoxResult
    Dim hWnd As Long
    
    Call checkTimers
    
    LogDebug "  Iniciando dispositivo gráfico {"
    
    DevType = D3DDEVTYPE_HAL
    
    ' Este ON Error esta correcto
    On Local Error Resume Next

    Call D3D.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, Caps8)

    If Err.Number Then
        LogError "Error A01.1: " & vbNewLine & " El juego no ha encontrado un dispositivo acelerador compatible con Direct3D 8.1 - se utilizará un dispositivo de referencia. Probablemente habrá problemas"
        
        resultado = MsgBox("El juego no ha encontrado un dispositivo acelerador compatible con Direct3D 8.1. Esto se puede deber a que no tenes actualizado los Drivers de tu computadora. Hace clic en 'Sí' para obtener más información de como actualizarlos. ¿Desea obtener más información?", vbYesNo + vbExclamation, "Tierras del Sur")
        
        If resultado = vbYes Then
            openUrl ("https://tierrasdelsur.cc/soportes.html?entrada=176.problemas-tcnicos-bajos-fps")
            GoTo errHandler
        End If
                
        DevType = D3DDEVTYPE_REF
        Call D3D.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, Caps8)
        Err.Clear
    End If

    Err.Clear

    InitParticles 1

    If FileExist(app.Path & "\MZEngine3.dll") Then
        LogDebug "  Engine DLL encontrada."
    Else
        LogDebug "  Engine DLL NO encontrada."
        MsgBox "No se encontro el MZEngine3.dll, por favor reinstale el juego."
        If Not IsIDE Then End
    End If

    ' Donde voy a dibujar
    hWnd = hoja
    
    ' El tamaño de cada tile
    TilePixelWidth = 32
    TilePixelHeight = 32
    
    ' Cantidad de tiles que se ven en pantalla
    WindowTileHeight = tilesAlto
    WindowTileWidth = tilesAncho
    
    HalfWindowTileHeight = tilesAlto \ 2
    HalfWindowTileWidth = tilesAncho \ 2
    
    ' El tamaño de la pantalla para mostrar todos estos tiles
    MainViewHeight = tilesAlto * TilePixelHeight
    MainViewWidth = tilesAncho * TilePixelWidth

    'lColorMod = D3DTOP_MODULATE Or D3DTOP_MODULATE2X
    If Not Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE, DevType, hWnd) Then
        'lColorMod = D3DTOP_MODULATE Or D3DTOP_MODULATE2X
        If Not Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING, DevType, hWnd) Then
            'lColorMod = D3DTOP_MODULATE Or D3DTOP_MODULATE2X
            If Not Engine_Init_D3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING, DevType, hWnd) Then
                'lColorMod = D3DTOP_MODULATE
                If Not Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING, DevType, hWnd) Then

                    MsgBox "Error ""A01.2"" No se puede iniciar el dispositivo gráfico - (" & ultimo_error & ") " & D3DX.GetErrorString(ultimo_error) & " | " & ultimo_error_desc
                    LogError "Error ""A01.2"" No se puede iniciar el dispositivo gráfico - (" & ultimo_error & ") " & D3DX.GetErrorString(ultimo_error) & " | " & ultimo_error_desc
                    End
                End If
            End If
        End If
    End If

    LogDebug "      Device iniciado."
    Instanciar_Engine
    LogDebug "      Dll iniciado."
    
    Get_Capabilities
    Init_Math_Const
    Init_Lights

    ' Posición donde se encuentra el personaje principal.
    user_screen_pos.Y = (tilesAlto \ 2) * TilePixelHeight
    user_screen_pos.X = (tilesAncho \ 2) * TilePixelWidth
    
    Engine_Set_TileBuffer_Size 5, 5
    
    Const MB_Reservados_Graficos As Long = 128

    tamaño_textures = MB_Reservados_Graficos * 1024& * 1024&
    
    Debug.Print "Memoria de video libre para texturas:"; tamaño_textures
    
    Call Init_TextureDB(tamaño_textures, 500, RecursosPath & "Graficos.TDS")
    Call Sonido_Init(32& * 1024& * 1024&, 50, RecursosPath & "Sonidos.TDS")
    
    Call Entidades_Iniciar(100)
    
    LogDebug "      Biblioteca de texturas iniciada con " & MB_Reservados_Graficos & "MB reservados."

    engineBaseSpeed = 0.018

    ReDim mapdata(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As MapBlock

    'Set FPS value to 60 for startup
    FPS = 65
    FramesPerSecCounter = 65

    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8

    'Esto es para saber las posiciones caminables sin que se vean los bordes negros
    MinXBorder = X_MINIMO_VISIBLE + HalfWindowTileWidth
    MaxXBorder = X_MAXIMO_VISIBLE - HalfWindowTileWidth
    MinYBorder = Y_MINIMO_VISIBLE + HalfWindowTileHeight
    MaxYBorder = Y_MAXIMO_VISIBLE - HalfWindowTileHeight

    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder

    tVerts(0).v.z = 0!
    tVerts(0).rhw = 1!
    tVerts(1).v.z = 0!
    tVerts(1).rhw = 1!
    tVerts(2).v.z = 0!
    tVerts(2).rhw = 1!
    tVerts(3).v.z = 0!
    tVerts(3).rhw = 1!
    
    Init_Tilesets

    With tBox
        .Z0 = 0!
        .Z1 = 0!
        .z2 = 0!
        .Z3 = 0!
        .rhw0 = 1!
        .rhw1 = 1!
        .rhw2 = 1!
        .rhw3 = 1!
    End With

    CargarParticle_Streams
    LogDebug "      Partículas cargadas."
    
    
    InitLostDevice
    
    Engine_Init_FontSettings
    
    Set Fps_Label = New clsGUIText

    new_text = True
    Engine_set_max_fps False, max_fps

    Font_Create "Tahoma", 8, True, 0
    LogDebug "      Fuentes cargadas."

    color_mod_day.r = 1
    color_mod_day.g = 1
    color_mod_day.b = 1

setup_ambient
timeBeginPeriod 1

    LogDebug "      Dibujando primera escena de prueba."
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    LogDebug "      Escena OK."

    Epsilon = 5.96046447753906E-08
    calculate_epsilon

    LogDebug "  } Dispositivo gráfico Iniciado correctamente."
    
    Set FPS_LIMITER = New clsPerformanceTimer
    Set FRAME_TIMER = New clsPerformanceTimer
    FPS_LIMITER.Time
    FRAME_TIMER.Time
    LogDebug "  Timers iniciados."

    Set_Blend_Mode 0
    
    
    
    LogDebug "  Iniciando shaders..."

    PixelShaderBump = CreateShaderFromCode(ShaderSombra)
    
    Engine_PixelShaders.Engine_PixelShaders_Iniciar
    
    Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.Ninguno, vbNullString, FVF
    
    Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.estandar, ShaderSombra, (FVF Or D3DFVF_TEX2 Or D3DFVF_TEX3)
    
    If cfgSoportaPointSprites Then
        Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.Particulas, vbNullString, particleFVF
    Else
        Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.Particulas, vbNullString, FVF
    End If
    
    Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.Normales, _
            "ps.1.0;" & vbNewLine & _
            "tex t1;" & vbNewLine & _
            "mov r0, t1;", _
        (FVF Or D3DFVF_TEX2 Or D3DFVF_TEX3)
        
    Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.ColoresAmbiente, _
            "ps.1.0;" & vbNewLine & _
            "tex t2;" & vbNewLine & _
            "mov r0, t2;", _
        (FVF Or D3DFVF_TEX2 Or D3DFVF_TEX3)
    
    Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.ColoresLuces, _
            "ps.1.0;" & vbNewLine & _
            "mov r0, v0;", _
        (FVF Or D3DFVF_TEX2 Or D3DFVF_TEX3)
        
    Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.Agua, _
            "ps.1.0;" & vbNewLine & _
"tex t0;" & vbNewLine & _
"tex t2;" & vbNewLine & _
"mov r0, t2;" & vbNewLine & _
"mul r0, t0, r0;" & vbNewLine & _
"mul r0, r0, c1;", _
        (FVF Or D3DFVF_TEX2 Or D3DFVF_TEX3)
            
    Engine_PixelShaders.Engine_PixelShaders_Setear ePixelShaders.Pisos, _
            ShaderSombra, _
        (FVF Or D3DFVF_TEX2 Or D3DFVF_TEX3), _
            "vs.1.0;" & vbNewLine & _
"mov r0, v0;" & vbNewLine & _
"mov oD0, v5;" & vbNewLine & _
"mov oT0, v7.xy;" & vbNewLine & _
"mov oT1, v7.xy;" & vbNewLine & _
"m4x4 r1, r0, c0; // Posicion en R1 de las tiles" & vbNewLine & _
"add r0.y, r0.y, -r0.z; // Altura de las tiles" & vbNewLine & _
"m4x4 oPos, r0, c0; // Pos en oPos de las tiles con altura" & vbNewLine & _
"mul r1, r1, c7;" & vbNewLine & _
"add r1.x, r1.x, c7.y;" & vbNewLine & _
"add r1.y, r1.y, -c7.y;" & vbNewLine & _
"mov oT2, r1.xy;"
        
    LogDebug "  Shaders OK."
    
    
    LogDebug "      Dibujando primera escena de prueba de luces"
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
    
    LightBackBuffer_Init
    
    Engine_LightsTexture_Init
    
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    LogDebug "      Escena OK."
    
    InitCamera
    ColoresAguaInit
    MapBox_Init
    
    prgRun = True

Exit Sub
errHandler:
MsgBox "Error: A01.0 [ " & Err.Number & " ] - " & D3DX.GetErrorString(Err.Number)
LogError "Error: A01.0 [ " & Err.Number & " ] - " & D3DX.GetErrorString(Err.Number)

prgRun = False

End Sub

Private Sub calculate_epsilon()
' EPSILON PUEDE CAMBIAR DEPENDIENDO DE LA ARQUITECTURA DEL PROCESADOR 10/01/2011 menduz: AJJAJAJ QUE PETE(?)
        Dim machEps!
        machEps = 1
        Do
           machEps = machEps / 2
        Loop While ((1 + (machEps / 2)) <> 1)
        
        If machEps <> 0 Then
            Epsilon = machEps
        End If
End Sub

Public Sub Engine_Deinit()

    timeEndPeriod 1
    Sonido_DeInit
    Erase mapdata
    Erase CharList
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set DX = Nothing
    #If esCLIENTE = 1 Then
        'TODO esto hace que no compile
       ' If conectado Then desconectar
    #End If
End Sub




    
    
    
    
    
    
    
    



Public Sub cTLVertex(ByRef tl As TLVERTEX, ByRef X As Single, ByRef Y As Single, ByRef Color As Long, ByRef tu As Single, ByRef tv As Single)
    tl.v.X = X
    tl.v.Y = Y
    tl.v.z = 0!
    tl.rhw = 1!
    tl.Color = Color
    tl.tu = tu
    tl.tv = tv
End Sub

Public Sub Engine_ActFPS()
    If GetTimer - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTimer
    End If
End Sub

Public Sub Draw_Barba(ByVal textura As Integer, ByVal XX%, ByVal YY%, ByVal heading As E_Heading, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal alt As Byte = 0, Optional ByVal mirror As Byte = 0, Optional ByVal mirrorv As Byte = 0, Optional ByVal alpha As Byte = 0)

    Dim CurrentGrhIndex As Integer
    Dim resto As Integer
    
        
    If map_x = 0 Then map_x = 1
    If map_y = 0 Then map_y = 1
    If alt = 0 Then YY = YY - AlturaPie(map_x, map_y)
        
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(textura)
    
    Colorear_TBOX tBox, map_x, map_y
    
    Dim w As Integer
    Dim h As Integer
    Dim pixelHeight As Integer
    Dim pixelWidth As Integer
    Dim sy As Integer
    Dim sx As Integer
    
    If heading = E_Heading.SOUTH Then
        sx = 0
    ElseIf heading = E_Heading.EAST Then
        sx = 32
    ElseIf heading = E_Heading.WEST Then
        sx = 64
    ElseIf heading = E_Heading.NORTH Then
        sx = 96
    End If
    
    pixelHeight = 32
    pixelWidth = 32
    sy = 0
    w = 128
    h = 32
    
    Dim TGRH As GrhData
    
    With TGRH
        .tu(0) = sx / w
        .tv(0) = (sy + pixelHeight) / h
        .tu(1) = .tu(0)
        .tv(1) = sy / h
        .tu(2) = (sx + pixelWidth) / w
        .tv(2) = .tv(0)
        .tu(3) = .tu(2)
        .tv(3) = .tv(1)
    End With
    
    With tBox
        .x0 = XX
        .y0 = YY + pixelHeight
        .x1 = .x0
        .y1 = YY
        .x2 = XX + pixelWidth
        .y2 = .y0
        .x3 = .x2
        .y3 = .y1

        If alpha Then
            .color0 = (.color0 And &HFFFFFF) Or Alphas(alpha)
            .Color1 = (.Color1 And &HFFFFFF) Or Alphas(alpha)
            .Color2 = (.Color2 And &HFFFFFF) Or Alphas(alpha)
            .color3 = (.color3 And &HFFFFFF) Or Alphas(alpha)
        End If
        
        If mirror Then
            .tu0 = TGRH.tu(2)
            .tv0 = TGRH.tv(2)
            .tu1 = TGRH.tu(3)
            .tv1 = TGRH.tv(3)
            .tu2 = TGRH.tu(0)
            .tv2 = TGRH.tv(0)
            .tu3 = TGRH.tu(1)
            .tv3 = TGRH.tv(1)
        Else
            .tu0 = TGRH.tu(0)
            .tv0 = TGRH.tv(0)
            .tv1 = TGRH.tv(1)
            .tu2 = TGRH.tu(2)
            .tu1 = .tu0
            .tv2 = .tv0
            .tu3 = .tu2
            .tv3 = .tv1
        End If
        
        If mirrorv Then
            Dim bsf!
            bsf = .tv0
            .tv0 = .tv1
            .tv1 = bsf
            bsf = .tv2
            .tv2 = .tv3
            .tv3 = bsf
        End If

    End With
    

    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.estandar
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
End Sub

Public Sub Draw_Grh(ByRef Grh As Grh, ByVal XX%, ByVal YY%, ByVal Animate As Byte, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal alt As Byte = 0, Optional ByVal mirror As Byte = 0, Optional ByVal mirrorv As Byte = 0, Optional ByVal alpha As Byte = 0, Optional ByVal texturaOverwride As Integer = 0)
    Dim CurrentGrhIndex As Integer
    Dim resto As Integer
    
    If Grh.GrhIndex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            ' Frame counter tiene que ir entre (1 y La cantidad frames.9999)
            ' Para elegir el frame finalmente se tomará la parte entera de FrameCounter
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            ' Si supera frames.9999 o sea que es igual a frame+1. o superior.
            ' Se resetea, pero calculando el que le corresponde
            If Grh.FrameCounter >= (GrhData(Grh.GrhIndex).NumFrames + 1) Then
                ' FrameCounter siempre va a ser superior a NumFrames en al menos "1", sino nunca ingresaría aquí.
                ' No se puede usar MOD acá porque se perderia la interpolacion ya que trabaja con enteros
                resto = Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames
                ' Esto es porque los frames va a de 1 a N...
                If resto = 0 Then resto = GrhData(Grh.GrhIndex).NumFrames
                Grh.FrameCounter = resto + (Grh.FrameCounter - Fix(Grh.FrameCounter)) ' Le sumo los decimales
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
                

            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).frames(Fix(Grh.FrameCounter))

    If CurrentGrhIndex > 0 Then
        XX = XX + GrhData(CurrentGrhIndex).offsetX
        YY = YY + GrhData(CurrentGrhIndex).offsetY
        
        If map_x = 0 Then map_x = 1
        If map_y = 0 Then map_y = 1
        If alt = 0 Then YY = YY - AlturaPie(map_x, map_y)
        
        Grh_Render_new CurrentGrhIndex, XX, YY, map_x, map_y, mirror, mirrorv, alpha, texturaOverwride
    End If
End Sub

Public Sub Draw_Grh_Interpolador(ByRef Grh As Grh, ByVal XX%, ByVal YY%, ByVal Animate As Byte, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal alt As Byte = 0, Optional ByVal mirror As Byte = 0, Optional ByVal mirrorv As Byte = 0, Optional ByVal mu As Single = 0, Optional ByVal alpha As Byte = 0, Optional ByVal texturaOverride As Integer = 0)
    Dim CurrentGrhIndex As Integer
    If Grh.GrhIndex = 0 Then Exit Sub
    Dim j As Integer
    
    'If Grh.Started = 1 Then
    '    Grh.Loops = 0
    '    Draw_Grh Grh, XX, YY, center, Animate, map_x, map_y, alt, mirror, mirrorv
    '    Exit Sub
    'End If
    
    
    If Animate Then
        'Grh.FrameCounter = (Interp(GrhData(Grh.GrhIndex).NumFrames, 0, mins(mu, 1)) Mod GrhData(Grh.GrhIndex).NumFrames) + 1
        j = GrhData(Grh.GrhIndex).NumFrames
        If j > 1 Then
            Grh.FrameCounter = (j * (1 - mu)) + 1
            
            If Fix(Grh.FrameCounter) = j + 1 Then
                Grh.FrameCounter = 1
                If Grh.Started = 1 Then Grh.Started = 0
            End If
        Else
            Grh.FrameCounter = 1
        End If
    Else
        Grh.FrameCounter = 1
    End If
    
    Grh.FrameCounter = Abs(Grh.FrameCounter)
On Error Resume Next
    CurrentGrhIndex = GrhData(Grh.GrhIndex).frames(Fix(Grh.FrameCounter))
    
    If CurrentGrhIndex = 0 Then
        Exit Sub
    End If
    XX = XX + GrhData(CurrentGrhIndex).offsetX
    YY = YY + GrhData(CurrentGrhIndex).offsetY
    
    If map_x = 0 Then map_x = 1
    If map_y = 0 Then map_y = 1
    If alt = 0 Then YY = YY - AlturaPie(map_x, map_y)
    
    Grh_Render_new CurrentGrhIndex, XX, YY, map_x, map_y, mirror, mirrorv, alpha, texturaOverride
End Sub

Public Sub Draw_GrhE(ByRef Grh As Grh, ByVal XX%, ByVal YY%, ByVal Animate As Byte, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal alt As Byte = 0, Optional ByVal mirror As Byte = 0, Optional ByVal mirrorv As Byte = 0, Optional ByVal mu As Single = 0)
    Dim CurrentGrhIndex As Integer
    Dim j As Integer
    
    If Grh.GrhIndex = 0 Then Exit Sub

    If Grh.GrhIndex = 0 Then Exit Sub
    
    If Animate Then
        'Grh.FrameCounter = (Interp(GrhData(Grh.GrhIndex).NumFrames, 0, mins(mu, 1)) Mod GrhData(Grh.GrhIndex).NumFrames) + 1
        j = GrhData(Grh.GrhIndex).NumFrames
        If j > 1 Then
            Grh.FrameCounter = (j * (1 - mu)) + 1
            
            If Fix(Grh.FrameCounter) = j + 1 Then
                Grh.FrameCounter = 1
                If Grh.Started = 1 Then Grh.Started = 0
            ElseIf Grh.FrameCounter < 1 Then
                Grh.FrameCounter = 1
            End If
        Else
            Grh.FrameCounter = 1
        End If
    Else
        Grh.FrameCounter = 1
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).frames(Fix(Grh.FrameCounter))

    XX = XX + GrhData(CurrentGrhIndex).offsetX
    YY = YY + GrhData(CurrentGrhIndex).offsetY
    
    If map_x = 0 Then map_x = 1
    If map_y = 0 Then map_y = 1
    If alt = 0 Then YY = YY - AlturaPie(map_x, map_y)
    
    Grh_Render_reflejo CurrentGrhIndex, XX, YY, map_x, map_y, mirror, mirrorv, 80
End Sub

Public Sub Draw_Grh_Alpha(ByRef Grh As Grh, ByVal X!, ByVal Y!, ByVal Animate As Byte, Optional ByVal alpha As Byte, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal alt As Byte = 0)
    Dim CurrentGrhIndex As Integer
    Dim resto As Integer
    
    If Grh.GrhIndex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            ' Frame counter tiene que ir entre (1 y La cantidad frames.9999)
            ' Para elegir el frame finalmente se tomará la parte entera de FrameCounter
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            ' Si supera frames.9999 o sea que es igual a frame+1. o superior.
            ' Se resetea, pero calculando el que le corresponde
            If Grh.FrameCounter >= (GrhData(Grh.GrhIndex).NumFrames + 1) Then
                ' FrameCounter siempre va a ser superior a NumFrames en al menos "1", sino nunca ingresaría aquí.
                ' No se puede usar MOD acá porque se perderia la interpolacion ya que trabaja con enteros
                resto = Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames
                ' Esto es porque los frames va a de 1 a N...
                If resto = 0 Then resto = GrhData(Grh.GrhIndex).NumFrames
                Grh.FrameCounter = resto + (Grh.FrameCounter - Fix(Grh.FrameCounter)) ' Le sumo los decimales
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).frames(Fix(Grh.FrameCounter))

    X = X + GrhData(CurrentGrhIndex).offsetX
    Y = Y + GrhData(CurrentGrhIndex).offsetY

    If map_x = 0 Then map_x = 1
    If map_y = 0 Then map_y = 1
    If alt = 0 Then Y = Y - AlturaPie(map_x, map_y)

    Grh_Render_new CurrentGrhIndex, X, Y, map_x, map_y, , , alpha

End Sub

Public Sub Draw_Grh_Techo(ByRef Grh As Grh, ByVal X!, ByVal Y!, ByVal Animate As Byte, Optional ByVal alpha As Byte = 255)
    Dim dest_x2%, dest_y2%
    Dim CurrentGrhIndex As Integer
    Dim Color As Long
    Dim resto As Integer
'calculo
    If Grh.GrhIndex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            ' Frame counter tiene que ir entre (1 y La cantidad frames.9999)
            ' Para elegir el frame finalmente se tomará la parte entera de FrameCounter
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            ' Si supera frames.9999 o sea que es igual a frame+1. o superior.
            ' Se resetea, pero calculando el que le corresponde
            If Grh.FrameCounter >= (GrhData(Grh.GrhIndex).NumFrames + 1) Then
                ' FrameCounter siempre va a ser superior a NumFrames en al menos "1", sino nunca ingresaría aquí.
                ' No se puede usar MOD acá porque se perderia la interpolacion ya que trabaja con enteros
                resto = Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames
                ' Esto es porque los frames va a de 1 a N...
                If resto = 0 Then resto = GrhData(Grh.GrhIndex).NumFrames
                Grh.FrameCounter = resto + (Grh.FrameCounter - Fix(Grh.FrameCounter)) ' Le sumo los decimales
                                
              '  Grh.FrameCounter = 1
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).frames(Fix(Grh.FrameCounter))
    
    If CurrentGrhIndex = 0 Then Exit Sub

    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(CurrentGrhIndex).filenum)
    If GrhData(CurrentGrhIndex).hardcor = 0 Then Init_grh_tutv CurrentGrhIndex
    
    X = X + GrhData(CurrentGrhIndex).offsetX
    Y = Y + GrhData(CurrentGrhIndex).offsetY
    
'render
    Color = base_light_techo Or Alphas(alpha)
    
    dest_y2 = Y + GrhData(CurrentGrhIndex).pixelHeight
    dest_x2 = X + GrhData(CurrentGrhIndex).pixelWidth
    
'Debug.Print Hex(Color)
    With tBox
        .x0 = X
        .y0 = dest_y2
        .x1 = X
        .y1 = Y
        .x2 = dest_x2
        .y2 = dest_y2
        .x3 = dest_x2
        .y3 = Y
        .color0 = Color
        .Color1 = Color
        .Color2 = Color
        .color3 = Color
        .tu0 = GrhData(CurrentGrhIndex).tu(0)
        .tv0 = GrhData(CurrentGrhIndex).tv(0)
        .tv1 = GrhData(CurrentGrhIndex).tv(1)
        .tu2 = GrhData(CurrentGrhIndex).tu(2)
        .tu1 = .tu0
        .tv2 = .tv0
        .tu3 = .tu2
        .tv3 = .tv1
    End With
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Normal Nothing
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
    Grh_Render_Complementario GrhData(CurrentGrhIndex).filenum
End Sub

Public Sub init_gui_tl(ByRef vert() As TLVERTEX, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal Height As Integer, Optional ByVal Color As Long = &HFFFFFFFF)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    vert(1) = Geometry_Create_TLVertex(left, top, Color, 0, 0)
    vert(3) = Geometry_Create_TLVertex(left + width, left, Color, width / 256, 0)
    vert(0) = Geometry_Create_TLVertex(left, top + Height, Color, 0, Height / 256)
    vert(2) = Geometry_Create_TLVertex(left + width, top + Height, Color, width / 256, Height / 256)
End Sub

Private Sub init_gui_tl_indexed(ByRef vert() As TLVERTEX, ByVal top As Single, ByVal left As Single, ByVal width As Single, ByVal Height As Single, ByVal sx As Single, ByVal sy As Single, ByVal tex_dimension As Integer, Optional ByVal Color As Long = &HFFFFFFFF)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    vert(0) = Geometry_Create_TLVertex(top, left, Color, sx / tex_dimension, sy / tex_dimension)
    vert(1) = Geometry_Create_TLVertex(top + width, left, Color, (width + 1 + sx) / tex_dimension, sy / tex_dimension)
    vert(2) = Geometry_Create_TLVertex(top, left + Height, Color, sx / tex_dimension, (Height + sy + 1) / tex_dimension)
    vert(3) = Geometry_Create_TLVertex(top + width, left + Height, Color, (width + sx + 1) / tex_dimension, (Height + 1 + sy) / tex_dimension)
End Sub

Public Function Device_Test_Cooperative_Level() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'Handle Alt-Tab and Ctrl-Alt-Del
'**************************************************************
    'Call TestCooperativeLevel to see what state the device is in.
    Dim hr As Long
    hr = D3DDevice.TestCooperativeLevel
    If hr = D3DERR_DEVICELOST Then
        Exit Function
    ElseIf hr = D3DERR_DEVICENOTRESET Then
        Device_Reset
        Exit Function
    End If

    Device_Test_Cooperative_Level = True
End Function

Public Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Color As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
    Geometry_Create_TLVertex.v.X = X
    Geometry_Create_TLVertex.v.Y = Y
    Geometry_Create_TLVertex.v.z = 0
    Geometry_Create_TLVertex.rhw = 1
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function



Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef textures_size As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim Radius As Single
    Dim x_cor As Single
    Dim y_cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.right - dest.left) / 2
        y_center = dest.top + (dest.bottom - dest.top) / 2
        
        'Calculate radius
        Radius = Sqr((dest.right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.right - x_center) / Radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = pi - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_cor = dest.left
        y_cor = dest.bottom
    Else
        x_cor = x_center + Cos(-left_point - angle) * Radius
        y_cor = y_center - Sin(-left_point - angle) * Radius
    End If
    
    
    '0 - Bottom left vertex
    If textures_size Then
        verts(0) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(0), src.left / textures_size, (src.bottom + 1) / textures_size)
    Else
        verts(0) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(0), 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_cor = dest.left
        y_cor = dest.top
    Else
        x_cor = x_center + Cos(left_point - angle) * Radius
        y_cor = y_center - Sin(left_point - angle) * Radius
    End If
    
    
    '1 - Top left vertex
    If textures_size Then
        verts(1) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(1), src.left / textures_size, src.top / textures_size)
    Else
        verts(1) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(1), 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_cor = dest.right
        y_cor = dest.bottom
    Else
        x_cor = x_center + Cos(-right_point - angle) * Radius
        y_cor = y_center - Sin(-right_point - angle) * Radius
    End If
    
    
    '2 - Bottom right vertex
    If textures_size Then
        verts(2) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(2), (src.right + 1) / textures_size, (src.bottom + 1) / textures_size)
    Else
        verts(2) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(2), 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_cor = dest.right
        y_cor = dest.top
    Else
        x_cor = x_center + Cos(right_point - angle) * Radius
        y_cor = y_center - Sin(right_point - angle) * Radius
    End If
    
    
    '3 - Top right vertex
    If textures_size Then
        verts(3) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(3), (src.right + 1) / textures_size, src.top / textures_size)
    Else
        verts(3) = Geometry_Create_TLVertex(x_cor, y_cor, rgb_list(3), 0, 0)
    End If
End Sub


Public Sub PS_SetearColoresAmbiente()
    PS_Constants(0, 1) = color_mod_day.r 'mins(C0, minl(C1, minl(C2, C3))) / 255
    PS_Constants(1, 1) = color_mod_day.g 'PS_Constants(0, 1)
    PS_Constants(2, 1) = color_mod_day.b 'PS_Constants(0, 1)
    PS_Constants(3, 1) = 127
    
    D3DDevice.SetPixelShaderConstant 0, PS_Constants(0, 0), 3
End Sub




Public Sub CalcularNormalColorTBOX_Textura(ByRef tBox As Box_Vertex)
    With tBox
        .tu02 = .x0 / D3DWindow.BackBufferWidth
        .tu12 = .x1 / D3DWindow.BackBufferWidth
        .tu22 = .x2 / D3DWindow.BackBufferWidth
        .tu32 = .x3 / D3DWindow.BackBufferWidth
        
        .tv02 = .y0 / D3DWindow.BackBufferHeight
        .tv12 = .y1 / D3DWindow.BackBufferHeight
        .tv22 = .y2 / D3DWindow.BackBufferHeight
        .tv32 = .y3 / D3DWindow.BackBufferHeight
    End With
End Sub

Public Sub CalcularNormalColorTBOX_TexturaVertical(ByRef tBox As Box_Vertex)
    With tBox
        .tu02 = .x0 / D3DWindow.BackBufferWidth
        .tu12 = .x1 / D3DWindow.BackBufferWidth
        .tu22 = .x2 / D3DWindow.BackBufferWidth
        .tu32 = .x3 / D3DWindow.BackBufferWidth
        
        .tv02 = .y0 / D3DWindow.BackBufferHeight
        .tv12 = (.y0 - 32) / D3DWindow.BackBufferHeight
        .tv22 = .y2 / D3DWindow.BackBufferHeight
        .tv32 = (.y2 - 32) / D3DWindow.BackBufferHeight
    End With
End Sub








'MARCE EH?
Public Sub Colorear_Cuadrado(v() As TLVERTEX, ByVal map_x As Byte, ByVal map_y As Byte)
    Dim tmpc As Long
    If map_x < X_MAXIMO_VISIBLE And map_y > Y_MINIMO_VISIBLE Then
        v(0).Color = ResultColorArray(map_x, map_y)
        v(1).Color = ResultColorArray(map_x, map_y - 1)
        v(2).Color = ResultColorArray(map_x + 1, map_y)
        v(3).Color = ResultColorArray(map_x + 1, map_y - 1)
    End If
End Sub


Public Sub Init_grh_tutv(ByVal GrhIndex As Integer)
    Dim h!, w!
    Call GetTextureDimension(GrhData(GrhIndex).filenum, h, w)
    If h = 0 Then Exit Sub
    With GrhData(GrhIndex)
        .tu(0) = .sx / w
        .tv(0) = (.sy + .pixelHeight) / h
        .tu(1) = .tu(0)
        .tv(1) = .sy / h
        .tu(2) = (.sx + .pixelWidth) / w
        .tv(2) = .tv(0)
        .tu(3) = .tu(2)
        .tv(3) = .tv(1)
        .hardcor = 1
    End With
'            .tu(0) = .SX / W
'            .tv(0) = (.SY + .pixelHeight + 1) / H
'            .tu(1) = .SX / W
'            .tv(1) = .SY / H
'            .tu(2) = (.SX + .pixelWidth + 1) / W
'            .tv(2) = (.SY + .pixelHeight + 1) / H
'            .tu(3) = (.SX + .pixelWidth + 1) / W
'            .tv(3) = .SY / H
'            .hardcor = 1


End Sub

Public Function d3d_format_id(ByVal fmt As Long) As Long
Select Case fmt
        Case D3DFMT_R5G6B5:     d3d_format_id = 1
        Case D3DFMT_X1R5G5B5:   d3d_format_id = 2
        Case D3DFMT_A1R5G5B5:   d3d_format_id = 3
        Case D3DFMT_X8R8G8B8:   d3d_format_id = 4
        Case D3DFMT_A8R8G8B8:   d3d_format_id = 5
        Case Else:              d3d_format_id = 0
End Select
End Function



Public Function Engine_Gfx_BeginScene() As Boolean
    Dim Mode As D3DDISPLAYMODE
    Dim hr As Long

    hr = D3DDevice.TestCooperativeLevel
    If (hr = D3DERR_DEVICELOST) Then
        Exit Function
    ElseIf (hr = D3DERR_DEVICENOTRESET) Then
        On Local Error GoTo Errh1
        Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, Mode)
        If Mode.format = D3DFMT_UNKNOWN Then
Errh1:
            MsgBox "No se pudo obtener la información del dispositivo."
        End If
        'Marce 'Marce On local error goto 0

        D3DWindow.BackBufferFormat = Mode.format

        If (d3d_format_id(Mode.format) < 4) Then
            nScreenBPP = 16
        Else
            nScreenBPP = 32
        End If

        If (Not Engine_Gfx_Restore()) Then Exit Function
    End If

    If Engine_Escene_Abierta Then
        MsgBox "Ya hay una escena abierta."
        Exit Function
    End If

    D3DDevice.BeginScene

    SceneBegin = GetTimer


    IndiceSeteado = False
    
    Engine_Escene_Abierta = True
    Engine_Gfx_BeginScene = True
End Function

Public Sub Engine_Gfx_Clear(Optional ByVal Color As Long = &HFF000000, Optional ByVal Villero As Byte = 0)
'    If Villero = 0 Then
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, Color, 0, 0
'    Else
'        Engine.Draw_FilledBox 0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, Color, 0, 0
'    End If
End Sub

Public Sub Engine_Gfx_EndScene()
    Dim mainrect As RECT
    
    Engine_Escene_Abierta = False
    
    mainrect.top = 0
    mainrect.left = 0
    mainrect.right = MainViewWidth
    mainrect.bottom = MainViewHeight

    D3DDevice.EndScene
    D3DDevice.Present mainrect, ByVal 0, frmMain.Renderer.hWnd, ByVal 0
    
    Exit Sub
errh3:
    MsgBox "No se pudo reiniciar el dispositivo gráfico."
    LogDebug "Engine_EndScene! ERROOOR!"
    'Engine_Deinit
    'End
End Sub

Private Function Engine_Gfx_Restore() As Boolean
On Error GoTo errh
    LogDebug "REINICIANDO EL DEVICE!"
    
    D3DDevice.SetIndices Nothing, 0
    D3DDevice.SetStreamSource 0, Nothing, TL_size
    D3DDevice.reset D3DWindow
    
    InitLostDevice
    Engine_Gfx_Restore = True
Exit Function
errh:
    MsgBox "No se pudo reiniciar el dispositivo gráfico."
    LogDebug "REINICIANDO EL DEVICE! ERROOOR!"
    Engine_Deinit
    End
End Function


Private Sub InitLostDevice()
    Engine.SetVertexShader FVF

    D3DDevice.SetRenderState D3DRS_LIGHTING, False

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(32)

    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 1
    
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0
    
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE 'ACAKB
    D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    
    IniciarIndexBuffers
End Sub

#If esCLIENTE = 1 Then

Public Sub Engine_Start()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
DoEvents

Dim cut_fps_ud As Long
Dim conteo As Double

LogDebug "Iniciando Bucle del programa."

Do While prgRun
    
    Rem Limitar FPS


    If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
        CheckKeys
        RENDERCRC = (2147483647 * Rnd) Xor GetTickCount
        
        Render
        
        cut_fps_ud = GetTimer
        FramesPerSecCounter = FramesPerSecCounter + 1
        If (cut_fps_ud - lFrameTimer) >= 1000 Then
            FPS = FramesPerSecCounter
            frmMain.FPS.Caption = FPS
            FramesPerSecCounter = 0
            lFrameTimer = cut_fps_ud
            'FPS_LIMITER.Time
            'Call ResetFPS
        End If
        
        Call rm2a
    Else
        RenderInterface
        Sleep 16&
    End If


    Do
        ' En esta parte del bucle se tienen que hacer las cosas que no afectan al
        ' tiempo de render, como por ejemplo leer el socket o reproducir sonidos.
        
        ' Como VB es un solo thread con eventos, corremos todos los eventos encolados ahora.
        DoEvents

        'Audio.Music_GetLoop
     Loop While (Not UsarVSync And BSleepForFrameRateLimit(150, SceneBegin))
Loop

DoEvents

LogDebug "Fin del bucle del programa."

End Sub
#End If

Public Function BSleepForFrameRateLimit(ByVal ulMaxFrameRate As Double, ByVal m_ulGameTickCount As Long) As Boolean

  ' Frame rate limiting
  Dim flDesiredFrameMilliseconds  As Single
  flDesiredFrameMilliseconds = 1000 / ulMaxFrameRate

  Dim flMillisecondsElapsed As Long
  flMillisecondsElapsed = GetTimer(False) - m_ulGameTickCount
  
  If flMillisecondsElapsed < flDesiredFrameMilliseconds Then
    ' If enough time is left sleep, otherwise just keep spinning so we don't go over the limit...
    If (flDesiredFrameMilliseconds - flMillisecondsElapsed) > 3 Then
        Sleep 5
    End If

    BSleepForFrameRateLimit = True
  Else
    BSleepForFrameRateLimit = False
  End If

End Function

Public Function CreateColorVal(a As Single, r As Single, g As Single, b As Single) As D3DCOLORVALUE
    CreateColorVal.a = a
    CreateColorVal.r = r
    CreateColorVal.g = g
    CreateColorVal.b = b
End Function

Public Function Engine_FToDW(f As Single) As Long
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Call DXCopyMemory(Engine_FToDW, f, 4)
End Function

Private Function VectorToRGBA(vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, g As Integer, b As Integer, a As Integer
    r = 127 * vec.X + 128
    g = 127 * vec.Y + 128
    b = 127 * vec.z + 128
    a = 255 * fHeight
    VectorToRGBA = D3DColorARGB(a, r, g, b)
End Function


Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal Height As Integer, Color As Long, outlinecolor As Long, Optional ByVal lh As Integer = 1)
    Static box_rect As RECT
    Static OutLine As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    rgb_list(0) = Color
    rgb_list(1) = Color
    rgb_list(2) = Color
    rgb_list(3) = Color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .bottom = Y + Height - lh
        .left = X + lh
        .right = X + width - lh
        .top = Y + lh
    End With
    
    With OutLine
        .bottom = Y + Height
        .left = X
        .right = X + width
        .top = Y
    End With
    
    Geometry_Create_Box Vertex2(), OutLine, OutLine, rgb_list2(), 0, 0
    Geometry_Create_Box vertex(), box_rect, box_rect, rgb_list(), 0, 0
    last_texture = 0
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
            
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), TL_size
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertex(0), TL_size
End Sub

Public Function FPSLimiter(bool As Boolean) As Single
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    Call QueryPerformanceCounter(start_time)
    FPSLimiter = (start_time - end_time) / timer_freq * 1000
    If bool = True Then _
        Call QueryPerformanceCounter(end_time)
End Function

Public Function FPSTIMER(Optional ByVal reset As Byte) As Double
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim start_time As Currency
    Dim end_time As Currency
    Static timer_freq As Double
    Static stt As Double
    Dim tt As Double
    
    If timer_freq = 0 Then
        QueryPerformanceFrequency end_time
        timer_freq = CDbl(end_time)
        Call QueryPerformanceCounter(start_time)
        stt = CDbl(start_time)
        
    End If
    
    Call QueryPerformanceCounter(end_time)
    
    tt = (CDbl(end_time) - stt) / timer_freq
    
    If tt > 1 Then
        Call QueryPerformanceCounter(start_time)
        stt = CDbl(start_time)
        zNFrames = 1
    End If
    
    If tt = 0 Then tt = 1
    
    FPSTIMER = Round(((zNFrames And &H7FFFFFFF) / tt), 1)
    zNFrames = zNFrames + 1
    
    
End Function



Public Sub text_render_graphic(T$, X!, Y!, Optional ByVal Color As Long = &HFFFFFFFF, Optional ByVal scalea As Single = 1)


    
    Dim i As Integer, j As Integer
    
    Dim lenght&
    lenght = Len(T)
    If lenght = 0 Then Exit Sub
    
    
    Dim ind&, TempStr$()
    Dim TLV() As Box_Vertex
    
    ReDim TLV((Len(T) * 4) - 1)

    'Call GetTexture(TexturaTexto)

    X = Round(X)
    Y = Round(Y)

    
    Dim Ascii() As Byte
    
    
    
    
    
    Dim KeyPhrase As Byte
    Dim TempColor As Long
    Dim ResetColor As Byte
    
    
    
    
    Dim lena&
    
    Dim TmpX!, TmpY!
    
    TempColor = Color
        TempStr = Split(T, vbCrLf)
        For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            TmpY = i * Font_Default.CharHeight * scalea + Y
            TmpX = X
        
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            lena = Len(TempStr(i)) - 1
            'Loop through the characters
            For j = 0 To lena
                If Ascii(j) = 255 Then 'If Ascii = "|"124
                    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
                    If KeyPhrase Then TempColor = &HFFAACCAA Else ResetColor = 1
                Else
                    'Copy from the cached vertex array to the temp vertex array
                    CopyMemory TLV(ind), Font_Default.HeaderInfo.CharVA(Ascii(j)).vertex, BV_size

                    TLV(ind).x0 = TmpX
                    TLV(ind).y0 = TmpY
                    
                    TLV(ind).x1 = TLV(ind).x1 + TmpX '* scalea
                    TLV(ind).y1 = TLV(ind).y0
    
                    TLV(ind).x2 = TmpX
                    TLV(ind).y2 = TLV(ind).y2 + TmpY '* scalea
    
                    TLV(ind).x3 = TLV(ind).x1
                    TLV(ind).y3 = TLV(ind).y2
                    
                    TLV(ind).color0 = TempColor
                    TLV(ind).Color1 = TempColor
                    TLV(ind).Color2 = TempColor
                    TLV(ind).color3 = TempColor
                        
                    ind = ind + 1
                    TmpX = TmpX + Font_Default.HeaderInfo.CharWidth(Ascii(j)) '* scalea
                
                End If
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
                
            Next j
            
        End If
    Next i
    
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(TexturaTexto2)
    
    If IndexBufferEnabled Then
        'd3ddevice.SetStreamSource
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLELIST, 0, ind * 4, ind * 2, StaticIndexBuffer(0), D3DFMT_INDEX16, TLV(0), TL_size
    Else
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, ind * 4 - 1, TLV(0), TL_size
    End If

End Sub

Public Function Engine_GetTextWidth(ByVal text As String) As Integer
    Dim i As Integer
    If LenB(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        Engine_GetTextWidth = Engine_GetTextWidth + Font_Default.HeaderInfo.CharWidth(Asc(mid$(text, i, 1)))
    Next i
End Function



' extract major/minor from version cap
Function D3DSHADER_VERSION_MAJOR(Version As Long) As Long

    D3DSHADER_VERSION_MAJOR = (((Version) \ 8) And &HFF&)

End Function

Function D3DSHADER_VERSION_MINOR(Version As Long) As Long

    D3DSHADER_VERSION_MINOR = (((Version)) And &HFF&)

End Function

'vertex shader version token
Function D3DVS_VERSION(Major As Long, Minor As Long) As Long

    D3DVS_VERSION = (&HFFFE0000 Or ((Major) * 2 ^ 8) Or (Minor))

End Function

Function LONGtoD3DCOLORVALUE(ByVal Color As Long) As D3DCOLORVALUE

  Dim a As Long, r As Long, g As Long, b As Long

    If Color < 0 Then
        a = ((Color And (&H7F000000)) / (2 ^ 24)) Or &H80&
      Else
        a = Color / (2 ^ 24)
    End If
    r = (Color And &HFF0000) / (2 ^ 16)
    g = (Color And &HFF00&) / (2 ^ 8)
    b = (Color And &HFF&)

    LONGtoD3DCOLORVALUE.a = a / 255
    LONGtoD3DCOLORVALUE.r = r / 255
    LONGtoD3DCOLORVALUE.g = g / 255
    LONGtoD3DCOLORVALUE.b = b / 255

End Function

Private Sub Get_Capabilities()
'cfnc = fnc.E_Get_Capabilities
'Marce On error resume next
    Dim d3dCaps As D3DCAPS8, i As Integer, j As Integer

    D3DDevice.GetDeviceCaps d3dCaps

    'check bump mapping

    ''//Does this device support the two bump mapping blend operations?
    If (d3dCaps.TextureOpCaps And D3DTEXOPCAPS_BUMPENVMAPLUMINANCE) Then
        act_caps.CanDo_BumpMapping = 1
    End If

    ''//Does this device support up to three blending stages?
    If d3dCaps.MaxTextureBlendStages < 3 Then
        act_caps.CandDo_3StagesTextureBlending = 0
      Else
        act_caps.CandDo_3StagesTextureBlending = 1

    End If

    ''//Does this device support multitexturing
    If d3dCaps.MaxSimultaneousTextures > 1 Then
        act_caps.CanDo_MultiTexture = 1
        act_caps.Max_TextureStages = d3dCaps.MaxSimultaneousTextures
    End If

    'anisotropic filter
    If d3dCaps.RasterCaps And D3DPRASTERCAPS_ANISOTROPY Then
        act_caps.Filter_Anisotropic = True

        act_caps.Max_AnisotropY = d3dCaps.MaxAnisotropy

    End If

    'trilinear

    If (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MINFLINEAR) Then

        If (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MAGFLINEAR) Then
            If (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MIPFLINEAR) Then

                act_caps.Filter_Trilinear = 1

            End If
        End If
    End If

    'flatcubic

    If ((d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MINFLINEAR) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MAGFAFLATCUBIC) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MIPFLINEAR)) Then

        act_caps.Filetr_FlatCubic = 1

    End If

    'Gaussian cubic

    If ((d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MINFLINEAR) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MAGFGAUSSIANCUBIC) + _
       (d3dCaps.TextureFilterCaps And D3DPTFILTERCAPS_MIPFLINEAR)) Then

        act_caps.Filter_GaussianCubic = 1

    End If

    If d3dCaps.TextureCaps And D3DPTEXTURECAPS_VOLUMEMAP Then

        act_caps.CanDo_VolumeTexture = 1

    End If

    If d3dCaps.TextureCaps And D3DPTEXTURECAPS_PROJECTED Then

        act_caps.CanDo_ProjectedTexture = 1

    End If

    If d3dCaps.TextureCaps And D3DPTEXTURECAPS_MIPMAP Then

        act_caps.CanDo_TextureMipMapping = 1

    End If

'    If (d3dCaps.RasterCaps And D3DPRASTERCAPS_WBUFFER) Then
'        act_caps.Wbuffer_OK = True
'        obj_Device.SetRenderState D3DRS_ZENABLE, D3DZB_USEW
'        IS_WBUFFER = True
'    End If

    If d3dCaps.MaxPointSize > 0 Then
        act_caps.CanDo_PointSprite = 1
    End If

  Dim MA As Long
  Dim MI As Long

    MA = D3DSHADER_VERSION_MAJOR(d3dCaps.VertexShaderVersion)
    MI = D3DSHADER_VERSION_MINOR(d3dCaps.VertexShaderVersion)

    'MA = D3DVS_VERSION(MA, MI)
    act_caps.Vertex_ShaderVERSION = Str(MI) + "." + CStr(MA)

    MA = D3DSHADER_VERSION_MAJOR(d3dCaps.PixelShaderVersion)
    MI = D3DSHADER_VERSION_MINOR(d3dCaps.PixelShaderVersion)
    act_caps.pxs_max = MA
    act_caps.pxs_min = MI
    'MA = D3DVS_VERSION(MA, MI)
    act_caps.Pixel_ShaderVERSIOn = Str(MI) + "." + CStr(MA)

    act_caps.Cando_VertexShader = d3dCaps.VertexShaderVersion >= D3DVS_VERSION(1, 0)
    act_caps.Cando_PixelShader = d3dCaps.PixelShaderVersion >= D3DVS_VERSION(1, 0)

    act_caps.CanDo_CubeMapping = (d3dCaps.TextureCaps And D3DPTEXTURECAPS_CUBEMAP)

    act_caps.CanDo_Dot3 = (d3dCaps.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3)
    
    usaBumpMapping = usaBumpMapping And act_caps.CanDo_Dot3 And act_caps.CandDo_3StagesTextureBlending

    act_caps.CanDoTableFog = (d3dCaps.RasterCaps And D3DPRASTERCAPS_FOGTABLE) And _
                              (D3DPRASTERCAPS_ZFOG) Or (d3dCaps.RasterCaps And D3DPRASTERCAPS_WFOG)

    act_caps.CanDoVertexFog = (d3dCaps.RasterCaps And D3DPRASTERCAPS_FOGVERTEX)

    act_caps.CanDoWFog = (d3dCaps.RasterCaps And D3DPRASTERCAPS_WFOG)

'  Dim nAdapters As Long 'How many adapters we found
'  Dim AdapterInfo As D3DADAPTER_IDENTIFIER8 'A Structure holding information on the adapter
'
'  Dim sTemp As String
'
'    '//This'll either be 1 or 2
'    nAdapters = obj_D3D.GetAdapterCount
'
'    For I = 0 To nAdapters - 1
'        'Get the relevent Details
'        obj_D3D.GetAdapterIdentifier I, 0, AdapterInfo
'
'        'Get the name of the current adapter - it's stored as a long
'        'list of character codes that we need to parse into a string
'        ' - Dont ask me why they did it like this; seems silly really :)
'        sTemp = "" 'Reset the string ready for our use
'
'        For J = 0 To 511
'            sTemp = sTemp & Chr$(AdapterInfo.Description(J))
'        Next J
'        sTemp = Replace(sTemp, Chr$(0), " ")
'        J = InStr(sTemp, "     ")
'        sTemp = Left$(sTemp, J)
'
'    Next I
'
'    If InStr(UCase(sTemp), "GEFORCE") Then
'
'        If act_caps.Wbuffer_OK = 0 Then
'            act_caps.Wbuffer_OK = 1
'            IS_WBUFFER = 1
'        End If
'
'    End If

End Sub

Sub Engine_Weather_UpdateFog()
'Update the fog effects

Dim i As Long
Dim X As Long
Dim Y As Long
Dim c As Long
Dim tx!, ty!


    If WeatherFogCount = 0 Then WeatherFogCount = 13

    WeatherFogX1 = WeatherFogX1 + (timerElapsedTime * (0.018 + Rnd * 0.01)) - (offset_mapO.X - offset_map.X)
    WeatherFogY1 = WeatherFogY1 + (timerElapsedTime * (0.013 + Rnd * 0.01)) - (offset_mapO.Y - offset_map.Y)
    
'    Engine_Parallax.GetParalllaxOffset WeatherFogX1, WeatherFogY1, 200, TX, TY
'
'    TX = (TX * 3) Mod 512
'    TY = (TY * 3) Mod 512
    
    WeatherFogX1 = WeatherFogX1 - tx
    WeatherFogY1 = WeatherFogY1 - ty
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop
    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop
    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop
    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (timerElapsedTime * (0.037 + Rnd * 0.01)) - (offset_mapO.X - offset_map.X)
    WeatherFogY2 = WeatherFogY2 - (timerElapsedTime * (0.021 + Rnd * 0.01)) - (offset_mapO.Y - offset_map.Y)
    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop
    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop
    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop
    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop


    
    'Render fog 2
    X = 2
    Y = -1
    
    Weatherfogalpha = AlphaNiebla
    
    c = D3DColorARGB(Weatherfogalpha, 255, 255, 255)
    For i = 1 To WeatherFogCount
        Grh_Render_Simple_box 1127, (X * 512) + WeatherFogX2, (Y * 512) + WeatherFogY2, c, 512!
        X = X + 1
        If X > (1 + (MainViewWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i
            
            

    
    'Render fog 1
    If Weatherfogalpha < 40 Then Exit Sub
    X = 0
    Y = 0
    c = D3DColorARGB(Weatherfogalpha - 40, 255, 255, 255)
    For i = 1 To WeatherFogCount
        Grh_Render_Simple_box 1128, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, c, 512!
        X = X + 1
        If X > (2 + (MainViewWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i
    
    WeatherFogX1 = WeatherFogX1 + tx
    WeatherFogY1 = WeatherFogY1 + ty
End Sub

Sub Engine_Weather_SandStorm()
'Update the fog effects

Dim i As Long
Dim X As Long
Dim Y As Long
Dim c As Long


    If Sandstorm_Count = 0 Then Sandstorm_Count = 13

    Sandstorm_X1 = Sandstorm_X1 + (timerElapsedTime * 0.4) - (offset_mapO.X - offset_map.X)
    Sandstorm_Y1 = Sandstorm_Y1 - (offset_mapO.Y - offset_map.Y) '+ (timerElapsedTime * 0.05 * Rnd) - (timerElapsedTime * 0.05 * Rnd)
    Do While Sandstorm_X1 < -1024
        Sandstorm_X1 = Sandstorm_X1 + 1024
    Loop
    Do While Sandstorm_Y1 < -512
        Sandstorm_Y1 = Sandstorm_Y1 + 512
    Loop
    Do While Sandstorm_X1 > 0
        Sandstorm_X1 = Sandstorm_X1 - 1024
    Loop
    Do While Sandstorm_Y1 > 0
        Sandstorm_Y1 = Sandstorm_Y1 - 512
    Loop
    

    'Render fog 1

    X = 0
    Y = 0
    c = D3DColorARGB(Engine_Meteorologic.AlphaArena, 255, 255, 255)
    For i = 1 To Sandstorm_Count
        Grh_Render_Simple_box 1132, (X * 1024) + Sandstorm_X1, (Y * 512) + Sandstorm_Y1, c, 512!, , 512!
        X = X + 1
        If X > (2 + (MainViewWidth \ 1024)) Then
            X = 0
            Y = Y + 1
        End If
    Next i

End Sub

Sub Engine_Weather_lightbeam()
On Error GoTo endaa
Dim i As Long
Dim X As Long
Dim Y As Long
Dim c As Long

    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Randomize
    
    i = Lightbeam_a1
    i = Abs(CInt(i + Rnd * 2 - Rnd * 2)) Mod 60
    Lightbeam_a1 = i

    c = D3DColorARGB(i, 255, 255, 255)
    Grh_Render_Simple_box 1129, 256, 0, c, 512!



    i = Lightbeam_a2
    i = Abs(CInt(i + Rnd * 2 - Rnd * 2)) Mod 64
    Lightbeam_a2 = i

    c = D3DColorARGB(i, 255, 255, 255)
    Grh_Render_Simple_box 1129, 0, 0, c, 512!
    
    
    
    i = Lightbeam_a3
    i = Abs(CInt(i + Rnd * 2 - Rnd * 2)) Mod 128
    Lightbeam_a3 = i

    c = D3DColorARGB(i, 255, 255, 255)
    Grh_Render_Simple_box 1130, 0, 0, c, 512!
    


    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    Exit Sub
endaa:
Debug.Assert False
End Sub




Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
If GrhIndex = 0 Then Exit Sub
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started

        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

'MARCE EH? Estan bien los ifs? debe ser < o <= MZ: Si, estan bien
Public Sub Colorear_TBOX(ByRef Box As Box_Vertex, ByVal X As Byte, ByVal Y As Byte)
    Box.color0 = ResultColorArray(X, Y)
    If X < X_MAXIMO_VISIBLE Then
        Box.Color2 = ResultColorArray(X + 1, Y)
        If Y > Y_MINIMO_VISIBLE Then
            Box.color3 = ResultColorArray(X + 1, Y - 1)
            Box.Color1 = ResultColorArray(X, Y - 1)
        End If
    End If
End Sub

'MARCE EH?
Public Sub Colorear_TBOX_Flip(ByRef Box As Box_Vertex, ByVal X As Byte, ByVal Y As Byte)
    Box.Color2 = ResultColorArray(X, Y)
    If X < X_MAXIMO_VISIBLE Then
        Box.color3 = ResultColorArray(X + 1, Y)
        If Y > Y_MINIMO_VISIBLE Then
            Box.Color1 = ResultColorArray(X + 1, Y - 1)
            Box.color0 = ResultColorArray(X, Y - 1)
        End If
    End If
End Sub



'Public Sub Grh_Render_relieve(ByVal GrhIndex As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte, ByVal flip As Byte)
''*********************************************
''Author: menduz
''*********************************************
'    Dim tBottom!, tRight! ', tTop!, tLeft!
'    Dim ll As Long
'    Dim TGRH As GrhData
'    Dim altU As AUDT
'
'    Dim colores(0 To 3) As Long
'
'
'    If GrhIndex = 0 Then Exit Sub
'    call GetTexture(GrhData(GrhIndex).FileNum)
'
'    If copy_tile_now Or MapData(map_x, map_y).tile_render = 0 Or MapData(map_x, map_y).is_water = True Then
'        If GrhData(GrhIndex).hardcor = 0 Then Init_grh_tutv GrhIndex
'        TGRH = GrhData(GrhIndex)
'        altU = hMapData(map_x, map_y)
'
'        tBottom = tTop + TGRH.pixelHeight
'        tRight = tLeft + TGRH.pixelWidth
'
'
'        With tBox 'With tBox
'            If flip Then
'                .x0 = tLeft
'                .y0 = tTop - altU.hs(1)
'                .color0 = colores(1)
'                .x1 = tRight
'                .y1 = tTop - altU.hs(3)
'                .color1 = colores(3)
'                .x2 = tLeft
'                .y2 = tBottom - altU.hs(0)
'                .color2 = colores(0)
'                .x3 = tRight
'                .y3 = tBottom - altU.hs(2)
'                .color3 = colores(2)
'
'                .tu0 = TGRH.tu(1)
'                .tv0 = TGRH.tv(1)
'                .tu1 = TGRH.tu(3)
'                .tv1 = TGRH.tv(3)
'                .tu2 = TGRH.tu(0)
'                .tv2 = TGRH.tv(0)
'                .tu3 = TGRH.tu(2)
'                .tv3 = TGRH.tv(2)
'            Else
'                .x0 = tLeft
'                .y0 = tBottom - altU.hs(0)
'                .color0 = colores(0)
'                .x1 = tLeft
'                .y1 = tTop - altU.hs(1)
'                .color1 = colores(1)
'                .x2 = tRight
'                .y2 = tBottom - altU.hs(2)
'                .color2 = colores(2)
'                .x3 = tRight
'                .y3 = tTop - altU.hs(3)
'                .color3 = colores(3)
'
'                .tu0 = TGRH.tu(0)
'                .tv0 = TGRH.tv(0)
'                .tu1 = TGRH.tu(1)
'                .tv1 = TGRH.tv(1)
'                .tu2 = TGRH.tu(2)
'                .tv2 = TGRH.tv(2)
'                .tu3 = TGRH.tu(3)
'                .tv3 = TGRH.tv(3)
'            End If
'
'            If ModSuperWater(map_x, map_y) Then
'                .y0 = .y0 + ModSuperWaterDD(map_x, map_y).hs(0)
'                .y1 = .y1 + ModSuperWaterDD(map_x, map_y).hs(1)
'                .y2 = .y2 + ModSuperWaterDD(map_x, map_y).hs(2)
'                .y3 = .y3 + ModSuperWaterDD(map_x, map_y).hs(3)
'            End If
'
'
'        End With
'
'        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
'        'batch_add_box tBox, GrhData(GrhIndex).FileNum, actual_blend_mode
'
'        'MapData(map_x, map_y).tile_render = 255
'        If copy_tile_now Or MapData(map_x, map_y).tile_render = 0 Then
'            CopyMemory MapData(map_x, map_y).tile, tBox, BV_size
'            MapData(map_x, map_y).tile_render = 255
'        End If
'    Else
'        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MapData(map_x, map_y).tile, TL_size
'        'batch_add_box MapData(map_x, map_y).tile, GrhData(GrhIndex).FileNum, actual_blend_mode
'    End If
'
'End Sub


Public Function IsIDE() As Boolean
On Error GoTo is_ide
Debug.Print 1 / 0
Exit Function
is_ide:
IsIDE = True
End Function



'Public Sub Grh_Render_Tileset(ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte, ByVal flip As Byte)
''*********************************************
''Author: menduz
''*********************************************
'    Dim tBottom!, tRight! ', tTop!, tLeft!
'
'    Dim tn As Byte
'
'    Dim altU As AUDT
'
'    Dim H!, W!, SX!, SY!
'
'    altU = hMapData(map_x, map_y)
'
'    tn = MapData(map_x, map_y).tile_number
'    If tn = 0 Then Exit Sub
'
'    Call GetTexture(MapData(map_x, map_y).tile_texture)
'    Call GetTextureDimension(MapData(map_x, map_y).tile_texture, H, W)
'    SX = (MapData(map_x, map_y).tile_number Mod (W \ 32)) * 32!
'    SY = (MapData(map_x, map_y).tile_number \ (H \ 32)) * 32!
'
'    Colorear_TBOX Tileset_Grh_Array(tn), map_x, map_y
'
'    tBottom = tTop + 32!
'    tRight = tLeft + 32!
'
'    With Tileset_Grh_Array(tn)
'        .x0 = tLeft
'        .y0 = tBottom - altU.hs(0)
'
'        .x1 = tLeft
'        .y1 = tTop - altU.hs(1)
'
'        .x2 = tRight
'        .y2 = tBottom - altU.hs(2)
'
'        .x3 = tRight
'        .y3 = tTop - altU.hs(3)
'    End With
'
'    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Tileset_Grh_Array(tn), TL_size
'
'End Sub



Public Sub Engine_Render_TBox(Box As Box_Vertex, ByVal textura As Integer, Optional ByVal BoxCount As Integer = 1)
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(textura)
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    If BoxCount = 1 Then
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Box, TL_size
    ElseIf BoxCount > 1 Then
        If IndexBufferEnabled Then
            D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLELIST, 0, BoxCount * 4, BoxCount * 2, StaticIndexBuffer(0), D3DFMT_INDEX16, Box, TL_size
        Else
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, BoxCount * 2, Box, TL_size
        End If
    End If
End Sub



Public Function FloatToColor(a!, r!, g!, b!) As Long
FloatToColor = D3DColorARGB(a * 255, r * 255, g * 255, b * 255)
End Function

Public Function Colision(ByVal X As Integer, ByVal Y As Integer, ByVal top As Integer, ByVal bottom As Integer, ByVal left As Integer, ByVal right As Integer) As Boolean
If X >= left Then
    If X <= right Then
        If Y >= top Then
            If Y <= bottom Then
                Colision = True
            End If
        End If
    End If
End If
End Function

Public Function ColisionRect(ByVal X As Integer, ByVal Y As Integer, ByRef r As RECT) As Boolean
If X >= r.left Then
    If X <= r.right Then
        If Y >= r.top Then
            If Y <= r.bottom Then
                ColisionRect = True
            End If
        End If
    End If
End If
End Function

Public Function GetVertexShader() As Long
    GetVertexShader = pVertexShader 'Engine.GetVertexShader()
End Function

Public Sub SetVertexShader(ByVal vs As Long)
'    If vs = pVertexShader Then Exit Sub
    D3DDevice.SetVertexShader vs
    pVertexShader = vs
End Sub


Public Function CreateShaderFromCode(PixelShader As String) As Long
    Dim shaderCode As D3DXBuffer
    Dim RetError As String
    Dim Arr() As Long
    Dim size As Long
    Dim i As Long
    If Not act_caps.Cando_PixelShader Then Exit Function
On Error GoTo exitt
    Set shaderCode = D3DX.AssembleShader(PixelShader, 0, Nothing)
    size = shaderCode.GetBufferSize() / 4
    ReDim Arr(size - 1)
    D3DX.BufferGetData shaderCode, 0, 4, size, Arr(0)
    CreateShaderFromCode = D3DDevice.CreatePixelShader(Arr(0))
    
    If CreateShaderFromCode Then
        Debug.Print "Pixel shader generado"
    Else
        Debug.Print "Error en el pixel shader"
    End If
    If PixelShader = "" Then
        Debug.Print "Pixel shader vacio"
    End If
    
    Set shaderCode = Nothing
    
    Exit Function
exitt:
    Beep
    CreateShaderFromCode = 0
    Debug.Print D3DX.GetErrorString(Err.Number)
    Err.Clear
End Function

