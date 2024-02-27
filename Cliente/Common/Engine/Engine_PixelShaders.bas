Attribute VB_Name = "Engine_PixelShaders"
Option Explicit

Public Const MAX_PIXEL_SHADERS = 7

Public Enum ePixelShaders
    Ninguno = 0
    estandar = 1
    Agua = 2
    Particulas = 3
    
    ' DEBUG
    Normales = 4
    ColoresLuces = 5
    ColoresAmbiente = 6
    
    Pisos = 7
End Enum

Private PixelShaderActual As Integer

Private UltimaTextura1  As Direct3DTexture8 ' Colores (diffuse)
Private UltimaTextura2  As Direct3DTexture8 ' Normal de los colores (normal map)
Private UltimaTextura3  As Direct3DTexture8 ' Normal de las luces (ambient map)
Private UltimaTextura4  As Direct3DTexture8 ' Complemento 2 (emissive)

Public Type tPixelShader
    heapRef             As Long
    tipo                As ePixelShaders
    codigo              As String
    vertexShader        As Long
    codigoVertexShader  As String
    FVF                 As Long
End Type

Public PixelShaderCatalog(0 To MAX_PIXEL_SHADERS) As tPixelShader

Public Sub Engine_PixelShaders_EngineReiniciado()
    Set UltimaTextura1 = Nothing
    Set UltimaTextura2 = Nothing
    Set UltimaTextura3 = Nothing
    Set UltimaTextura4 = Nothing
    
    PixelShaderActual = -1
    
    Dim i As Integer
    
    For i = 0 To MAX_PIXEL_SHADERS
        PixelShaderCatalog(i).vertexShader = 0
        PixelShaderCatalog(i).heapRef = 0
        
        If PixelShaderCatalog(i).codigoVertexShader <> vbNullString Then
            PixelShaderCatalog(i).vertexShader = 0
            CompilarVertexShader (i)
        End If
        
        If PixelShaderCatalog(i).codigo <> vbNullString Then
            PixelShaderCatalog(i).heapRef = 0
            CompilarPixelShader (i)
        End If
    Next i
End Sub

Public Sub Engine_PixelShaders_Iniciar()
    PixelShaderActual = -1
End Sub


Private Function creaVShader(codigo As String, FVF As Long) As Long
On Error GoTo fin
    Dim shaderArray() As Long
    Dim shader As Long

    Dim shaderCode As D3DXBuffer
    Dim errors As String
    
    errors = Space$(255)
    
    Dim decl As D3DXDECLARATOR
 
    D3DX.DeclaratorFromFVF FVF, decl

    Set shaderCode = D3DX.AssembleShader(codigo, 0, Nothing, errors)
 
    ReDim shaderArray(shaderCode.GetBufferSize() / 4)
    
    D3DX.BufferGetData shaderCode, 0, 1, shaderCode.GetBufferSize(), shaderArray(0)
 
    Set shaderCode = Nothing
 
    D3DDevice.CreateVertexShader decl.value(0), shaderArray(0), shader, 0
    creaVShader = shader
    
    LogDebug "VS: OK!"
    
    Exit Function
fin:
    Call LogDebug("VS: ERROR! " + errors)
End Function


Public Sub Engine_PixelShaders_Setear( _
    ByVal cualPixelShader As ePixelShaders, _
    Optional codigo As String = vbNullString, _
    Optional FVF As Long = 0, _
    Optional codigoVertexShader As String = vbNullString)
    
    With PixelShaderCatalog(cualPixelShader)
        .codigo = codigo
        .codigoVertexShader = codigoVertexShader
        .vertexShader = 0
        .heapRef = 0
        .FVF = FVF
    End With
    
    If PixelShaderCatalog(cualPixelShader).codigoVertexShader <> vbNullString Then
        PixelShaderCatalog(cualPixelShader).vertexShader = 0
        CompilarVertexShader cualPixelShader
    End If
    
    If PixelShaderCatalog(cualPixelShader).codigo <> vbNullString Then
        PixelShaderCatalog(cualPixelShader).heapRef = 0
        CompilarPixelShader cualPixelShader
    End If
    
    If PixelShaderActual = cualPixelShader Then PixelShaderActual = -1
End Sub

Private Sub CompilarPixelShader(ByVal cualPixelShader As ePixelShaders)
    PixelShaderCatalog(cualPixelShader).heapRef = CreateShaderFromCode(PixelShaderCatalog(cualPixelShader).codigo)
End Sub

Private Sub CompilarVertexShader(ByVal cualPixelShader As ePixelShaders)
    PixelShaderCatalog(cualPixelShader).vertexShader = creaVShader(PixelShaderCatalog(cualPixelShader).codigoVertexShader, PixelShaderCatalog(cualPixelShader).FVF)
End Sub

Public Sub Engine_PixelShaders_Utilizar(ByVal cualPixelShader As ePixelShaders)
    If PixelShaderActual = cualPixelShader Then Exit Sub
    With PixelShaderCatalog(cualPixelShader)
        If .heapRef Then
            D3DDevice.SetPixelShader .heapRef
        Else
            D3DDevice.SetPixelShader 0
        End If
        
        If .vertexShader Then
        On Error GoTo errorShader
            D3DDevice.SetVertexShader .vertexShader
            GoTo shaderOK
errorShader:
            .vertexShader = 0
            Call LogDebug("Borro el shader por error!")
shaderOK:
        Else
            D3DDevice.SetVertexShader .FVF
        End If
    End With
    PixelShaderActual = cualPixelShader
End Sub

Public Sub Engine_PixelShaders_SetTexture_Diffuse(ByRef tex As Direct3DTexture8)
    If UltimaTextura1 Is tex Then Exit Sub
    D3DDevice.SetTexture 0, tex
    Set UltimaTextura1 = tex
End Sub

Public Sub Engine_PixelShaders_SetTexture_Normal(ByRef tex As Direct3DTexture8)
    If UltimaTextura2 Is tex Then Exit Sub
    D3DDevice.SetTexture 1, tex
    Set UltimaTextura2 = tex
End Sub

Public Sub Engine_PixelShaders_SetTexture_Ambient(ByRef tex As Direct3DTexture8)
    If UltimaTextura3 Is tex Then Exit Sub
    D3DDevice.SetTexture 2, tex
    Set UltimaTextura3 = tex
End Sub

Public Sub Engine_PixelShaders_SetTexture_Emissive(ByRef tex As Direct3DTexture8)
    If UltimaTextura4 Is tex Then Exit Sub
    D3DDevice.SetTexture 3, tex
    Set UltimaTextura4 = tex
End Sub
