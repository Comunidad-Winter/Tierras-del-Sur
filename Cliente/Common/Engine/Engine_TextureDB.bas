Attribute VB_Name = "Engine_TextureDB"
' ESTE ARCHIVO ESTA COMPARTIDO POR TODOS LOS PROGRAMAS.

' TODO: Eliminar la tabla de indices. Usar en cambio un indice para cada textura y un flag bCargado

Option Explicit

Private Type TEXT_DB_ENTRY
    FileName        As Integer
    UltimoAcceso    As Long
    texture         As Direct3DTexture8
    alto            As Single
    ancho           As Single
    size            As Long
    png             As Byte
    
    complemento_1   As Integer
    complemento_2   As Integer
    complemento_3   As Integer
    complemento_4   As Integer
End Type

Private Tabla() As Integer
Private TablaMax As Integer

Private mGraficos() As TEXT_DB_ENTRY

Private mMaxEntries As Integer
Private mCantidadGraficos As Integer
Private mFreeMemoryBytes As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Enum TEX_FLAGS
    TEX_NOUSE_MIPMAPS = 1
End Enum

Public pakGraficos As clsEnpaquetado

Public last_texture As Integer

Public Sub DeInit_TextureDB()
'Marce On error resume next
    Dim i As Long
    
    For i = 1 To mCantidadGraficos
        Set mGraficos(i).texture = Nothing
    Next i
    
    Erase mGraficos
    
End Sub

Public Sub ReloadAllTextures()
    ReDim Tabla(TablaMax)
End Sub

Public Function PeekTexture(ByVal FileName As Integer) As Direct3DTexture8
    Dim Index As Long

    If FileName = 0 Then
        Exit Function
    End If

    If TablaMax < FileName Then
        TablaMax = FileName + 128
        ReDim Preserve Tabla(TablaMax)
    End If
    
    Index = Tabla(FileName)

    If Index <> 0 Then
        mGraficos(Index).UltimoAcceso = SceneBegin
        Set PeekTexture = mGraficos(Index).texture
    Else

    #Const MEM_VIDEO = False
    #If MEM_VIDEO = True Then
        If mMaxEntries = mCantidadGraficos Or mFreeMemoryBytes < 4000000 Then '~4kb
    #Else
        If mMaxEntries = mCantidadGraficos Then
    #End If
            'Sacamos el que hace más que no usamos, y utilizamos el slot
            Index = CrearGrafico(FileName, BorraMenosUsado())
            Set PeekTexture = mGraficos(Index).texture
        Else
            'Agrego una textura nueva a la lista
            Index = CrearGrafico(FileName)
            Set PeekTexture = mGraficos(Index).texture
        End If
        Tabla(FileName) = Index
    End If
End Function




Public Sub PreLoadTexture(ByVal FileName As Integer)
Dim Index As Integer
If FileName > 0 Then
    If TablaMax < FileName Then
        ReDim Preserve Tabla(FileName)
    End If
    
    If Tabla(FileName) = 0 Then
        If mMaxEntries = mCantidadGraficos Then
            'Sacamos el que hace más que no usamos, y utilizamos el slot
            Index = CrearGrafico(FileName, BorraMenosUsado())
        Else
            'Agrego una textura nueva a la lista
            Index = CrearGrafico(FileName)
        End If
        Tabla(FileName) = Index
    End If
End If
End Sub

Public Sub GetTextureDimension(ByVal FileName As Integer, ByRef h As Single, ByRef w As Single)
    Dim Index As Integer
    Index = Tabla(FileName)
    If Index Then
        h = mGraficos(Index).alto
        w = mGraficos(Index).ancho
    End If
End Sub

Public Function GetTexturePNG(ByVal FileName As Integer) As Byte
    Dim Index As Integer
    Index = Tabla(FileName)
    If Index Then
        GetTexturePNG = mGraficos(Index).png
    End If
End Function

Public Function Init_TextureDB(ByVal MaxMemory As Long, ByVal MaxEntries As Long, path_Pack As String) As Boolean
    mMaxEntries = MaxEntries

    If mMaxEntries < 1 Then 'por lo menos 1 gráfico
        Exit Function
    End If

    mCantidadGraficos = 0

    mFreeMemoryBytes = MaxMemory

    Init_TextureDB = True

    'mFreeMemoryBytes = D3DDevice.GetAvailableTextureMem(D3DPOOL_MANAGED)

    ReDim Tabla(32767)
    TablaMax = 32767

    Set pakGraficos = New clsEnpaquetado
    pakGraficos.Cargar path_Pack
End Function

Public Sub Borrar_TextureDB()
    Dim i As Long
    
    For i = 1 To mCantidadGraficos
        Set mGraficos(i).texture = Nothing
    Next i
    ReDim Tabla(3000)
    TablaMax = 3000
    ReDim mGraficos(0)
    mCantidadGraficos = 0
End Sub

Private Function CrearGrafico(ByVal archivo As Integer, Optional ByVal Index As Integer = -1) As Integer
'On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim srcData() As Byte
    Dim header As Long
    Dim fmt1 As CONST_D3DFORMAT, fmt2 As CONST_D3DFORMAT
    Dim bUseMip As Long
    
    If Index < 0 Then
        Index = mCantidadGraficos + 1
        ReDim Preserve mGraficos(1 To Index)
    End If
    Err.Clear
    
    If Index = 0 Then Index = mCantidadGraficos
    
    With mGraficos(Index)
        .FileName = archivo
        .UltimoAcceso = GetTimer
        
    On Local Error Resume Next
        Dim IH As INFOHEADER
'Debug.Assert Archivo < 20000
        If pakGraficos.IH_Get(archivo, IH) And pakGraficos.Leer(archivo, srcData(), rGrh) Then
            .png = (IH.file_type = eTiposRecursos.rPng)
            
            .complemento_1 = IH.complemento_1
            .complemento_2 = IH.complemento_2
            .complemento_3 = IH.complemento_3
            .complemento_4 = IH.complemento_4

'            Debug.Assert .complemento_1 = 0
            DXCopyMemory header, srcData(0), 4
            
            If header = &H20534444 Then 'DDS magic header
                fmt1 = D3DFMT_UNKNOWN
                fmt2 = D3DFMT_A8R8G8B8
            Else
                fmt1 = D3DFMT_A8R8G8B8
                fmt2 = D3DFMT_UNKNOWN
            End If
            
            If IH.flags And TEX_NOUSE_MIPMAPS Then
                bUseMip = 0
            Else
                bUseMip = 1
            End If
            
            Set .texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, srcData(0), UBound(srcData) + 1, _
                    D3DX_DEFAULT, D3DX_DEFAULT, bUseMip, 0, fmt1, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                    D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
                    
            If .texture Is Nothing Then
                Err.Clear
                Set .texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, srcData(0), UBound(srcData) + 1, _
                    D3DX_DEFAULT, D3DX_DEFAULT, bUseMip, 0, fmt2, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                    D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
            End If
            Erase srcData
            If Not .texture Is Nothing Then
            .texture.GetLevelDesc 0, surface_desc
            End If
        Else
            Set .texture = Nothing
        End If
        
    On Local Error GoTo 0
    
        If Err.Number Or .texture Is Nothing Then
            LogError "A5.0 Error en carga de gráficos[" & archivo & "]. - " & D3DX.GetErrorString(Err.Number)
            Set .texture = Nothing
        End If
        
        .ancho = surface_desc.Width
        .alto = surface_desc.Height
        .size = surface_desc.size
            
        mFreeMemoryBytes = D3DDevice.GetAvailableTextureMem(D3DPOOL_MANAGED) 'mFreeMemoryBytes - surface_desc.size
    End With
    
    mCantidadGraficos = mCantidadGraficos + 1
    
    CrearGrafico = Index
Exit Function
errHandler:

LogError "A5.0 Error en carga de gráficos.  - " & D3DX.GetErrorString(Err.Number)


End Function

Private Function BorraMenosUsado() As Integer
    Dim valor As Long
    Dim i As Long
    'Inicializamos todo
    valor = GetTimer 'mGraficos(1).UltimoAcceso
    BorraMenosUsado = 1
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To mCantidadGraficos
        If mGraficos(i).UltimoAcceso < valor Then
            valor = mGraficos(i).UltimoAcceso
            BorraMenosUsado = i
        End If
    Next i
    'Disminuimos el contador
    mCantidadGraficos = mCantidadGraficos - 1
    'Borramos la texture
    
    Set mGraficos(BorraMenosUsado).texture = Nothing
    Tabla(mGraficos(BorraMenosUsado).FileName) = 0
    mGraficos(BorraMenosUsado).alto = 0
    mGraficos(BorraMenosUsado).ancho = 0
    mFreeMemoryBytes = mFreeMemoryBytes + mGraficos(BorraMenosUsado).size
    mGraficos(BorraMenosUsado).size = 0
End Function

Public Sub BorrarTexturaDeMemoria(ByVal numero As Integer)
If numero <= TablaMax Then
    Tabla(numero) = 0
End If
End Sub


Public Property Get MaxEntries() As Integer
    MaxEntries = mMaxEntries
End Property

Public Property Let MaxEntries(ByVal vNewValue As Integer)
    mMaxEntries = vNewValue
End Property

Public Property Get cantidadGraficos() As Integer
    cantidadGraficos = mCantidadGraficos
End Property

Public Function definir_complementarios(FileName As Integer, C1 As Integer, C2 As Integer, Optional C3 As Integer = -1, Optional C4 As Integer = -1) As Boolean
    definir_complementarios = False
    If Tabla(FileName) <> 0 Then
        mGraficos(Tabla(FileName)).complemento_1 = C1
        mGraficos(Tabla(FileName)).complemento_2 = C2
        If C3 <> -1 Then mGraficos(Tabla(FileName)).complemento_3 = C3
        If C4 <> -1 Then mGraficos(Tabla(FileName)).complemento_3 = C4
        definir_complementarios = True
    End If
End Function

Public Function Obtener_Texturas_Complementarias(FileName As Integer, ByRef C1 As Integer, ByRef C2 As Integer, Optional ByRef C3 As Integer, Optional ByRef C4 As Integer) As Boolean
    Obtener_Texturas_Complementarias = False
    Dim ElemEnTabla As Long
    
    ElemEnTabla = Tabla(FileName)
    
    If ElemEnTabla <> 0 Then
        C1 = mGraficos(ElemEnTabla).complemento_1
        C2 = mGraficos(ElemEnTabla).complemento_2
        C3 = mGraficos(ElemEnTabla).complemento_3
        C4 = mGraficos(ElemEnTabla).complemento_4
        
        Obtener_Texturas_Complementarias = (C1 > 0) Or (C2 > 0)
    End If
End Function


