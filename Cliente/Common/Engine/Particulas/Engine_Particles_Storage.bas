Attribute VB_Name = "Engine_Particles_Storage"
Option Explicit

Public GlobalParticleGroup()        As Engine_Particle_Group
Public GlobalParticleGroupCount     As Integer

Public Const PartGroupNONE          As Integer = -1



Public Function PartGroupValid(ByVal Index As Integer) As Boolean
    PartGroupValid = Index <> PartGroupNONE And Index <= GlobalParticleGroupCount And Index > 0
End Function

Public Function IniciarGrupoParticulas(ByVal PGID As Integer, ByRef grupo As Engine_Particle_Group) As Boolean
    If PartGroupValid(PGID) Then
        If Not GlobalParticleGroup(PGID) Is Nothing Then
            GlobalParticleGroup(PGID).ClonarEn grupo
            IniciarGrupoParticulas = True
        End If
    End If
End Function

Public Function PersistirGruposParticulas(ByVal Path As String) As Boolean
'On Error GoTo ErrHandler
    If FileExist(Path) Then
        Kill Path
    End If

    Dim handle As Integer
    Dim i As Integer
    
    handle = FreeFile
    
    Open Path For Binary Access Write As handle
        Put handle, , GlobalParticleGroupCount
        If GlobalParticleGroupCount > 0 Then
            For i = 1 To GlobalParticleGroupCount
                GlobalParticleGroup(i).PGID = 0
                GlobalParticleGroup(i).persistir handle
            Next i
        End If
    Close handle
    
    handle = 0
    
    PersistirGruposParticulas = True
'Exit Function
'ErrHandler:
'If handle Then Close handle
'MsgBox "Ocurrio un error al guardar los grupos de particulas globales."

End Function

Public Function CargarGruposParticulas(ByVal Path As String) As Boolean
'On Error GoTo ErrHandler
    If Not FileExist(Path) Then
        LogError "No se pudo cargar el archivo de particulas porque no existe. ->" & Path
        Exit Function
    End If

    Dim handle As Integer
    
    Dim i As Integer
    
    handle = FreeFile
    
    Open Path For Binary Access Read As handle
        Get handle, , GlobalParticleGroupCount
        ReDim GlobalParticleGroup(1 To GlobalParticleGroupCount)
        
        If GlobalParticleGroupCount > 0 Then
            For i = 1 To GlobalParticleGroupCount
                Set GlobalParticleGroup(i) = New Engine_Particle_Group
                
                If Not GlobalParticleGroup(i).Cargar(handle) Then
                    Set GlobalParticleGroup(i) = Nothing
                End If
                
                
            Next i
        End If
    Close handle
    
    handle = 0
    
    CargarGruposParticulas = True
'Exit Function
'ErrHandler:
'If handle Then Close handle
'MsgBox "Ocurrio un error al cargar los grupos de particulas globales."

End Function

Public Sub CargarParticle_Streams()
    CargarGruposParticulas RecursosPath & "Particles.bin"
End Sub

Public Sub PersistirParticleStreams()
    PersistirGruposParticulas RecursosPath & "\Particles.bin"
End Sub

Public Sub DibujarParticulas(ByRef puntero() As PARTVERTEX, ByVal cantidad As Long, ByVal textura As Long, ByVal blend_mode As Long)
    Dim estadoAnterior As Long
    
    'estadoAnterior = Engine.GetVertexShader
    
    If cantidad <= 0 Then Exit Sub
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(textura)
    Set_Blend_Mode blend_mode
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Particulas
    
    
    If cfgSoportaPointSprites Then
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, cantidad, puntero(0), Part_size
    Else
        Dim dest() As Box_Vertex
        ReDim dest(cantidad)
        
        IniciarParticulaBox dest
        particulaABox puntero, dest, cantidad

        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, cantidad * 2, dest(0), TL_size
    End If
    
    'Engine.SetVertexShader estadoAnterior
End Sub


Public Sub IniciarParticulaBox(ByRef destino() As Box_Vertex)
    Dim i As Integer
    Dim tmax As Integer
    
    tmax = UBound(destino)
    
    For i = 0 To tmax
        With destino(i)
            .tu0 = 0
            .tv0 = 1
            .tv1 = 0
            .tu2 = 1
            
            .tu1 = 0
            .tv2 = 1
            .tu3 = 1
            .tv3 = 0
            
            .rhw0 = 1
            .rhw1 = 1
            .rhw2 = 1
            .rhw3 = 1
        End With
    Next i
End Sub


Public Sub particulaABox(ByRef origen() As PARTVERTEX, ByRef destino() As Box_Vertex, ByVal cantidad As Long)
    'TODO PASAR ESTO A Cpp
    Dim i As Integer
    Dim tmax As Integer
    
    tmax = minl(minl(cantidad - 1, UBound(origen)), UBound(destino))
    
    Dim half_size As Single, x As Single, y As Single, Color As Long, y2 As Single, x2 As Single

    For i = 0 To tmax
        With origen(i)
            half_size = .Tamanio / 2
            x = .v.x - half_size
            y = .v.y - half_size
            x2 = .v.x + half_size
            y2 = .v.y + half_size
            Color = .Color
        End With
        
        With destino(i)
            .color0 = Color
            .Color1 = Color
            .Color2 = Color
            .color3 = Color
            
            .x0 = x
            .x1 = x
            .x2 = x2
            .x3 = x2
            
            .y0 = y2
            .y1 = y
            .y2 = y2
            .y3 = y
        End With
    Next i
End Sub


