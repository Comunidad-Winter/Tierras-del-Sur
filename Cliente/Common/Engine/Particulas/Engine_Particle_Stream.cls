VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine_Particle_Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta clase guarda una instancia NUEVA de un emisor de partículas. Los datos de este emisor son cargados desde un archivo
'o desde un paquete de datos o desde cualquier lado.

'Las particulas pertenecientes a los personajes se dibujan en 2 etapas; la primera son las particulas que estan atrás del personaje. que es la primera mitad de la cantidad total.
'La segunda es la otra mitad, se dibujan encima del personaje. Con esto podemos dibujar particulas al redcedor del pj y da efecto de profundidad.


Option Explicit

Private Particulas()            As Particle         'Lista de particulas
Private Particulas_Vertices()   As PARTVERTEX       'Lista de vertices que se van a dibujar.
Private Particulas_Cajas()      As Box_Vertex       'Lista de vertices que se van a dibujar como caja.

Private PixelPos                As mzVECTOR2        'Posicion en pixeles del mapa

Private TickMuerte              As Long

Private CantidadAEmitir         As Long             'Cantidad de particulas que se tienen que emitir durante toda la vida del emisor.

Private DatosEmisor             As ParticleStream

Private Nombre                  As String * 16

Public PEID                     As Long             ' Particle Emissor ID

Private Const EmitirInfinitasParticulas As Long = -1

Private CantidadADibujar As Integer

Private UltimoRender            As Long
Private UltimoRenderX           As Long
Private UltimoRenderY           As Long

Public AbsolutePosition         As Boolean

#If esMe = 1 Then
    Private VentanaEdicion      As vw_Part_Stream
    Private Ventana             As vWindow
    Private DatosEmisorOriginal As ParticleStream
    Private EstoyEditando       As Boolean
#End If

Private AcumuladorTimer         As Single

Friend Sub SetEmisor(grupo As ParticleStream)
    DatosEmisor = grupo
    
    DatosEmisor.trig_factor = (pi * 2) / maxl(DatosEmisor.NumOfParticles, 1)
    
    ReDim Preserve Particulas(DatosEmisor.NumOfParticles)
    ReDim Preserve Particulas_Vertices(DatosEmisor.NumOfParticles)
    
    'If Not cfgSoportaPointSprites Then
        ReDim Preserve Particulas_Cajas(DatosEmisor.NumOfParticles)
    'End If
    
    If DatosEmisor.CantidadAEmitir <> 0 Then
        CantidadAEmitir = DatosEmisor.CantidadAEmitir
    Else
        CantidadAEmitir = EmitirInfinitasParticulas
    End If
    
    IniciarListaParticulas
End Sub

Public Sub reset()
    ReDim Particulas(DatosEmisor.NumOfParticles)
    ReDim Particulas_Vertices(DatosEmisor.NumOfParticles)
    
    If Not cfgSoportaPointSprites Then
        ReDim Particulas_Cajas(DatosEmisor.NumOfParticles)
    End If
    
    If DatosEmisor.CantidadAEmitir <> 0 Then
        CantidadAEmitir = DatosEmisor.CantidadAEmitir
    Else
        CantidadAEmitir = EmitirInfinitasParticulas
    End If
    
    IniciarListaParticulas
End Sub

Private Sub IniciarListaParticulas()
    Dim i As Integer
    For i = 0 To DatosEmisor.NumOfParticles
        Particulas_Vertices(i).rhw = 1
    Next i
    
    If Not cfgSoportaPointSprites Then
        IniciarParticulaBox Particulas_Cajas
    End If
End Sub


Friend Function Clone() As Engine_Particle_Stream
    Dim T As Engine_Particle_Stream
    Set T = New Engine_Particle_Stream
    #If esMe = 1 Then
    T.SetNombre GetNombre
    #End If
    Call T.SetEmisor(DatosEmisor)
    Set Clone = T
End Function

Friend Function GetEmisor() As ParticleStream
    GetEmisor = DatosEmisor
End Function

Public Sub SetPixelPos(ByVal x As Integer, ByVal y As Integer)
    PixelPos.x = x
    PixelPos.y = y
End Sub

Private Sub Class_Initialize()
    CantidadAEmitir = EmitirInfinitasParticulas
    Nombre = Space$(16)
End Sub

Public Sub Matar(Optional ByVal Milisegundos As Long = 0)
    TickMuerte = Milisegundos + GetTimer
End Sub

Public Function CantidadDibujada() As Long
    CantidadDibujada = CantidadADibujar
End Function

Public Function Render() As Boolean
    
    Dim TickActual As Long
    
    TickActual = GetTimer

    If TickMuerte > 0 Then
        If TickMuerte < TickActual Then
            CantidadAEmitir = 0
        End If
    End If
        
    Render = True
    'cfgSoportaPointSprites = False
    
    If DatosEmisor.NumOfParticles > 0 Then
        AcumuladorTimer = AcumuladorTimer + timerTicksPerFrame

        If bCameraCanged Or (TickActual - UltimoRender) > 24 Then
            Dim vec_offset As mzVECTOR2
            
            If AbsolutePosition = False Then
                vec_offset = offset_map_part
            End If
        
            CantidadADibujar = UpdateParticles(Particulas(0), Particulas_Vertices(0), AcumuladorTimer, PixelPos, DatosEmisor, vec_offset, CantidadAEmitir, TickActual)
            
            If Not cfgSoportaPointSprites Then
                'Call partToBox(Particulas_Vertices(0), CantidadADibujar, Particulas_Cajas(0))
                particulaABox Particulas_Vertices, Particulas_Cajas, CantidadADibujar
            End If
            'cfgSoportaPointSprites = True
            UltimoRender = TickActual
            AcumuladorTimer = 0
        End If
        
        
        If CantidadADibujar = -1 Then
            Render = False
        ElseIf CantidadADibujar > 0 Then
            If DatosEmisor.Line = 1 Then
                Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
                Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Particulas
                D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, CantidadADibujar - 1, Particulas_Vertices(0), Part_size
            Else
                'DibujarParticulas Particulas_Vertices(0), CantidadADibujar, DatosEmisor.texture, DatosEmisor.blend_mode
                Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Particulas
                Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(DatosEmisor.texture)
                Set_Blend_Mode DatosEmisor.blend_mode
                
                
                If cfgSoportaPointSprites Then
                    'Engine.SetVertexShader particleFVF
                    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, CantidadADibujar, Particulas_Vertices(0), Part_size
                Else
                    'Engine.SetVertexShader FVF
                    If IndexBufferEnabled Then
                        D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLELIST, 0, CantidadADibujar * 4, CantidadADibujar * 2, StaticIndexBuffer(0), D3DFMT_INDEX16, Particulas_Cajas(0), TL_size
                    Else
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, CantidadADibujar * 2, Particulas_Cajas(0), TL_size
                    End If
                End If
                
            End If
        End If
    End If
    
End Function

Public Function LeerEmisor(handle As Integer, Optional ByVal Offset As Long = 0)
    If Offset > 0 Then Seek handle, Offset
    Get handle, , PEID
    Get handle, , Nombre
    Get handle, , DatosEmisor
    
    SetEmisor DatosEmisor
End Function

Public Sub EscribirEmisor(handle As Integer, Optional ByVal Offset As Long = 0)
    If Offset > 0 Then Seek handle, Offset
    Put handle, , PEID
    Put handle, , Nombre
    Put handle, , DatosEmisor
End Sub


#If esMe = 1 Then

Public Sub EmpezarEdicion()
    If EstoyEditando = False Then
        DatosEmisorOriginal = DatosEmisor
        EstoyEditando = True
    End If
    
    If VentanaEdicion Is Nothing Then
        Set VentanaEdicion = New vw_Part_Stream
        Set Ventana = VentanaEdicion
        VentanaEdicion.BindEmisor Me
    End If
    
    Ventana.Show
End Sub

Public Sub CancelarEdicion()
    If EstoyEditando = True Then
        DatosEmisor = DatosEmisorOriginal
        EstoyEditando = False
    End If
    
    Ventana.Hide
    
    Set Ventana = Nothing
    Set VentanaEdicion = Nothing
End Sub

Public Sub GuardarEdicion()
    If EstoyEditando = True Then
    
        
        Call SetEmisor(VentanaEdicion.GetDatos)
        
        EstoyEditando = False
        Ventana.Hide
        
        Set Ventana = Nothing
        Set VentanaEdicion = Nothing
        
        If PEID <> 0 Then
            'COMO ES UN EMISOR GLOBAL LO TENGO QUE ACTUALIZAR EN TODOS LOS GRUPOS QUE HAYA CREADOS
        End If
    End If
End Sub

Public Function ObtenerLabel() As String
    ObtenerLabel = Trim$(PEID & " - " & Nombre)
End Function

Public Function GetNombre() As String
    GetNombre = Trim$(Nombre)
End Function

Public Sub SetNombre(val As String)
    val = Trim$(val)
    If Len(val) > 16 Then
        Nombre = left$(val, 16)
    Else
        Nombre = val
    End If
End Sub

#End If
