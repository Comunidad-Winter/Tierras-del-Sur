VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine_Particle_Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PixelPos            As mzVECTOR2

Private VelocidadEnPixeles  As Long

Private ListaEmisores()     As Engine_Particle_Stream
Private NuevoOrden()        As Integer
Private CantidadEmisores    As Long

Private Nombre              As String * 16

Public PGID                 As Integer

Public MapX As Integer
Public MapY As Integer

Private AbsolutePosition    As Boolean

Private UltimoRender        As Long

#If esMe = 1 Then
    Private BindedList As ControlList
    Private VentanaEdicion As vw_Part_Group
#End If

Public Property Let CargarPGID(v As Integer)
Attribute CargarPGID.VB_UserMemId = 0
Attribute CargarPGID.VB_MemberFlags = "200"
    PGID = v
    If PartGroupValid(v) Then
        GlobalParticleGroup(v).ClonarEn Me
    End If
End Property

Public Function Clonar() As Engine_Particle_Group
    Dim Tmp As New Engine_Particle_Group
    Call Me.ClonarEn(Tmp)
    Set Clonar = Tmp
End Function

Public Sub LimpiarEmisores()
    CantidadEmisores = 0
    Dim i As Integer
    
    For i = 0 To CantidadEmisores - 1
        Set ListaEmisores(i) = Nothing
    Next i
    
    ReDim ListaEmisores(0)
End Sub

Public Function ClonarEn(T As Engine_Particle_Group)
    If Not T Is Nothing Then
    
        T.LimpiarEmisores
        #If esMe = 1 Then
            T.SetNombre Trim$(Nombre)
        #End If
        
        Dim i As Integer
        For i = 0 To CantidadEmisores - 1
            T.AgregarEmisor ListaEmisores(i).Clone
            ClonarEn = True
        Next i
    End If
End Function

Private Sub ReconstruirLista()
    ReDim Preserve ListaEmisores(maxl(CantidadEmisores - 1, 0))
    ReDim Preserve NuevoOrden(maxl(CantidadEmisores - 1, 0))
End Sub

Private Sub LimpiarLista()
    Dim i As Integer
    For i = 0 To CantidadEmisores - 1
        Set ListaEmisores(i) = Nothing
        Set ListaEmisores(i) = New Engine_Particle_Stream
    Next i
End Sub

Public Function Render() As Boolean
    If CantidadEmisores > 0 Then
        'Engine.SetVertexShader particleFVF
    
        If MapX > 0 And MapY > 0 Then
            offset_map_part.y = offset_map.y - AlturaPie(MapX, MapY) + minYOffset * 32
        Else
            offset_map_part.y = offset_map.y + minYOffset * 32
        End If
        
        Dim i As Integer
        Dim contador As Integer
        
        For i = 0 To CantidadEmisores - 1
            ListaEmisores(i).SetPixelPos PixelPos.x, PixelPos.y
            If ListaEmisores(i).Render Then
                NuevoOrden(contador) = i
                contador = contador + 1
            End If
        Next i
        
        Set_Blend_Mode 0
        
        If contador <> CantidadEmisores Then
            For i = 0 To CantidadEmisores - 1
                If i <= contador Then
                    Set ListaEmisores(i) = ListaEmisores(NuevoOrden(i))
                Else
                    Set ListaEmisores(i) = Nothing
                End If
            Next i
            
            CantidadEmisores = contador
            ReDim Preserve ListaEmisores(contador)
            ReDim NuevoOrden(contador)
        End If
        

        'Engine.SetVertexShader FVF
        
        Render = (contador > 0)
    End If
End Function

Public Sub SetPosAbsolute(ByVal x!, ByVal y!)
    PixelPos.x = x
    PixelPos.y = y
    
    Dim i As Integer
    If CantidadEmisores > 0 Then
        For i = 0 To CantidadEmisores - 1
            ListaEmisores(i).SetPixelPos x, y
            ListaEmisores(i).AbsolutePosition = True
        Next i
    End If
    
    MapX = 0
    MapY = 0
    
    AbsolutePosition = True
End Sub

Public Sub SetPixelPos(ByVal x!, ByVal y!)
    PixelPos.x = x
    PixelPos.y = y
    
    Dim i As Integer
    If CantidadEmisores > 0 Then
        For i = 0 To CantidadEmisores - 1
            ListaEmisores(i).SetPixelPos x, y
            ListaEmisores(i).AbsolutePosition = False
        Next i
    End If
    
    MapX = x \ 32
    MapY = y \ 32
    
    AbsolutePosition = False
End Sub

Public Sub SetPos(ByVal map_x As Integer, ByVal map_y As Integer)
    Call SetPixelPos(map_x * 32 + 16, map_y * 32 + 16)
End Sub

Public Sub Matar(Optional ByVal Milisegundos As Long = 0)
    Dim i As Integer
    If CantidadEmisores > 0 Then
        For i = 0 To CantidadEmisores - 1
            ListaEmisores(i).Matar Milisegundos
        Next i
    End If
End Sub

Public Sub persistir(handle As Integer, Optional ByVal Offset As Long = 0)
    If Offset > 0 Then
        Seek handle, Offset
    End If
    
    Put handle, , PGID
    
    'Si PGID es cero significa que el emisor esta hecho "al vuelo" y esta es la UNICA instancia en todo el juego de esta particula.
    'si PGID es cero, se guarda todo el grupo en el mapa
    'si PGID NO ES CERO se guarda todo el grupo en el archivo global de particulas
    
    If PGID = 0 Then
        Put handle, , Nombre
        Put handle, , CantidadEmisores
    
        Dim i As Integer
        
        If CantidadEmisores > 0 Then
            For i = 0 To CantidadEmisores - 1
                ListaEmisores(i).EscribirEmisor handle
            Next
        End If
    End If
End Sub

Public Function Cargar(handle As Integer, Optional ByVal Offset As Long = 0) As Boolean
    If Offset > 0 Then
        Seek handle, Offset
    End If
    
    Get handle, , PGID
    
    If PGID = 0 Then
        Get handle, , Nombre
        Get handle, , CantidadEmisores
    
        ReconstruirLista
    
        If CantidadEmisores > 0 Then
            Dim i As Integer
        
            For i = 0 To CantidadEmisores - 1
                Set ListaEmisores(i) = New Engine_Particle_Stream
                ListaEmisores(i).LeerEmisor handle
            Next
            Cargar = True
        End If
    Else
        If PartGroupValid(PGID) Then
            Cargar = GlobalParticleGroup(PGID).ClonarEn(Me)
        End If
    End If
End Function

Public Function AgregarEmisor(T As Engine_Particle_Stream) As Integer ' Return index
    CantidadEmisores = CantidadEmisores + 1
    
    ReconstruirLista
    
    Set ListaEmisores(CantidadEmisores - 1) = T
    
    If AbsolutePosition Then
       T.AbsolutePosition = True
    End If
    
    T.SetPixelPos PixelPos.x, PixelPos.y
     
    AgregarEmisor = CantidadEmisores - 1
End Function

#If esMe = 1 Then

Public Function IniciarEdicion() As vw_Part_Group
    If VentanaEdicion Is Nothing Then
        Set VentanaEdicion = New vw_Part_Group
    End If
    
    VentanaEdicion.BindGroup Me
    
    Set IniciarEdicion = VentanaEdicion
    
    GUI_SetFocus VentanaEdicion
End Function

Public Sub RefreshList()
    If Not BindedList Is Nothing Then
        BindedList.Clear
        
        Dim i As Integer
        
        For i = 0 To CantidadEmisores - 1
            BindedList.AddItem ListaEmisores(i).ObtenerLabel
        Next
    End If
End Sub

Public Sub BindList(list As ControlList)
    Set BindedList = list
    RefreshList
End Sub

Public Sub SubirEmisor()
    If Not BindedList Is Nothing Then
        Dim cual As Integer
        cual = BindedList.SelectedIndex
        If cual < CantidadEmisores And cual > 0 Then
            Dim T As Engine_Particle_Stream
            Set T = ListaEmisores(cual - 1)
            Set ListaEmisores(cual - 1) = ListaEmisores(cual)
            Set ListaEmisores(cual) = T
            Set T = Nothing
        End If
        RefreshList
    End If
End Sub

Public Sub BajarEmisor()
    If Not BindedList Is Nothing Then
        Dim cual As Integer
        cual = BindedList.SelectedIndex
        If cual < CantidadEmisores - 1 And cual <> -1 Then
            Dim T As Engine_Particle_Stream
            Set T = ListaEmisores(cual + 1)
            Set ListaEmisores(cual + 1) = ListaEmisores(cual)
            Set ListaEmisores(cual) = T
            Set T = Nothing
        End If
        RefreshList
    End If
End Sub

Public Sub QuitarEmisor()
    If Not BindedList Is Nothing Then
        Dim cual As Integer
        cual = BindedList.SelectedIndex
        If cual < CantidadEmisores And cual <> -1 Then
            Dim i As Integer
            Dim contador As Integer
            
            For i = 0 To CantidadEmisores - 1
                If i <> cual Then
                    Set ListaEmisores(contador) = ListaEmisores(i)
                    contador = contador + 1
                End If
            Next i
            
            CantidadEmisores = contador
            ReDim Preserve ListaEmisores(contador)
            ReDim NuevoOrden(contador)
        End If
        RefreshList
    End If
End Sub

Public Sub EditarEmisorSeleccionado()
    If Not BindedList Is Nothing Then
        Dim Index As Integer
        Index = BindedList.SelectedIndex
        
        If Index > -1 And Index < CantidadEmisores Then
            ListaEmisores(Index).EmpezarEdicion
        End If
    End If
End Sub

Public Function AgregarNuevoEmisor() As Integer ' Return index
     Dim T As Engine_Particle_Stream
     
     Set T = New Engine_Particle_Stream
     
     Dim EmisorBase As ParticleStream
     
     With EmisorBase
        .life = 2000
        .rnd_x = 1
        .rnd_y = 1
        .mod_timer = 1
        .rnd_alpha_factor = 1
        .alpha_coseno = 1
        .progress_factor = 1
        .NumOfParticles = 20
     End With

    T.SetEmisor EmisorBase

     AgregarNuevoEmisor = AgregarEmisor(T)
     
     RefreshList
     
     If Not BindedList Is Nothing Then
        BindedList.SelectedIndex = AgregarNuevoEmisor
     End If
     
     EditarEmisorSeleccionado
End Function

Public Function ObtenerLabel() As String
    ObtenerLabel = Trim$(Nombre)
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

Private Sub Class_Initialize()
    Nombre = Space$(16)
End Sub
