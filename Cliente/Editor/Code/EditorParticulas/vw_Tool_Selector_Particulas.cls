VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vw_Part_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements vWindow

Private WithEvents BotonNuevoGrupo          As ControlBoton
Attribute BotonNuevoGrupo.VB_VarHelpID = -1
Private WithEvents BotonEditarGrupo         As ControlBoton
Attribute BotonEditarGrupo.VB_VarHelpID = -1
Private WithEvents BotonBorrarGrupo         As ControlBoton
Attribute BotonBorrarGrupo.VB_VarHelpID = -1

'Private WithEvents BotonInsertarGrupo       As ControlBoton

Private WithEvents BotonGuardar             As ControlBoton
Attribute BotonGuardar.VB_VarHelpID = -1
Private WithEvents BotonCargar              As ControlBoton
Attribute BotonCargar.VB_VarHelpID = -1

Private WithEvents ListaGrupos              As ControlList
Attribute ListaGrupos.VB_VarHelpID = -1

Private WithEvents Ventana                  As ControlVentana
Attribute Ventana.VB_VarHelpID = -1

Private WithEvents Controles                As vControles
Attribute Controles.VB_VarHelpID = -1

Private WithEvents BotonCapa1               As ControlCheckBox
Attribute BotonCapa1.VB_VarHelpID = -1
Private WithEvents BotonCapa2               As ControlCheckBox
Attribute BotonCapa2.VB_VarHelpID = -1
Private WithEvents BotonCapa3               As ControlCheckBox
Attribute BotonCapa3.VB_VarHelpID = -1

Private WithEvents BotonAccionInsertar      As ControlCheckBox
Attribute BotonAccionInsertar.VB_VarHelpID = -1
Private WithEvents BotonAccionBorrar        As ControlCheckBox
Attribute BotonAccionBorrar.VB_VarHelpID = -1

Private GrupoSeleccionadoIndex              As Integer

Private vX As Integer, vY As Integer


'Esta es la particula (con sus 3 capas) que esta seleccionando el usuario
Private ParticulasSeleccionadas(2) As Engine_Particle_Group

'Implementacion de ventanas. No modificar
' LA PUTA QUE LO PARIOOOOOOO PORQUE NO LO PUEDO EXTENDER Y YA!?!?. LA VIDA SERIA MEJOR SI SE PUDIERA EXTENDER UNA CLASE EN ESTE VB DE MIERDA
    Private pNext As vWindow
    Private pPrev As vWindow
    
    
Private Sub BotonAccionBorrar_Click()
    BotonAccionBorrar = True
    BotonAccionInsertar = False
    SetElementosSeleccion
End Sub

Private Sub BotonAccionInsertar_Click()
    BotonAccionBorrar = False
    BotonAccionInsertar = True
    SetElementosSeleccion
End Sub

Private Sub BotonCapa1_Click()
    BotonCapa1 = True
    BotonCapa2 = False
    BotonCapa3 = False
    SetElementosSeleccion
End Sub

Private Sub BotonCapa2_Click()
    BotonCapa2 = True
    BotonCapa1 = False
    BotonCapa3 = False
    SetElementosSeleccion
End Sub

Private Sub BotonCapa3_Click()
    BotonCapa3 = True
    BotonCapa2 = False
    BotonCapa1 = False
    SetElementosSeleccion
End Sub

Private Sub BotonCargar_Click()
    'Engine_Particles_Storage.CargarGruposParticulas "c:\particulas.bin"
    CargarParticle_Streams
    BotonParticulasGlobales_Click
End Sub

Private Sub BotonGuardar_Click()
    'Engine_Particles_Storage.PersistirGruposParticulas "c:\particulas.bin"
    PersistirParticleStreams
End Sub

Private Sub Ventana_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ContinuarPropagacion As Boolean)
    ContinuarPropagacion = False
End Sub

Private Sub Ventana_KeyPress(ByVal KeyAscii As Integer, ContinuarPropagacion As Boolean)
    ContinuarPropagacion = False
End Sub

Private Sub Ventana_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    GUI_SetFocus Me
End Sub

    Private Property Get vWindow_GetControl() As vControles
        Set vWindow_GetControl = Controles
    End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function
    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
'/Implementacion

Private Sub BotonEditarGrupo_Click()
    If PartGroupValid(GrupoSeleccionadoIndex) Then
        If Not GlobalParticleGroup(GrupoSeleccionadoIndex) Is Nothing Then
            GlobalParticleGroup(GrupoSeleccionadoIndex).IniciarEdicion().id = GrupoSeleccionadoIndex
        End If
    End If
End Sub

Private Sub BotonNuevoGrupo_Click()
    EdPar_Grupo_Nuevo
End Sub

Private Sub BotonParticulasGlobales_Click()
    EsListaGlobal = True

    ListaGrupos.Clear
    'Agregar a listagrupos todos los grupos globales guardados en la memoria
    If GlobalParticleGroupCount > 0 Then
        Dim i As Integer
        For i = 1 To GlobalParticleGroupCount
            ListaGrupos.AddItem (i & " - " & GlobalParticleGroup(i).ObtenerLabel)
        Next i
    End If
    
    If GrupoSeleccionadoIndex < ListaGrupos.count And GrupoSeleccionadoIndex <> -1 Then
        ListaGrupos.SelectedIndex = GrupoSeleccionadoIndex
    End If
    
    SetElementosSeleccion
End Sub



Private Sub ListaGrupos_Change(SelectedIndex As Integer)
    If SelectedIndex = -1 Then
        BotonEditarGrupo.Enabled = False
        BotonBorrarGrupo.Enabled = False
        
        GrupoSeleccionadoIndex = PartGroupNONE
        
        BotonAccionBorrar = True
        BotonAccionInsertar = False
    Else
        BotonEditarGrupo.Enabled = True
        BotonBorrarGrupo.Enabled = True And False
        
        GrupoSeleccionadoIndex = SelectedIndex + 1
        
        BotonAccionInsertar = True
        BotonAccionBorrar = False
    End If
    
    SetElementosSeleccion
End Sub

Private Sub Class_Initialize()

'Inicio los controles
    Set Controles = New vControles

    Set BotonNuevoGrupo = Controles.Add(New ControlBoton)
    Set BotonEditarGrupo = Controles.Add(New ControlBoton)
    Set BotonBorrarGrupo = Controles.Add(New ControlBoton)
    
    Set BotonGuardar = Controles.Add(New ControlBoton)
    Set BotonCargar = Controles.Add(New ControlBoton)
    
    Set BotonCapa1 = Controles.Add(New ControlCheckBox)
    Set BotonCapa2 = Controles.Add(New ControlCheckBox)
    Set BotonCapa3 = Controles.Add(New ControlCheckBox)
    
    Set BotonAccionInsertar = Controles.Add(New ControlCheckBox)
    Set BotonAccionBorrar = Controles.Add(New ControlCheckBox)
    
    Set ListaGrupos = Controles.Add(New ControlList)
    
    Set Ventana = Controles.Add(New ControlVentana)
    
' Seteo los estados de los campos
    
    Call BotonParticulasGlobales_Click
    Call ListaGrupos_Change(-1)
    
    Ventana.SetSize 300, 400
    Ventana.Titulo = "Partículas"
    Ventana.ventanaHost = Me
    
    BotonCapa1.text = "Capa 1"
    BotonCapa2.text = "Capa 2"
    BotonCapa3.text = "Capa 3"
    
    BotonAccionInsertar.text = "Insertar partícula"
    BotonAccionBorrar.text = "Borrar partícula"
    

' Posiciono los controles


    ListaGrupos.SetSize 280, 230
    BotonGuardar.SetSize 64, 16
    BotonCargar.SetSize 64, 16

    BotonGuardar.text = "Guardar"
    BotonCargar.text = "ReCargar"

    BotonNuevoGrupo.SetSize 64, 16
    BotonEditarGrupo.SetSize 64, 16
    BotonBorrarGrupo.SetSize 64, 16


    BotonNuevoGrupo.text = "Nuevo +"
    BotonEditarGrupo.text = "Editar"
    BotonBorrarGrupo.text = "Borrar"
    
    vWindow_SetPos 10, 32
End Sub

Public Sub SetElementosSeleccion()
    If Not BotonCapa1.value Then
        Set ParticulasSeleccionadas(0) = Nothing
    Else
        If BotonAccionInsertar Then
            Set ParticulasSeleccionadas(0) = New Engine_Particle_Group
            ParticulasSeleccionadas(0).PGID = GrupoSeleccionadoIndex
        Else
            Set ParticulasSeleccionadas(0) = New Engine_Particle_Group
            ParticulasSeleccionadas(0).PGID = PartGroupNONE
        End If
    End If
    
    If Not BotonCapa2.value Then
        Set ParticulasSeleccionadas(1) = Nothing
    Else
        If BotonAccionInsertar Then
            Set ParticulasSeleccionadas(1) = New Engine_Particle_Group
            ParticulasSeleccionadas(1).PGID = GrupoSeleccionadoIndex
        Else
            Set ParticulasSeleccionadas(1) = New Engine_Particle_Group
            ParticulasSeleccionadas(1).PGID = PartGroupNONE
        End If
    End If
    
    If Not BotonCapa3.value Then
        Set ParticulasSeleccionadas(2) = Nothing
    Else
        If BotonAccionInsertar Then
            Set ParticulasSeleccionadas(2) = New Engine_Particle_Group
            ParticulasSeleccionadas(2).PGID = GrupoSeleccionadoIndex
        Else
            Set ParticulasSeleccionadas(2) = New Engine_Particle_Group
            ParticulasSeleccionadas(2).PGID = PartGroupNONE
        End If
    End If
    
    Call establecerParticula(ParticulasSeleccionadas)
End Sub

Private Sub vWindow_Hide()

End Sub

Private Sub vWindow_SetPos(ByVal x As Integer, ByVal y As Integer)
    vX = x
    vY = y
    Ventana.SetPos x, y
    
    BotonCargar.SetPos vX + 16, vY + 32
    BotonGuardar.SetPos vX + 32 + 64, vY + 32
    ListaGrupos.SetPos vX + 10, vY + 64
'    ListaSeleccion.SetPos vX + 16, vY + 64 + 128 + 48 + 48
    
    BotonNuevoGrupo.SetPos vX + 10, vY + 64 + 228 + 16
    BotonEditarGrupo.SetPos vX + 32 + 64 - 6, vY + 64 + 228 + 16
    BotonBorrarGrupo.SetPos vX + 32 + 64 + 64 + 10, vY + 64 + 228 + 16
    
    BotonAccionInsertar.SetPos vX + 100, vY + 330
    BotonAccionBorrar.SetPos vX + 100, vY + 350
    
    BotonCapa1.SetPos vX + 10, vY + 330
    BotonCapa2.SetPos vX + 10, vY + 350
    BotonCapa3.SetPos vX + 10, vY + 370
End Sub

Private Sub vWindow_Show()
    GUI_SetFocus Me
End Sub

Private Function vWindow_Render() As Boolean
    vWindow_Render = True
    Controles.Render
    vWindow_SetPos vX, vY
End Function

Public Sub Add(group As Engine_Particle_Group)
    GlobalParticleGroupCount = GlobalParticleGroupCount + 1
    
    ReDim Preserve GlobalParticleGroup(1 To GlobalParticleGroupCount)
    
    Set GlobalParticleGroup(GlobalParticleGroupCount) = group
    
    BotonParticulasGlobales_Click
End Sub

Public Sub SaveIntoGlobals(group As Engine_Particle_Group, id As Integer)
    If PartGroupValid(id) Then
        Set GlobalParticleGroup(id) = group
        BotonParticulasGlobales_Click
    Else
        MsgBox "Id invalido."
    End If
End Sub


