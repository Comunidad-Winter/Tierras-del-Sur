VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vw_Part_Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Implements vWindow






Private WithEvents ListaEmisores        As ControlList
Attribute ListaEmisores.VB_VarHelpID = -1

Private WithEvents BotonSubirEmisor     As ControlBoton
Attribute BotonSubirEmisor.VB_VarHelpID = -1
Private WithEvents BotonBajarEmisor     As ControlBoton
Attribute BotonBajarEmisor.VB_VarHelpID = -1

Private WithEvents BotonQuitarEmisor    As ControlBoton
Attribute BotonQuitarEmisor.VB_VarHelpID = -1
Private WithEvents BotonAgregarEmisor   As ControlBoton
Attribute BotonAgregarEmisor.VB_VarHelpID = -1
Private WithEvents BotonEditarEmisor    As ControlBoton
Attribute BotonEditarEmisor.VB_VarHelpID = -1

Private WithEvents BotonAplicar         As ControlBoton
Attribute BotonAplicar.VB_VarHelpID = -1
Private WithEvents BotonGuardar         As ControlBoton
Attribute BotonGuardar.VB_VarHelpID = -1
Private WithEvents BotonCancelar        As ControlBoton
Attribute BotonCancelar.VB_VarHelpID = -1

Private WithEvents TextTituloGroup      As ControlInput
Attribute TextTituloGroup.VB_VarHelpID = -1

Private WithEvents Ventanita            As ControlVentana
Attribute Ventanita.VB_VarHelpID = -1


Public group As Engine_Particle_Group

Public GroupOut As Engine_Particle_Group

Private Controles As vControles

Private GrupoBackUp As Engine_Particle_Group

Private vX As Integer, vY As Integer

Public id As Integer


'Implementacion de ventanas. No modificar
' LA PUTA QUE LO PARIOOOOOOO PORQUE NO LO PUEDO EXTENDER Y YA!?!?. LA VIDA SERIA MEJOR SI SE PUDIERA EXTENDER UNA CLASE EN ESTE VB DE MIERDA
    Private pNext As vWindow
    Private pPrev As vWindow

Private Sub BotonEditarEmisor_Click()
    group.EditarEmisorSeleccionado
End Sub

Private Sub BotonGuardar_Click()
    group.SetNombre TextTituloGroup

    group.ClonarEn GroupOut
    
    If id = -1 Then
        VentanaSelectorParticulas.Add GroupOut
    Else
        VentanaSelectorParticulas.SaveIntoGlobals GroupOut, id
    End If
    
    GUI_Quitar Me
End Sub

Private Sub Ventanita_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ContinuarPropagacion As Boolean)
    ContinuarPropagacion = False
End Sub

Private Sub Ventanita_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer, ContinuarPropagacion As Boolean)
    ContinuarPropagacion = False
End Sub

Private Sub Ventanita_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
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

Private Sub Class_Terminate()
    Set Controles = Nothing
End Sub

Public Sub BindGroup(e As Engine_Particle_Group)
    Set GroupOut = e
    
    Set group = e.Clonar
    Titulo = group.ObtenerLabel
    TextTituloGroup.text = Trim$(group.ObtenerLabel)
    
    Set GrupoBackUp = e.Clonar()
    
    group.BindList ListaEmisores
End Sub

Private Sub Class_Initialize()
    Set Controles = New vControles
    
    Set ListaEmisores = Controles.Add(New ControlList)
    
    Set BotonSubirEmisor = Controles.Add(New ControlBoton)
    Set BotonBajarEmisor = Controles.Add(New ControlBoton)
    
    Set BotonAgregarEmisor = Controles.Add(New ControlBoton)
    Set BotonEditarEmisor = Controles.Add(New ControlBoton)
    Set BotonQuitarEmisor = Controles.Add(New ControlBoton)
    
    Set BotonAplicar = Controles.Add(New ControlBoton)
    Set BotonCancelar = Controles.Add(New ControlBoton)
    Set BotonGuardar = Controles.Add(New ControlBoton)
    
    Set TextTituloGroup = Controles.Add(New ControlInput)
    
    Set Ventanita = Controles.Add(New ControlVentana)
    
    Dim CancelarEventosMouseVentanaVieja As ControlMouseReader
    Set CancelarEventosMouseVentanaVieja = Controles.Add(New ControlMouseReader)
    CancelarEventosMouseVentanaVieja.SetSize 300, 400
    CancelarEventosMouseVentanaVieja.SetPos 10, 32
    
    BotonSubirEmisor.text = "Subir"
    BotonBajarEmisor.text = "Bajar"
    
    BotonAgregarEmisor.text = "Agregar"
    BotonEditarEmisor.text = "Editar"
    BotonQuitarEmisor.text = "Quitar"
    
    BotonCancelar.text = "Cancelar"
    BotonGuardar.text = "Guardar"
    
    TextTituloGroup.Label = "Nombre"
    
    BotonSubirEmisor.SetSize 64, 16
    BotonBajarEmisor.SetSize 64, 16
    BotonAgregarEmisor.SetSize 64, 16
    BotonEditarEmisor.SetSize 64, 16
    BotonQuitarEmisor.SetSize 64, 16

    
    ListaEmisores.SetSize 128, 128

    
    BotonCancelar.SetSize 64, 16
    BotonGuardar.SetSize 64, 16
    
    vWindow_SetPos 48, 48
    
    Ventanita.SetSize 250, 224
    Ventanita.Titulo = "Editar grupo de partículas"
    Ventanita.ventanaHost = Me
    
    
    Call ListaEmisores_Change(-1)
    
End Sub


Private Sub vWindow_Hide()
    GUI_Quitar Me
End Sub

Private Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
    Ventanita.SetPos vX, vY
    TextTituloGroup.SetPos vX + 16, vY + 29
    ListaEmisores.SetPos vX + 16, vY + 48
    BotonSubirEmisor.SetPos vX + 128 + 32, vY + 48
    BotonBajarEmisor.SetPos vX + 128 + 32, vY + 48 + 19
    BotonAgregarEmisor.SetPos vX + 128 + 32, vY + 90 + 32
    BotonEditarEmisor.SetPos vX + 128 + 32, vY + 90 + 51
    BotonQuitarEmisor.SetPos vX + 128 + 32, vY + 90 + 70
    BotonCancelar.SetPos vX + 16, vY + 48 + 16 + 128
    BotonGuardar.SetPos vX + 37 + 48, vY + 48 + 16 + 128
End Sub

Private Sub vWindow_Show()
    GUI_SetFocus Me
End Sub

Private Function vWindow_Render() As Boolean
    vWindow_Render = True
    vWindow_SetPos 20, 100
    Controles.Render
    
    Call group.SetPosAbsolute(384 + (D3DWindow.BackBufferWidth - 384) / 2, D3DWindow.BackBufferHeight / 2)
    
    If group.Render = False Then
        
    End If
End Function

Private Sub BotonAgregarEmisor_Click()
    group.AgregarNuevoEmisor
End Sub

Private Sub BotonBajarEmisor_Click()
    group.BajarEmisor
End Sub

Private Sub BotonCancelar_Click()
    GrupoBackUp.ClonarEn GroupOut
    GUI_Quitar Me
End Sub

Private Sub BotonQuitarEmisor_Click()
    group.QuitarEmisor
End Sub

Private Sub BotonSubirEmisor_Click()
    group.SubirEmisor
End Sub

Private Sub ListaEmisores_Change(SelectedIndex As Integer)
    If SelectedIndex = -1 Then
        BotonQuitarEmisor.Enabled = False
        BotonEditarEmisor.Enabled = False
    Else
        BotonQuitarEmisor.Enabled = True
        BotonEditarEmisor.Enabled = True
    End If
    
    If ListaEmisores.Count > 1 Then
        BotonSubirEmisor.Enabled = True
        BotonBajarEmisor.Enabled = True
    Else
        BotonSubirEmisor.Enabled = False
        BotonBajarEmisor.Enabled = False
    End If
End Sub
