VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlFichaPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControl

Private Nombre As ControlLabel
Attribute Nombre.VB_VarHelpID = -1
Private cabeza As ControlImage
Private cuerpo As ControlImage
Attribute cuerpo.VB_VarHelpID = -1
Private fondo As ControlBoton
Private indicadorModoCandado As ControlImage
Private indicadorOnline As ControlImage

Private Pj As personaje

Private vwControles As vControles

Private CX%, cY%, cH%, cW%

Private ControlRect As RECT

Private callback As vwCuenta

Private ImagenBarba As ControlPicture
Private ImagenPelo As ControlPicture
Private ImagenRopaInterior As ControlPicture
Private ImagenCuerpo As ControlPicture
Private ImagenCara As ControlPicture

Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

Private Sub Class_Initialize()
    Set vwControles = New vControles

    Set Nombre = vwControles.Add(New ControlLabel)
    
    ' El personaje
    Set ImagenBarba = vwControles.Add(New ControlPicture)
    Set ImagenPelo = vwControles.Add(New ControlPicture)
    Set ImagenRopaInterior = vwControles.Add(New ControlPicture)
    Set ImagenCuerpo = vwControles.Add(New ControlPicture)
    Set ImagenCara = vwControles.Add(New ControlPicture)
    
    Set indicadorModoCandado = vwControles.Add(New ControlImage)
    Set indicadorOnline = vwControles.Add(New ControlImage)
    Set fondo = vwControles.Add(New ControlBoton)
    
    fondo.textura = 4602
    fondo.SetSize 128, 128
    
    Nombre.SetCentrado True
End Sub

Public Sub setCallBack(Form As vwCuenta)
    Set callback = Form
End Sub

Public Property Let personaje(personaje As personaje)
    Set Pj = personaje
    Call actualizarPersonaje
End Property


Private Sub actualizarEstetica()
    Dim grhId As Integer

    Dim centro As Integer
    Dim posicionCabeza As Integer
    Dim posicionCuerpoX As Integer
    Dim posicionCuerpoY As Integer
    
    fondo.SetPos CX, cY
       
    posicionCuerpoX = CX + (cW / 2) - 32 / 2
    posicionCuerpoY = cY + cH / 2 - 64 / 2
    
    centro = posicionCuerpoX

    'cabeza.SetPos posicionCuerpoX + BodyData(Pj.cuerpo).HeadOffset.X + 5, posicionCuerpoY + BodyData(Pj.cuerpo).HeadOffset.Y + 30
   ' cuerpo.SetPos posicionCuerpoX, posicionCuerpoY
   
    If Pj.cuerpoDesnudo > 0 Then
        grhId = BodyData(Pj.cuerpoDesnudo).Walk(E_Heading.NORTH).GrhIndex

        ImagenCuerpo.setOrigin 32, 0, 32, 64
        ImagenCuerpo.setTextura (CInt(GrhData(GrhData(grhId).frames(1)).filenum)), 256, 256
        ImagenCuerpo.IControl_SetSize 32, 64
        ImagenCuerpo.SetPos centro, posicionCuerpoY

        If BodyData(Pj.cuerpo).HeadOffset.Y = -28 Then
            posicionCabeza = posicionCuerpoY + BodyData(Pj.cuerpo).HeadOffset.Y + 32
        Else
            posicionCabeza = posicionCuerpoY + BodyData(Pj.cuerpo).HeadOffset.Y + 32
        End If
    End If
    
    If posicionCabeza > 0 Then
     If Pj.cabeza > 0 Then
         ImagenCara.setTextura (GrhData(HeadData(Pj.cabeza).Head(E_Heading.NORTH).GrhIndex).filenum), 128, 32
         ImagenCara.setOrigin 0, 0, 32, 32
         ImagenCara.IControl_SetSize 32, 32
         ImagenCara.SetPos centro, posicionCabeza
     End If
     
     If Pj.barba > 0 Then
         ImagenBarba.setOrigin 0, 0, 32, 32
         ImagenBarba.setTextura Pj.barba, 128, 32
         ImagenBarba.SetPos centro, posicionCabeza
         ImagenBarba.IControl_SetSize 32, 32
    End If
     
    If Pj.ropaInterior > 0 Then
        ImagenRopaInterior.setOrigin 32, 0, 32, 64
        ImagenRopaInterior.setTextura Pj.ropaInterior, 256, 256
        ImagenRopaInterior.IControl_SetSize 32, 64
        ImagenRopaInterior.SetPos centro, posicionCuerpoY
     End If
     
     If Pj.pelo > 0 Then
          ImagenPelo.setOrigin 0, 0, 32, 32
          ImagenPelo.setTextura Pj.pelo, 128, 32
          ImagenPelo.SetPos centro, posicionCabeza
          ImagenPelo.IControl_SetSize 32, 32
     End If
    End If
End Sub

Private Sub actualizarPersonaje()
    
        
    Nombre.text = Pj.Nick
    
    If Pj.modoCandado Then
        indicadorModoCandado.grhId = GRH_MODOCANDADO_ON
    Else
        indicadorModoCandado.grhId = GRH_MODOCANDADO_OFF
    End If
    
    If Pj.online Then
        indicadorOnline.grhId = GRH_PERSONAJE_ON
    Else
        indicadorOnline.grhId = GRH_PERSONAJE_OFF
    End If
    
    If Pj.bloqueado = False Then
        Nombre.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
        fondo.Tooltip = ""
    Else
        Nombre.SetStyle 3, mzRed, mzInterfaceColor2, dVertical
        fondo.Tooltip = "Personaje bloqueado."
    End If
    
    Call actualizarPosiciones
    
End Sub

Private Sub actualizarPosiciones()
    fondo.SetPos CX, cY

    Nombre.SetPos CX + cW / 2, cY + cH - 8
    Nombre.SetCentrado True
    
    indicadorOnline.SetPos CX + cW - 32, cY + 7
    indicadorModoCandado.SetPos CX + cW - 16, cY + 7
    
    Call actualizarEstetica
    
    ReDoRECT
End Sub


Private Function IControl_Click(ByVal X As Integer, ByVal Y As Integer) As Boolean

End Function

Private Property Get IControl_IsFocuseable() As Boolean
    IControl_IsFocuseable = False
End Property

Private Property Get IControl_IsMouseFocus() As Boolean
    IControl_IsMouseFocus = False
End Property

Private Function IControl_KeyDown(KeyCode As Integer, Shift As Integer) As Boolean

End Function

Private Function IControl_KeyPress(KeyAscii As Integer) As Boolean

End Function

Private Function IControl_KeyUp(KeyCode As Integer, Shift As Integer) As Boolean

End Function

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
     IControl_MouseDown = ColisionRect(X, Y, ControlRect)
     
     If IControl_MouseDown Then
        Call callback.setPersonajeSeleccionado(Pj)
     End If
End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer) As Boolean
    vwControles.MouseMove Button, Shift, X, Y, z
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    vwControles.Render
End Sub

Public Sub IControl_SetPos(ByVal X As Integer, ByVal Y As Integer)
    CX = X
    cY = Y
    
    Call actualizarPosiciones

End Sub

Public Sub IControl_SetSize(ByVal ancho As Integer, ByVal alto As Integer)
    cW = ancho
    cH = alto
    
    ReDoRECT
End Sub

Private Sub ReDoRECT()

    With ControlRect
        .top = cY
        .bottom = cY + cH
        .left = CX
        .right = CX + cW
    End With
    
End Sub
