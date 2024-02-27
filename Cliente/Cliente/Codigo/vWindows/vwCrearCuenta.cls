VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements vWindow

Option Explicit

Private vX As Integer
Private vY As Integer
Private vW As Integer
Private vH As Integer

Private Visible As Boolean
Private vControles As vControles

Private WithEvents BotonCrearCuenta As ControlBoton
Attribute BotonCrearCuenta.VB_VarHelpID = -1
Private WithEvents BotonAtras As ControlBoton
Attribute BotonAtras.VB_VarHelpID = -1

Private InputNombre As ControlInput
Private InputMail As ControlInput
Private InputPassword As ControlInput

Private FondoVentana As ControlImage

Private FondoTituloPassword As ControlImage
Private TituloPassword As ControlLabel


Private FondoTituloMail As ControlImage
Private TituloMail As ControlLabel

Private FondoTituloNombre As ControlImage
Private TituloNombre As ControlLabel

Private lineaSeparadoraNombre As ControlImage
Private lineaSeparadora As ControlImage

Private LabelAdvertencia As ControlLabel

Private LabelError As ControlLabel

Public Event Atras()
Public Event CrearCuenta(nombreCuenta As String, Password As String, Email As String)

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow

Private Sub BotonAtras_Click()
    Call Sonido_Play(SND_CLICK)
    RaiseEvent Atras
End Sub

Private Function validar() As Boolean
    Dim Nombre As String
    Dim Password As String
    Dim Email As String
    
    Nombre = Trim$(InputNombre.text)
    Password = Trim$(InputPassword.text)
    Email = Trim$(InputMail.text)
    
    If Len(Password) < 5 Then
        mostrarError 0, "La contraseña debe tener al menos 5 letras. Es por tu seguridad."
        Exit Function
    End If
    
    If Len(Nombre) < 3 Then
        mostrarError 0, "El nombre debe tener al menos 3 letras."
        Exit Function
    End If
    
    If Len(Nombre) > 12 Then
        mostrarError 0, "La cuenta debe tener como máximo 12 letras."
        Exit Function
    End If
    
    If Not CheckMailString(Email) Then
        mostrarError 0, "El correo electrónico no parece ser válido."
        Exit Function
    End If

    validar = True
End Function

Private Sub BotonCrearCuenta_Click()
    If Not validar Then Exit Sub

    Call Sonido_Play(SND_CLICK)
    
    RaiseEvent CrearCuenta(InputNombre.text, InputPassword.text, InputMail.text)
End Sub

Private Sub Class_Initialize()
    Set vControles = New vControles
    
    Set InputNombre = vControles.Add(New ControlInput)
    Set InputPassword = vControles.Add(New ControlInput)
    Set InputMail = vControles.Add(New ControlInput)
    
    Set BotonCrearCuenta = vControles.Add(New ControlBoton)
    Set BotonAtras = vControles.Add(New ControlBoton)
        
    Set TituloPassword = vControles.Add(New ControlLabel)
    Set TituloMail = vControles.Add(New ControlLabel)
    Set TituloNombre = vControles.Add(New ControlLabel)
    
    Set FondoTituloPassword = vControles.Add(New ControlImage)
    Set FondoTituloMail = vControles.Add(New ControlImage)
    Set FondoTituloNombre = vControles.Add(New ControlImage)
    
    Set lineaSeparadora = vControles.Add(New ControlImage)
    Set lineaSeparadoraNombre = vControles.Add(New ControlImage)
    
    Set FondoVentana = vControles.Add(New ControlImage)
    
    FondoTituloPassword.grhId = 21868
    FondoTituloMail.grhId = 21868
    FondoTituloNombre.grhId = 21868
    
    InputNombre.MaxLenght = 12
    InputNombre.TipoTexto = ALFANUMERICO
    
    InputMail.MaxLenght = 32
    InputMail.TipoTexto = ALFANUMERICO
    
    InputPassword.MaxLenght = 20
    InputPassword.TipoTexto = Contraseña
    
    TituloNombre.text = "Nombre de la cuenta"
    TituloNombre.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloNombre.SetCentrado True
    
    TituloPassword.text = "Elegi una clave"
    TituloPassword.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloPassword.SetCentrado True
    
    TituloMail.text = "Escribi tu correo electronico"
    TituloMail.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloMail.SetCentrado True
    
    FondoVentana.grhId = 20352
    lineaSeparadora.grhId = 21867
    lineaSeparadoraNombre.grhId = 21867
    
    InputPassword.MaxLenght = 30
    InputPassword.TipoTexto = Contraseña
    
    InputMail.MaxLenght = 40
    InputMail.TipoTexto = ALFANUMERICO
    
    vControles.ObjetoFocus = InputNombre
    
    ' Botones de accion
    BotonAtras.text = "Volver"
    BotonCrearCuenta.text = "Crear Cuenta"
End Sub

Private Sub Class_Terminate()
    Debug.Print "Terminate"
End Sub

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = vControles
End Property
Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function

Private Property Get vWindow_IsVisible() As Boolean
    vWindow_IsVisible = Visible
End Property

Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub

Public Sub vWindow_Hide()
    Visible = False
End Sub

Public Sub vWindow_Show()
    Visible = True
End Sub

Private Function vWindow_Render() As Boolean
    If Visible Then vControles.Render
    
    vWindow_Render = True
End Function


Public Sub mostrarError(error As Byte, errordesc As String)
    If error = modLogin.JUEGO_DESACTUALIZADO Then
        Call juegoDesactualizado
    End If
    
    If LabelError Is Nothing Then
        Set LabelError = vControles.AddFirst(New ControlLabel)
    End If
   
    LabelError.SetPos vX + vW / 2, vY + vH - 35
    LabelError.Color = mzRed
    LabelError.SetCentrado True
    LabelError.SetSize vW - 50, vY
    
    ' Y el texto
    If errordesc = "" And error > 0 Then
        LabelError.text = modLogin.erroresDescripcion(error)
    ElseIf error > 0 Then
        LabelError.text = modLogin.erroresDescripcion(error) & " Error: " & errordesc
    Else
        LabelError.text = errordesc
    End If
End Sub

Public Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
    
    Call actualizarPosiciones
End Sub

Public Sub vWindow_SetSize(ByVal width As Integer, ByVal Height As Integer)
    vW = width
    vH = Height
    
    Call actualizarPosiciones
End Sub

Private Sub actualizarPosiciones()
    FondoVentana.IControl_SetSize vW, vH
    FondoVentana.SetPos vX, vY
    
    Dim base As Integer
    
    base = 0
    
    FondoTituloNombre.SetPos vX + vW / 2 - TituloNombre.width / 2, vY + 2
    FondoTituloNombre.IControl_SetSize TituloNombre.width, 32
    
    TituloNombre.SetPos vX + vW / 2, vY + base + 7
       
    InputNombre.SetPos vX + vW / 2 - 300 / 2, vY + base + 30
    InputNombre.SetSize 300, 26
    
    lineaSeparadoraNombre.SetPos vX, vY + 90
    lineaSeparadoraNombre.IControl_SetSize vW, 10

    ' Inputs
    base = 90
    
    FondoTituloPassword.SetPos vX + vW / 2 - TituloPassword.width / 2, vY + base + 5
    FondoTituloPassword.IControl_SetSize TituloPassword.width, 32
    
    TituloPassword.SetPos vX + vW / 2, vY + base + 7

    InputPassword.SetPos vX + vW / 2 - 300 / 2, vY + 30 + base
    InputPassword.SetSize 300, 26
    
    base = 180
    
    lineaSeparadora.SetPos vX, vY + base
    lineaSeparadora.IControl_SetSize vW, 10
    
    
   TituloMail.SetPos vX + vW / 2, vY + base + 7

    FondoTituloMail.SetPos vX + vW / 2 - TituloMail.width / 2, vY + base + 5
    FondoTituloMail.IControl_SetSize TituloMail.width, 32

    InputMail.SetPos vX + vW / 2 - 300 / 2, vY + base + 30
    InputMail.SetSize 300, 26
    
    ' Botones finales
    BotonCrearCuenta.SetPos vX + vW - 128, vY + vH + 10
    BotonAtras.SetPos vX, vY + vH + 10
    

    If Not LabelError Is Nothing Then
        LabelError.SetPos vX + vW / 2, vY + vH - 35
    End If
End Sub

Public Property Get vWindow_X() As Integer
    vWindow_X = vX
End Property

Public Property Get vWindow_Y() As Integer
    vWindow_Y = vY
End Property

Public Property Get vWindow_Width() As Integer
    vWindow_Width = vW
End Property

Public Property Get vWindow_Height() As Integer
    vWindow_Height = vH
End Property


