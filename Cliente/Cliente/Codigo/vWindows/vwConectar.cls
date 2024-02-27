VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwConectar"
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

Private visible As Boolean

Private LabelError As ControlLabel
' Superior
Private WithEvents Boton As ControlBoton
Attribute Boton.VB_VarHelpID = -1
Private WithEvents BotonCrear As ControlBoton
Attribute BotonCrear.VB_VarHelpID = -1

Private WithEvents Nombre As ControlInput
Attribute Nombre.VB_VarHelpID = -1
Private WithEvents Password As ControlInput
Attribute Password.VB_VarHelpID = -1
Private WithEvents recordarClave As ControlCheckBox
Attribute recordarClave.VB_VarHelpID = -1

Private WithEvents vControles As vControles
Attribute vControles.VB_VarHelpID = -1

Private ObjetoFocus As Object

Public Event conectar(usuario As String, Password As String, recordarClave As Boolean)
Public Event crear()

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow

Private Progreso As clsProgreso

Private Sub BotonSalir_Click()
    prgRun = False
End Sub


Private Sub emitirConectar()
    Call Sonido_Play(SND_CLICK)
    
    Call ocultarError
    
    If recordarClave Then
        Call GuardarPassword(Nombre.text, CryptStr(Password.text, 1))
    Else
        Call EliminarPassword(Nombre.text)
    End If
    
    RaiseEvent conectar(Trim$(Nombre.text), Password.text, recordarClave.value)
End Sub

Private Sub BotonCrear_Click()
    Call Sonido_Play(SND_CLICK)
        
    RaiseEvent crear
End Sub


Private Sub intentarCOnPasswordRecordada()
    If recordarClave.value And Password.text = "" Then
        Dim passwordGuardado As String
        
        passwordGuardado = BuscarPassword(Nombre.text)
        
        If Not passwordGuardado = vbNullString Then
            Password.text = passwordGuardado
            Call emitirConectar
        End If
        
    ElseIf Not Password.text = "" Then
        Call emitirConectar
    End If
    
    
    
End Sub

Private Sub Nombre_KeyPress(ByVal KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        intentarCOnPasswordRecordada
    End If
End Sub

Private Sub Password_Click(ByVal X As Integer, ByVal Y As Integer)
    Call intentarCOnPasswordRecordada
End Sub

Private Sub password_KeyPress(ByVal KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call intentarCOnPasswordRecordada
    End If
End Sub

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = vControles
End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function

Private Property Get vWindow_IsVisible() As Boolean
    vWindow_IsVisible = visible
End Property

    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub

'/Implementacion

Private Sub Class_Terminate()
 
Set LabelError = Nothing
Set Boton = Nothing
Set BotonCrear = Nothing
Set Nombre = Nothing
Set Password = Nothing
Set recordarClave = Nothing
Set vControles = Nothing
Set Progreso = Nothing

End Sub

Private Sub Class_Initialize()
Set vControles = New vControles

Set Nombre = vControles.Add(New ControlInput)
Set Password = vControles.Add(New ControlInput)
Set recordarClave = vControles.Add(New ControlCheckBox)
Set Boton = vControles.Add(New ControlBoton)
Set BotonCrear = vControles.Add(New ControlBoton)

recordarClave.Enabled = True
recordarClave.value = False
recordarClave.texto = "Recordar clave"

Boton.text = "Ingresar"

BotonCrear.text = "Crear Cuenta"

Nombre.SetStyle 2, mzInterfaceColor1, mzInterfaceColor2, dVertical
Nombre.text = ""
Nombre.Enabled = True
Nombre.label = "Cuenta"
Nombre.MaxLenght = 30
Nombre.TipoTexto = eGuiTiposTexto.ALFANUMERICO

Password.SetStyle 2, mzInterfaceColor1, mzInterfaceColor2, dVertical
Password.text = ""
Password.Enabled = True
Password.label = "Contraseña"
Password.MaxLenght = 30
Password.TipoTexto = eGuiTiposTexto.Contraseña

recordarClave.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical

vControles.ObjetoFocus = Nombre

' Aparece el recordar clave tildado por defecto?
recordarClave.value = Configuracion_Usuario.Recpassword

End Sub

Private Sub Boton_Click()
    emitirConectar
End Sub

Public Sub vWindow_Hide()
    visible = False
End Sub

Private Function vWindow_Render() As Boolean
    vWindow_Render = True
    
    
    If Progreso.isCompleto = False Then
        BotonCrear.setAlpha Progreso.Calcular
        Boton.setAlpha Progreso.Calcular
        recordarClave.setAlpha Progreso.Calcular
        Password.setAlpha Progreso.Calcular
        Nombre.setAlpha Progreso.Calcular
    End If
    
    If visible Then vControles.Render
End Function

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

Public Sub vWindow_Show()
    visible = True
    Set Progreso = New clsProgreso
    Progreso.SetRango 0, 255
    Progreso.SetTicks GetTimer + 1000
End Sub

Private Sub ocultarError()
    If LabelError Is Nothing Then
        Exit Sub
    End If
    
    Dim Control As IControl
    Set Control = LabelError
    Call vControles.RemoveControl(LabelError)
    
    Set LabelError = Nothing
End Sub
Public Sub mostrarError(error As Byte, errordesc As String)
    If error = modLogin.JUEGO_DESACTUALIZADO Then
        Call juegoDesactualizado
    End If
    
    If LabelError Is Nothing Then
        Set LabelError = vControles.Add(New ControlLabel)
    End If
    
    LabelError.Color = mzRed
    LabelError.SetCentrado False
    LabelError.SetSize 200, vY
    
    ' Y el texto
    If errordesc = "" Then
        LabelError.text = modLogin.erroresDescripcion(error)
    ElseIf error > 0 Then
        LabelError.text = modLogin.erroresDescripcion(error) & " Error: " & errordesc
    Else
        LabelError.text = errordesc
    End If
    
    Call actualizarPosiciones
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

Private Sub actualizarPosiciones()
    
Boton.SetSize vW, 32
BotonCrear.SetSize vW, 32

Nombre.SetPos vX, vY
Nombre.SetSize vW, 26

Password.SetPos vX, vY + 50
Password.SetSize vW, 26

recordarClave.SetPos vX, vY + 100
Boton.SetPos vX, vY + 140
BotonCrear.SetPos Boton.GetPosX, Boton.GetPosY + 40

If Not LabelError Is Nothing Then
    LabelError.SetPos vX, BotonCrear.GetPosY + BotonCrear.IControl_Height + 10
End If

End Sub
