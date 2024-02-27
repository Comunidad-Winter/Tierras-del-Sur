Attribute VB_Name = "CLI_GUI"
Option Explicit

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Sub UnloadAllForms()
  'on error Resume Next
  Dim mifrm As Form

    For Each mifrm In Forms
        Unload mifrm
    Next

End Sub

'Cursores
'0=Manito Abierta
'1=Manito cerrada
Public Sub CambiarCursor(Form As Form, Optional cursor As Byte)

    If CursorPer = 1 Then
        Form.MouseIcon = LoadResPicture(101 + cursor, vbResCursor)
        Form.MousePointer = 99
      Else
        Form.MousePointer = 1
    End If

End Sub

Public Sub DameImagen(Control As Image, Imagen As Integer)
   Control.Picture = clsEnpaquetado_LeerIPicture(pakGUI, Imagen)
   'Control.Picture = LoadPicture(app.Path & "\Recursos\Interface\" & Imagen & ".jpg")
End Sub

Public Sub DameImagenForm(Control As Form, Imagen As Integer)
    Control.Picture = clsEnpaquetado_LeerIPicture(pakGUI, Imagen)
    'Control.Picture = LoadPicture(app.Path & "\Recursos\interface\" & Imagen & ".jpg")
End Sub


Public Sub DameImagenPicture(Control As PictureBox, Imagen As Integer)
    Control.Picture = clsEnpaquetado_LeerIPicture(pakGUI, Imagen)
    'Control.Picture = LoadPicture(app.Path & "\Recursos\interface\" & Imagen & ".jpg")
End Sub


Public Sub MostrarFormulario(Formulario As Form, Padre As Form)
    Padre.Enabled = False
    Formulario.Show vbModeless, Padre
End Sub

