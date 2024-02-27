Attribute VB_Name = "modPosicionarFormulario"
Option Explicit


Public Sub posicionarAbajoCentro(padre As Form, hijo As Form)
    'Posiciono el formulario
    hijo.top = padre.top + padre.height - hijo.height
    hijo.left = padre.left + (padre.width \ 2) - (hijo.width \ 2)
End Sub

Public Sub posicionarAbajoDerecha(padre As Form, hijo As Form)
    'Posiciono el formulario
    hijo.top = padre.top + padre.height - hijo.height
    hijo.left = padre.left + padre.width - hijo.width
End Sub


Public Sub setEnabledHijos(estado As Boolean, padre As Control, Form As Form)
  Dim Control As Control
  
    For Each Control In Form.Controls
        If TypeName(Control) <> "Timer" Then
            On Error Resume Next
            If Control.Container Is padre Then
                If Err.Number = 0 Then
                    Control.Enabled = estado
                End If
            End If
            Err.Clear
        End If
    Next
End Sub

