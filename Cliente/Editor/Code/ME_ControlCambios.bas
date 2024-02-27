Attribute VB_Name = "ME_ControlCambios"
Option Explicit

' Asi nomas! Recordar que no es bueno el new en la declaracion!
Public pendientes As New Collection
Public cambiosPendientes As Boolean

Public Sub SetHayCambiosSinActualiar(tipo As String)
    cambiosPendientes = True
    
    Call pendientes.Add(tipo)
    
    frmMain.mnuCambiosPendientes.visible = True
    frmMain.mnuCambiosPendientes.Enabled = True
End Sub

Public Function hayCambiosSinActualizar() As Boolean
    hayCambiosSinActualizar = cambiosPendientes
End Function

Public Function hayCambiosSinActualizarDe(tipo As String) As Boolean
    Dim loopC As Byte
    
    loopC = buscar(pendientes, tipo)
    
    If loopC > 0 Then
        hayCambiosSinActualizarDe = True
    Else
        hayCambiosSinActualizarDe = False
    End If
End Function

Public Sub SetCambioActualizado(tipo As String)
    cambiosPendientes = True
    
   Call remover(tipo)
    
    If pendientes.Count = 0 Then
        frmMain.mnuCambiosPendientes.visible = False
        frmMain.mnuCambiosPendientes.Enabled = False
    End If
End Sub

Private Function buscar(coleccion As Collection, elemento As Variant) As Integer
    Dim loopC As Byte
    
    For loopC = 1 To coleccion.Count
        If coleccion.Item(loopC) = elemento Then
            buscar = loopC
            Exit Function
        End If
    Next
    buscar = 0
End Function

Private Sub remover(elemento As Variant)
    Dim loopC As Byte
    
    loopC = buscar(pendientes, elemento)
    
    If loopC > 0 Then
        Call pendientes.Remove(loopC)
    End If
End Sub
