Attribute VB_Name = "SV_Simulador"
Option Explicit

Private CharsInactivos As ColaConBloques
Private EntidadDisponibles As ColaConBloques

Public Sub InitChars()
    Dim i As Integer
    
    Set CharsInactivos = New ColaConBloques

    For i = MaxChar To 1 Step -1
        CharsInactivos.agregar i
    Next
    
    Set EntidadDisponibles = New ColaConBloques

    For i = 200 To 1 Step -1
        EntidadDisponibles.agregar i
    Next
    
End Sub

Public Function NextOpenChar(Optional ByVal NoUserCharIndex As Boolean = False) As Integer 'SOLO SE USA EN EL ME
    NextOpenChar = MaxChar
    
    'Quedan chars inactivos?
    If CharsInactivos.getCantidadElementos() > 0 Then
    
        NextOpenChar = CharsInactivos.sacar()
        Debug.Print "Abro el char ", NextOpenChar
    End If
       
End Function

Public Sub EraseIndexChar(ByVal Charindex As Integer)
    If Charindex <> UserCharIndex Then
        Call CharsInactivos.agregar(Charindex)
    End If
    
    Debug.Print "Elimino el char ", Charindex
End Sub

Public Sub EliminarIDEntidad(id As Integer)
    Call EntidadDisponibles.agregar(id)
    Debug.Print "Libero Entidad ID " & id
End Sub
Public Function ObtenerIDEntidad() As Integer
    'Quedan chars inactivos?
    If EntidadDisponibles.getCantidadElementos() > 0 Then
        ObtenerIDEntidad = EntidadDisponibles.sacar()
    End If
    
    Debug.Print "Entidad ID " & ObtenerIDEntidad
End Function


