VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoNoRepPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private idUsuario() As Long
Private cantidadPjsRegistrados As Integer

Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
    Dim posicionLibre As Byte
    Dim loopUsuario As Byte
    cantidadPjsRegistrados = cantidadPjsRegistrados + 1
    
    For loopUsuario = 1 To UBound(integrantesIndex)
        For posicionLibre = 1 To UBound(idUsuario)
            If idUsuario(posicionLibre) = 0 Then
                idUsuario(posicionLibre) = UserList(integrantesIndex(loopUsuario)).id
                Exit For
            End If
        Next
    Next
    
    
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    iCondicionEvento_obtenerDescripcion = "No puede estar el mismo personaje en dos equipos."
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean

    Dim UserIndex As Integer
    Dim loopC As Byte
    Dim loopYaInscripto As Integer
    Dim UserID As Long
    
    For loopC = 1 To UBound(integrantesIndex)
        UserIndex = integrantesIndex(loopC)
        UserID = UserList(UserIndex).id
        
        For loopYaInscripto = 1 To UBound(idUsuario)
            If idUsuario(loopYaInscripto) = UserID Then
                iCondicionEvento_puedeIngresarEquipo = False
                Exit Function
            End If
        Next
    Next loopC
    
    iCondicionEvento_puedeIngresarEquipo = True
    
End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantesid() As Long)
    Dim posicionLibre As Byte
    Dim loopIntegrante As Byte
    
    cantidadPjsRegistrados = cantidadPjsRegistrados - 1
    
    For loopIntegrante = 1 To UBound(integrantesid)
        For posicionLibre = 1 To UBound(idUsuario)
            If idUsuario(posicionLibre) = integrantesid(loopIntegrante) Then
                idUsuario(posicionLibre) = 0
                Exit For
            End If
        Next
    Next
            
End Sub

Public Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
    ReDim idUsuario(1 To slots) As Long
    cantidadPjsRegistrados = 0
End Sub
