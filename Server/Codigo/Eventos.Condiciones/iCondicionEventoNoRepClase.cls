VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoNoRepClase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private maxClasesRepetidas As Byte

Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)

End Sub

Public Sub setParametros(ByVal maxClasesRepetidas_ As Byte)
    maxClasesRepetidas = maxClasesRepetidas_
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    If maxClasesRepetidas = 0 Then
        iCondicionEvento_obtenerDescripcion = "No vale repetir clase."
    Else
        iCondicionEvento_obtenerDescripcion = "Solo se puede repetir clase " & maxClasesRepetidas & " veces."
    End If
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean
    Dim UserIndex As Integer
    Dim loopC As Byte
    Dim loopB As Byte
    
    Dim loopClanes As Byte
    Dim cantidadDeClase As Byte
    Dim clase As eClases
    
    For loopC = 1 To UBound(integrantesIndex)
    
        UserIndex = integrantesIndex(loopC)
        
        clase = UserList(UserIndex).clase
        cantidadDeClase = 0
        
        For loopB = 1 To UBound(integrantesIndex)
            
            If UserIndex <> integrantesIndex(loopB) Then
                If clase = UserList(integrantesIndex(loopB)).clase Then
                    cantidadDeClase = cantidadDeClase + 1
                    If cantidadDeClase > maxClasesRepetidas Then
                            iCondicionEvento_puedeIngresarEquipo = False
                        Exit Function
                    End If
                End If
            End If
            
        Next loopB
    Next loopC
    
    iCondicionEvento_puedeIngresarEquipo = True
    
End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantesid() As Long)
End Sub

Public Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
End Sub
