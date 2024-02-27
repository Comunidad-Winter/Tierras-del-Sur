VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoClases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private clasesPermitidas As Long

Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
End Sub


Public Sub setParametros(clasesPermitidas_ As Long)
    clasesPermitidas = clasesPermitidas_
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    iCondicionEvento_obtenerDescripcion = "Clases permitidas:" & clasesToString(clasesPermitidas) & "."
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean

Dim loopIntegrante As Byte

If clasesPermitidas = eClases.indefinido Then
    iCondicionEvento_puedeIngresarEquipo = True
    Exit Function
End If

' Recorro cada uno de los integrantes, si alguno tiene una clase no permitida no puede jugar el team.
For loopIntegrante = 1 To UBound(integrantesIndex)
    If (clasesPermitidas And UserList(integrantesIndex(loopIntegrante)).clase) = 0 Then
        iCondicionEvento_puedeIngresarEquipo = False
        Exit Function
    End If
Next loopIntegrante

'Todos los integrantes tiene las clases permitidas
iCondicionEvento_puedeIngresarEquipo = True

End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantes() As Long)
End Sub

Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
End Sub
