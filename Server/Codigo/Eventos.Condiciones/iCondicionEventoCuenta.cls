VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iCondicionEventoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private seNecesitaPremium As Boolean

Public Sub setParametros(seNecesitaPremium_ As Boolean)
    seNecesitaPremium = seNecesitaPremium_
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    If seNecesitaPremium Then
        iCondicionEvento_obtenerDescripcion = "Solo personajes adheridos a una Cuenta Premium."
    Else
        iCondicionEvento_obtenerDescripcion = "Solo personajes adheridos a una Cuenta (Estándar o Premium)."
    End If
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean

    Dim loopC As Byte
    Dim UserIndex As Integer
    
    For loopC = 1 To UBound(integrantesIndex)
    
        UserIndex = integrantesIndex(loopC)
        
        If UserList(UserIndex).IDCuenta = 0 Then
            iCondicionEvento_puedeIngresarEquipo = False
            Exit Function
        ElseIf seNecesitaPremium And Not UserList(UserIndex).Premium Then
            iCondicionEvento_puedeIngresarEquipo = False
            Exit Function
        End If
        
    Next loopC
    
    iCondicionEvento_puedeIngresarEquipo = True
    
End Function


Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
End Sub
Private Sub iCondicionEvento_removerEquipoInscripto(integrantesid() As Long)
End Sub
Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
End Sub
