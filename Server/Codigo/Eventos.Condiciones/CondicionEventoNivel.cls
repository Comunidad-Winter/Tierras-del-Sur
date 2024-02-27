VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
'Condiciones: personajes del equipo, restriccion de nivel de cada integrante.
'@see
Option Explicit

Implements iCondicionEvento

Private minimoNivel As Byte
Private maximoNivel As Byte

Private Sub Class_Initialize()
    minimoNivel = 1
    maximoNivel = STAT_MAXELV
End Sub

Public Sub setParametros(ByVal minimo As Byte, ByVal maximo As Byte)

    minimoNivel = minimo
    maximoNivel = maximo
    
End Sub
Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
    'No necesitamos guardar nada de info
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    If minimoNivel = maximoNivel Then
        iCondicionEvento_obtenerDescripcion = "Personajes de nivel " & minimoNivel & "."
    Else
        iCondicionEvento_obtenerDescripcion = "Personajes de nivel " & minimoNivel & " a " & maximoNivel & "."
    End If
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean

    Dim loopC As Byte
    Dim UserIndex As Integer
    
    For loopC = 1 To UBound(integrantesIndex)
    
        UserIndex = integrantesIndex(loopC)
        
        If UserList(UserIndex).Stats.ELV < minimoNivel Or UserList(UserIndex).Stats.ELV > maximoNivel Then
            iCondicionEvento_puedeIngresarEquipo = False
            Exit Function
        End If
        
    Next loopC
    
    iCondicionEvento_puedeIngresarEquipo = True
End Function


Private Sub iCondicionEvento_removerEquipoInscripto(integrantes() As Long)
    'Como no guardamos nada no necesitamos borrar nada
End Sub

Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
    'No necesitamos hacer nada
End Sub
