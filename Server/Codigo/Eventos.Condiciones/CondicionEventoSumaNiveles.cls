VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoSumaNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private sumamax As Integer


Public Sub setParametros(ByVal sumamax_ As Integer)
    sumamax = sumamax_
End Sub
Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
    'No necesitamos guardar nada de info
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    iCondicionEvento_obtenerDescripcion = "La suma de los niveles de cada integrante del equipo no puede superar " & sumamax & "."
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean

    Dim loopC As Byte
    Dim UserIndex As Integer
    Dim sumatoria As Integer
    
    sumatoria = 0
    
    ' Sumamos!!
    For loopC = 1 To UBound(integrantesIndex)
    
        UserIndex = integrantesIndex(loopC)
        
        sumatoria = sumatoria + UserList(UserIndex).Stats.ELV

    Next loopC
    
    'Chequeamos!
    If sumatoria <= sumamax Then
        iCondicionEvento_puedeIngresarEquipo = True
    Else
        iCondicionEvento_puedeIngresarEquipo = False
    End If
    
End Function


Private Sub iCondicionEvento_removerEquipoInscripto(integrantes() As Long)
    'Como no guardamos nada no necesitamos borrar nada
End Sub

Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
    'No necesitamos hacer nada
End Sub
