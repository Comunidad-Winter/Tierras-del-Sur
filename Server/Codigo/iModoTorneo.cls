VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iModoTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub setParametros(parametros As String)
End Sub

Public Sub setNecesidadRoundGanadosGanador(cantidad As Byte)
End Sub
 
Friend Sub setTablaEquipos(tabla() As tEquipoTablaTorneo)
End Sub

Friend Sub setRings(ring() As tRing)
End Sub

Friend Sub setRing(ring As tRing)
End Sub

'Establece un nombre para el evento
Public Sub setMaximaCantidadAdvertencias(ByVal cantidad As Byte)
End Sub

'Establece un nombre para el evento
Public Sub setNombreEvento(nombre_ As String)
End Sub

' Devuelve la tabla de equipos.
Friend Function obtenerTabla() As tEquipoTablaTorneo()
End Function

Friend Function obtenerEquipo(idEquipo As Byte) As tEquipoTablaTorneo
End Function

Public Sub comenzar(delay As Byte)
End Sub

Public Sub cancelar()
End Sub
' Devuelve true si el sub evento termino
Public Function isTerminado() As Boolean
End Function
' *************************************************************************************
' Eventos de los usuarios
Public Sub usuarioIngreso(UserINdex As Integer, UserID As Long)
End Sub

Public Sub usuarioAbandono(UserINdex As Integer)
End Sub

Public Sub usuarioCerro(UserINdex As Integer)
End Sub

Public Sub usuarioMuere(UserINdex As Integer)
End Sub

Public Sub usuarioRevive(UserIndexRevivido As Integer, UserIndexRevividor As Integer)
End Sub

' *************************************************************************************
'Eventos automaticos
Public Sub pasarMinuto()
End Sub

Public Sub cargar(idliga As Long)
End Sub

Public Sub guardar()
End Sub

Public Function obtenerIDPersistencia() As Integer
End Function

