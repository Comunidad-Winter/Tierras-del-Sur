VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Eventos de los Usuarios que son tomados por el Evento
Public Sub usuarioMuere(UserIndex As Integer)
End Sub

Public Sub usuarioCerro(UserIndex As Integer)
End Sub
Public Sub usuarioAbandono(UserIndex As Integer)
End Sub

Public Sub usuarioIngreso(UserIndex As Integer, UserID As Long)
End Sub

Public Sub usuarioRevive(UserIndexRevivido As Integer, UserIndexRevividor As Integer)
End Sub

'*****************************************************************************
' Limite de acciones que pueden realizar los usuarios durante el evento      '
' ****************************************************************************

' ¿El usuario puede lanzar un hechizo? ' TODO tener en cuenta el usuario destino
Public Function puedeLanzar(indexHechizo As Byte) As Boolean
End Function

' ¿El usuario puede tirar un objeto?
Public Function puedeTirarObjeto(UserIndex As Integer, Objeto As Integer, ByVal cantidad As Long, destino As eDestinoObjeto, destinoIndex As Integer) As Boolean
End Function

'*****************************************************************************

Public Sub cancelar()
End Sub

Public Sub publicar()
End Sub

Public Function agregarEquipo(integrantesIndexs() As Integer) As eResultadoInscripcion
End Function

Public Function obtenerInfoExtendida() As String
End Function

Public Function obtenerInfoEstado() As String
End Function

Public Function getDescripcion() As String
End Function
Public Function getNombre() As String
End Function

Public Function getIDGanador() As Byte
End Function

Public Function getEstadoEvento() As eEstadoEvento
End Function

Public Function getTipoEvento() As eTipoEvento
End Function
Public Sub setHechizosPermitidos(reglas() As Boolean)
End Sub
'Esta es para debug
Public Function getTimeTranscurrido() As Integer
End Function

Public Function getCantidadParticipantesEquipo() As Byte
End Function

Public Function getCantidadEquiposMax() As Byte
End Function

Public Function isInscripcionesAbiertas() As Boolean
End Function

Public Function isCumpleCondicionEquipo(userIndexs() As Integer) As Boolean
End Function

'Procedimientos de timeOut. Se llama cada X tiempo
Public Sub timeOutMinuto()
End Sub

'Aún sin uso
Public Sub timeOutSegundo()
End Sub

Public Sub agregarCondicionIngreso(condicion As iCondicionEvento)
End Sub

Public Sub setHojaDeRuta(hojaRuta As String, parametros As String)
End Sub

Public Sub configurarApuestas(ByVal activadas As Boolean, ByVal pozoInicial As Long, ByVal tiempo As Byte)
End Sub

Public Sub establecerTablaDePagos(tablaPagos() As Long, tipoExpresion As eEventoPremio)
End Sub

Public Sub setTiporing(tipoRing_ As eRingTipo)
End Sub
'*****************************************************************************
' Persistencia
Public Sub guardar()
End Sub

Public Function obtenerIDPersistencia() As Integer
End Function

Public Sub setIDPersistencia(IDPersistencia As Long)
End Sub

'**************************************************************************************
Public Function establecerGanadorManualmente(nombreEquipoGanador As String) As Boolean
End Function

