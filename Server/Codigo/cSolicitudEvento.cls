VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSolicitudEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private evento As iEvento  'Evento al cual esta relacionado la solicitud
Private cantidadSolicitudesAceptadas As Byte
Private participantesID() As Long 'ID de las personas que conforman la solicitud. El 1 es el lider
Private isSolicitudAceptada() As Byte '1 a 1 con los participantes. Si es 1 quiere decir que acepto. 0 no acepto aún.


Public Sub crear(ByRef evento_ As iEvento)

    Set evento = evento_
    
    ReDim participantesID(1 To evento.getCantidadParticipantesEquipo) As Long
    ReDim isSolicitudAceptada(1 To evento.getCantidadParticipantesEquipo) As Byte
    
    cantidadSolicitudesAceptadas = 1    'El que lo creo
    isSolicitudAceptada(1) = 1 'El que lo creo
    
End Sub

Public Sub setEquipo(IDParticipantes_() As Long)
    participantesID = IDParticipantes_
End Sub

'0. Solicitud aceptada.
'1. Solicitud aceptada, comenzo el evento.
'2. Nadie lo invito
'3. Ya habia aceptado
'4. El equipo no esta completo.

Private Function inscribirEquipo() As Boolean

'Armo el equipo con las indexs
Dim cantidadParticipantes As Byte
Dim IntegranteIndexs() As Integer
Dim loopParticipante As Byte

Dim hayError As Boolean
Dim cantidadParticipantesOff As Byte
Dim cantidadParticipantesNoListos As Byte
Dim resultadoInscripcion As eResultadoInscripcion

cantidadParticipantes = UBound(participantesID)

ReDim IntegranteIndexs(1 To cantidadParticipantes) As Integer
              
'Chequeo que todo el equipo este online.
'Chequeo que cumplan las condiciones para jugar un evento
'Chequeo que cumplan las condiciones del evento en particiular
hayError = False
cantidadParticipantesOff = 0
cantidadParticipantesNoListos = 0

For loopParticipante = 1 To cantidadParticipantes

    IntegranteIndexs(loopParticipante) = IDIndex(participantesID(loopParticipante))

    If IntegranteIndexs(loopParticipante) > 0 Then
        If Not puedeParticipar(IntegranteIndexs(loopParticipante)) Then
            isSolicitudAceptada(loopParticipante) = 0 'Va a tener que aceptar de nuevo la soli

            hayError = True
            cantidadParticipantesNoListos = cantidadParticipantesNoListos + 1
        End If
    Else
        isSolicitudAceptada(loopParticipante) = 0 'Va a tener que volver a aceptar
        'Registro el error
        hayError = True
        cantidadParticipantesOff = cantidadParticipantesOff + 1
    End If

Next loopParticipante

'Sino hay error me fijo si el equipo entero cumple las condiciones
If Not hayError Then
    If evento.isCumpleCondicionEquipo(IntegranteIndexs) Then
    
        resultadoInscripcion = evento.agregarEquipo(IntegranteIndexs)
        
        If resultadoInscripcion = eResultadoInscripcion.correcta Then
            hayError = False
        Else
            If resultadoInscripcion = noTieneOro Then
                Call enviarMensaje(IntegranteIndexs, "No se tiene el oro necesario para participar.")
            Else
                Call LogTorneos("ERROR. " & evento.getNombre & " al momento de incribirse un equipo.")
            End If
            hayError = True
        End If
    Else
        hayError = True
        Call enviarMensaje(IntegranteIndexs, "El equipo no cumple las condiciones para participar de este evento.")
    End If
Else
    If cantidadParticipantesOff > 0 Then
        If cantidadParticipantesOff = 1 Then
            Call enviarMensaje(IntegranteIndexs, "Un integrate del equipo no se encuentra online.")
        Else
            Call enviarMensaje(IntegranteIndexs, cantidadParticipantesOff & " integrantes no se encuentran online.")
        End If
    Else
        If cantidadParticipantesNoListos = 1 Then
            Call enviarMensaje(IntegranteIndexs, "Un integrate del equipo no se encuentra en condiciones para participar de un evento.")
        Else
            Call enviarMensaje(IntegranteIndexs, cantidadParticipantesNoListos & " integrantes no se encuentran en condiciones de participar en un evento.")
        End If
    End If
End If
inscribirEquipo = hayError
End Function
Public Function aceptarSolicitud(IDParticipante As Long) As Byte
    
    Dim loopParticipante As Byte
    Dim cantidadParticipantes As Byte
    
    cantidadParticipantes = UBound(participantesID)
    
    For loopParticipante = 1 To cantidadParticipantes
    
        If participantesID(loopParticipante) = IDParticipante Then
            
            If isSolicitudAceptada(loopParticipante) = 0 Then
                isSolicitudAceptada(loopParticipante) = 1
                cantidadSolicitudesAceptadas = cantidadSolicitudesAceptadas + 1
                
                'Aceptaron todos?
                If cantidadSolicitudesAceptadas = cantidadParticipantes Then
                    'El evento todavia tiene abiertas sus inscripciones?
                    If evento.isInscripcionesAbiertas() Then
                       If inscribirEquipo Then
                            aceptarSolicitud = 0
                       Else
                            aceptarSolicitud = 4
                        End If
                    Else 'inscripciones cerradas
                        aceptarSolicitud = 0
                    End If
                    Exit Function
                Else 'Debes esperar a que los demas acepten
                    aceptarSolicitud = 1
                    Exit Function
                End If
            Else 'Ya habia aceptado la solicitud
                aceptarSolicitud = 3
                Exit Function
            End If
        End If
    
    Next loopParticipante
    
    'Nadie lo invito
    aceptarSolicitud = 2
End Function

Public Function getNombreEvento() As String
    getNombreEvento = evento.getNombre
End Function

Public Function getEvento() As iEvento
    Set getEvento = evento
End Function

Private Sub enviarMensaje(userIndexs() As Integer, mensaje As String)
    Dim loopParticipante As Byte
    
    For loopParticipante = 1 To UBound(userIndexs)
        If userIndexs(loopParticipante) > 0 Then
            EnviarPaquete Paquetes.mensajeinfo, mensaje, userIndexs(loopParticipante), ToIndex
        End If
    Next
End Sub
