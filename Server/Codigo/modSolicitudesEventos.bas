Attribute VB_Name = "modSolicitudesEventos"
Option Explicit


'Crea una solicitud o si el evento es de un solo usuario, directamnete lo inscribe en el evento
'Formato NombreEvento-Nick Compañero 1 - NickCompañero dos
Public Sub crear(liderIndex As Integer, ByRef infoCreacion As String)

Dim info() As String
Dim evento As iEvento
Dim cantidadParticipantes As Byte
Dim loopParticipante As Byte
Dim loopC As Byte
Dim UserIndex As Integer
Dim participantesID() As Long
Dim participantesIndexs() As Integer

Dim presentacionString As String 'Descripcion de los compañeros que se juntan.
Dim presentacionCompaneros As String 'Nombres de los companeros para el que crea la solicitud.

Dim hayError As Boolean
'Al menos deb

If Len(infoCreacion) > 0 Then
    'Parseo la información
    info = Split(infoCreacion, "-")

    'Obtengo el vento con este nombre
     Set evento = modEventos.getEventoByNombre(info(0))

    'El evento existe?
    If Not evento Is Nothing Then
        'El evento tiene las inscripciones abiertas?
        If evento.isInscripcionesAbiertas Then
        
            cantidadParticipantes = UBound(info) + 1 '+1 porque empiza en 0. No pongo un -1 para sacar el nombre del evento, porque también suma el que esta mandando esto
            
            'La cantidad de pariticpantes es igual a la cantidad que solicita al evento porequipo?
            If cantidadParticipantes = evento.getCantidadParticipantesEquipo Then
                If cantidadParticipantes = 1 Then
                
                    If puedeParticipar(liderIndex) Then
                        'Agrego al usuario al evento
                        ReDim participantesIndexs(1) As Integer
                        participantesIndexs(1) = liderIndex
                        
                        If evento.isCumpleCondicionEquipo(participantesIndexs) Then
                            Call evento.agregarEquipo(participantesIndexs)
                        Else
                            EnviarPaquete Paquetes.MensajeTalk, "No cumples las condiciones necesarias para participar de este evento.", liderIndex, ToIndex
                        End If
                    End If
                    
                Else
                    
                    ReDim participantesIndex(1 To cantidadParticipantes) As Integer
                    'Debo crear una solicitud para los otros usuarios
                   
                    hayError = False
                    participantesIndex(1) = liderIndex
                    
                    'Cadena que voy a utilizar para presentarle el evento a los compañeros
                    
                    presentacionString = UserList(liderIndex).Name & " te invita a participar del evento " & evento.getNombre & " junto a "
                    presentacionCompaneros = ""
                    
                    For loopParticipante = 1 To cantidadParticipantes - 1
                    
                            If Not UCase$(info(loopParticipante)) = UCase$(UserList(liderIndex).Name) Then
                                UserIndex = NameIndex(info(loopParticipante))
                            
                                'Esta online?
                                If UserIndex > 0 Then
                                   'Esta en zona segura?
                                    If MapInfo(UserList(UserIndex).pos.map).Pk = False Then
                                    
                                        'Veo si puso dos veces el mismo nick
                                        For loopC = 1 To loopParticipante - 1
                                            
                                            If participantesIndex(loopC) = UserIndex Then
                                                EnviarPaquete Paquetes.mensajeinfo, "En la lista de tus compañeros has repetido al personaje " & UserList(UserIndex).Name, liderIndex, ToIndex
                                                Exit Sub 'Aca si termino
                                            End If
                                        
                                        Next loopC
                                    
                                        'Esta todo ok
                                        If Not hayError Then
                                        
                                            participantesIndex(loopParticipante + 1) = UserIndex
                                            
                                            If loopParticipante < cantidadParticipantes - 2 Then
                                                presentacionString = presentacionString & UserList(UserIndex).Name & ", "
                                                presentacionCompaneros = presentacionCompaneros & UserList(UserIndex).Name & ", "
                                            ElseIf loopParticipante = cantidadParticipantes - 2 Then
                                                presentacionString = presentacionString & UserList(UserIndex).Name & " y "
                                                presentacionCompaneros = presentacionCompaneros & UserList(UserIndex).Name & " y "
                                            Else
                                                presentacionString = presentacionString & UserList(UserIndex).Name
                                                presentacionCompaneros = presentacionCompaneros & UserList(UserIndex).Name
                                            End If
                                        End If
                                    Else
                                        hayError = True
                                        EnviarPaquete Paquetes.mensajeinfo, "El personaje " & info(loopParticipante) & " no puede recibir tu invitación ya que esta en una zona insegura.", liderIndex, ToIndex
                                    End If
                                Else
                                    hayError = True
                                    EnviarPaquete Paquetes.mensajeinfo, "El personaje " & info(loopParticipante) & " no se encuentra online.", liderIndex, ToIndex
                                End If
                            Else
                                hayError = True
                                EnviarPaquete Paquetes.mensajeinfo, "No debés ponerte a ti mismo en la lista de compañeros.", liderIndex, ToIndex
                                Exit Sub
                            End If
                    Next loopParticipante
                    
                    If hayError Then Exit Sub
                    
                    'Termino de generar el mensaje de presentacion del evento
                    presentacionString = presentacionString & ". Para aceptar escribe /ACEPTAR " & UserList(liderIndex).Name & ". Y /INFO " & evento.getNombre & " para obtener información sobre el evento."
                    Debug.Print presentacionString
                    
                    
                    'Obtengo el userIndex de cada participante y le mando el mensaje
                    ReDim participantesID(1 To cantidadParticipantes) As Long
                    participantesID(1) = UserList(participantesIndex(1)).id
                    
                   'Empiezo desde el segundo para evitar al lider
                    For loopParticipante = 2 To cantidadParticipantes
                    
                        participantesID(loopParticipante) = UserList(participantesIndex(loopParticipante)).id
                        
                        EnviarPaquete Paquetes.MensajeTalk, presentacionString, participantesIndex(loopParticipante), ToIndex
                    Next loopParticipante
        
                    'Los participantes estan ok... creo la solicitud para que la puedan aceptar
                    Set UserList(liderIndex).solicitudEvento = New cSolicitudEvento
                    Call UserList(liderIndex).solicitudEvento.crear(evento)
                    Call UserList(liderIndex).solicitudEvento.setEquipo(participantesID)
                    
                    'Le avisamos que esta todo ok.
                    If cantidadParticipantes = 2 Then ' El lider más uno
                        EnviarPaquete Paquetes.mensajeinfo, "Has invitado a " & presentacionCompaneros & " al " & info(0) & ". Ahora debes esperar que acepte tu invitación.", liderIndex, ToIndex
                    Else
                        EnviarPaquete Paquetes.mensajeinfo, "Has invitado a " & presentacionCompaneros & " al " & info(0) & ". Ahora debes esperar que todos ellos acepten.", liderIndex, ToIndex
                    End If
                End If
            Else 'No puso la cantidad necesaria
                If evento.getCantidadParticipantesEquipo = 1 Then
                    EnviarPaquete Paquetes.mensajeinfo, "Este torneo es individual. No se aceptan compañeros y no debes ponerte a vos mismo en la lista de compañeros. Simplemente es /PARTICIPAR " & info(0) & ".", liderIndex, ToIndex
                Else
                    EnviarPaquete Paquetes.mensajeinfo, "Debes seleccionar al menos " & (evento.getCantidadParticipantesEquipo - 1) & " compañeros para participar del " & info(0) & ".", liderIndex, ToIndex
                End If
            End If
        Else 'El evento ya cerro su inscripcion
            EnviarPaquete Paquetes.mensajeinfo, "El evento en el cual deseas participar no tiene abiertas sus inscripciones.", liderIndex, ToIndex
        End If
    Else 'No existe ele vento
        EnviarPaquete Paquetes.mensajeinfo, "El evento " & info(0) & " no existe.", liderIndex, ToIndex
    End If
Else 'Mando mal el comando
    EnviarPaquete Paquetes.mensajeinfo, "Debe enviar al menos el nombre del evento al cual desea participar.", liderIndex, ToIndex
End If


End Sub


Public Sub aceptarSolicitud(UserIndex As Integer, nombreLider As String)

Dim liderIndex As Integer
Dim resultadoAceptacion As Byte

' ¿Reglas generales para participar de un evento automatico?
If Not puedeParticipar(UserIndex) Then Exit Sub

liderIndex = NameIndex(nombreLider)

'Me fijo si el tipo esta online
If liderIndex = 0 Then
    EnviarPaquete Paquetes.mensajeinfo, nombreLider & " no te invitó a ningún evento o la invitación venció.", UserIndex, ToIndex
    Exit Sub
End If

' No se puede invitar a Game Masters
#If testeo = 0 Then
    If UserList(liderIndex).flags.Privilegios > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, nombreLider & " no te invitó a ningún evento o la invitación venció.", UserIndex, ToIndex
        Exit Sub
    End If
#End If

' ¿Se acepta a el mismo?
If liderIndex = UserIndex Then
     EnviarPaquete Paquetes.mensajeinfo, "No debés aceptarte a vos mismo una solicitud de ingreso a un evento.", UserIndex, ToIndex
     Exit Sub
End If

'El usuario no tiene ninguna solicitud...
If UserList(liderIndex).solicitudEvento Is Nothing Then
    EnviarPaquete Paquetes.mensajeinfo, nombreLider & " no te invitó a ningún evento o la invitación venció.", UserIndex, ToIndex
    Exit Sub
End If

'Tiene una solicitud. intento aceptarla
resultadoAceptacion = UserList(liderIndex).solicitudEvento.aceptarSolicitud(UserList(UserIndex).id)

If resultadoAceptacion = 0 Then
    'La solicitud no me sirve más
    Set UserList(liderIndex).solicitudEvento = Nothing
ElseIf resultadoAceptacion = 1 Then
    'La solicitud fue aceptada, ahora debes esperar.
    EnviarPaquete Paquetes.mensajeinfo, "Has aceptado participar en el evento que te invito " & UserList(liderIndex).Name & ". Ahora debes esperar que todos los demás invitados también acepten.", UserIndex, ToIndex
    EnviarPaquete Paquetes.mensajeinfo, UserList(UserIndex).Name & " aceptó tu invitación para participar del evento.", liderIndex, ToIndex
ElseIf resultadoAceptacion = 2 Then
    EnviarPaquete Paquetes.mensajeinfo, nombreLider & " no te invitó a ningún evento o la invitación venció.", UserIndex, ToIndex
ElseIf resultadoAceptacion = 3 Then
    EnviarPaquete Paquetes.mensajeinfo, "Tu ya has aceptado la invitación para ingresar al evento. Debes esperar que los otros invitados también acepten.", UserIndex, ToIndex
End If

End Sub

Public Function puedeParticipar(UserIndex As Integer) As Boolean
'Esta muerto?
With UserList(UserIndex)
    
    If .flags.Muerto = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes participar de un evento si estas muerto.", UserIndex, ToIndex
        puedeParticipar = False
        Exit Function
    End If
    
    'Me fijo si el tipo quiere entrar ya estando en un evento!
    If Not .evento Is Nothing Then
        If .evento.getEstadoEvento = Desarrollandose Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes participar de un evento si ya estas participando en otro.", UserIndex, ToIndex
            puedeParticipar = False
            Exit Function
        End If
    End If
    
    'Solo puede aceptar si esta en zona segura
    If MapInfo(.pos.map).Pk = True Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes participar de un evento si estas en zona insegura.", UserIndex, ToIndex
        puedeParticipar = False
        Exit Function
    End If
    
    'No puede participar si esta en la carce
    If .Counters.Pena > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes participar de un evento si estas en zona insegura.", UserIndex, ToIndex
        puedeParticipar = False
        Exit Function
    End If
End With

'Si llegamos hasta acá es porque esta todo ok
puedeParticipar = True
End Function
