Attribute VB_Name = "modOcultar"
Option Explicit


Private Function isDebePermanencerOculto(ByRef personaje As User, tiempoActual As Long)

isDebePermanencerOculto = True

' Hay un tiempo minimo de ocultacion
If personaje.Counters.TimerOculto < 3000 Then
    Exit Function
End If

Dim maximoTiempo As Long
 
If personaje.clase = eClases.Guerrero Then
    maximoTiempo = (personaje.Stats.ELV \ 4) * 1000
    
    If maximoTiempo <= personaje.Counters.TimerOculto Then
        isDebePermanencerOculto = False
        Exit Function
    End If
ElseIf personaje.clase = eClases.Cazador Then
    If modPersonaje.tieneArmaduraCazador(personaje) Then
        Exit Function
    Else
        maximoTiempo = (personaje.Stats.ELV \ 3) * 1000
        
        If maximoTiempo <= personaje.Counters.TimerOculto Then
            isDebePermanencerOculto = False
            Exit Function
        End If
    End If
Else
    If personaje.Counters.TimerOculto >= 5000 Then
        isDebePermanencerOculto = False
    Else
        Dim azar As Integer
        
        azar = RandomNumberInt(1, 101)
        
        If azar > personaje.Stats.UserSkills(Ocultarse) Then
            isDebePermanencerOculto = False
        End If
    End If
End If

End Function
Public Sub DoPermanecerOculto(ByRef personaje As User, tiempo As Long)

' Tiempo oculto
personaje.Counters.TimerOculto = personaje.Counters.TimerOculto + tiempo

If isDebePermanencerOculto(personaje, tiempo) = False Then
    ' Pierde
    Call quitarOcultamiento(personaje)
    ' Mensaje
    EnviarPaquete Paquetes.MensajeSimple, Chr$(23), personaje.UserIndex
End If

End Sub

Public Sub quitarOcultamiento(personaje As User)
    personaje.eventoOcultar.Posicion.x = 0
    personaje.eventoOcultar.Posicion.y = 0
                    
    personaje.flags.Oculto = 0
    personaje.Counters.TimerOculto = 0
    
    EnviarPaquete Paquetes.Desocultar, ITS(personaje.Char.charIndex), personaje.UserIndex, ToMap
End Sub


Public Sub DoOcultarse(ByRef personaje As User)

Dim ahora As Long

' Si esta Invisible se le va.`
ahora = GetTickCount()

' Intervalo de ocultarse es de un segundo.
If personaje.Counters.ultimoIntentoOcultar + 1000 > ahora Then
    Exit Sub
End If

personaje.Counters.ultimoIntentoOcultar = ahora

' Cuando tenes exito y te ocultas, cuando te desocultas tenes que esperar 5 segundos para volver a tocar la O.
If personaje.eventoOcultar.fecha + 5000 > ahora Then
    Exit Sub
End If

If personaje.flags.Muerto = 1 Then Exit Sub
If personaje.flags.Navegando = 1 Then Exit Sub
If personaje.flags.Oculto = 1 Then Exit Sub
If MapInfo(personaje.pos.map).AntiHechizosPts = 1 Then Exit Sub
        

'En un evento no vale el ocultar
If Not personaje.evento Is Nothing Then
    If personaje.evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then Exit Sub
End If

Dim skillsOcultarse As Integer
Dim Suerte As Integer
Dim res As Integer

skillsOcultarse = personaje.Stats.UserSkills(Ocultarse)

' Obtenemos la suerte para esto
If skillsOcultarse >= 1 And skillsOcultarse <= 20 Then
    Suerte = 20
ElseIf skillsOcultarse <= 50 Then
    Suerte = 50
ElseIf skillsOcultarse <= 75 Then
    Suerte = 75
ElseIf skillsOcultarse <= 99 Then
    Suerte = 85
Else
    Suerte = 100
End If

If Not (personaje.clase = eClases.Cazador Or personaje.clase = eClases.Guerrero) Then
    Suerte = Suerte * 0.5
End If

' Tiramos el dado
res = RandomNumberInt(1, 100)

If res <= Suerte Then
    ' Activamos
    If personaje.flags.Invisible = 1 Then
       Call quitarInvisibilidad(personaje)
    End If

   personaje.flags.Oculto = 1

   
   ' Guardamos informacion del Evento
   personaje.eventoOcultar.fecha = GetTickCount
   personaje.eventoOcultar.Posicion.x = personaje.pos.x
   personaje.eventoOcultar.Posicion.y = personaje.pos.y
   personaje.Counters.TimerOculto = 0
   
   ' Mensajes
   EnviarPaquete Paquetes.ocultar, ITS(personaje.Char.charIndex), personaje.UserIndex, ToMap
   EnviarPaquete Paquetes.MensajeSimple, Chr$(93), personaje.UserIndex
   
   ' Subir Skill
   Call SubirSkill(personaje.UserIndex, Ocultarse)
Else
    ' Fallo
    If Not personaje.flags.UltimoMensaje = 4 Then
      EnviarPaquete Paquetes.MensajeSimple, Chr$(1), personaje.UserIndex
      personaje.flags.UltimoMensaje = 4
    End If
End If

End Sub

