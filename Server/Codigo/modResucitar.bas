Attribute VB_Name = "modResucitar"
Option Explicit

Private resucitacionesPendientes As Collection

Private Const TIEMPO_RESUCITACION_MILISEGUNDOS = 1 * 1000

    
Public Sub iniciar()
    Set resucitacionesPendientes = New Collection
End Sub

Public Sub cancelarResucitacion(ByRef resucitacionPendiente As resucitacionPendiente)
    Dim loopC As Integer
    For loopC = 1 To resucitacionesPendientes.Count
        If resucitacionesPendientes(loopC) Is resucitacionPendiente Then
            resucitacionesPendientes.Remove loopC
            Exit For
        End If
    Next loopC
    
    Dim userIndexResucitador As Integer
    Dim userIndexResucitado As Integer
    
    userIndexResucitador = resucitacionPendiente.resucitado
    userIndexResucitado = resucitacionPendiente.resucitador
    
    Set UserList(resucitacionPendiente.resucitado).resucitacionPendiente = Nothing
    Set UserList(resucitacionPendiente.resucitador).resucitacionPendiente = Nothing
        
    EnviarPaquete Paquetes.mensajeinfo, UserList(userIndexResucitado).Name & " perdió la concentración. Se cancela el proceso de resucitación.", userIndexResucitado, ToIndex
    EnviarPaquete Paquetes.mensajeinfo, "Te desconcentras. Se cancela el proceso de resucitación.", userIndexResucitador, ToIndex
End Sub

Public Sub agregarResucitacion(personajeResucitador As User, personajeResucitado As User)

    Dim resucitacionPendiente As resucitacionPendiente
    
    Set resucitacionPendiente = New resucitacionPendiente
    
    resucitacionPendiente.fecha = GetTickCount + TIEMPO_RESUCITACION_MILISEGUNDOS
    resucitacionPendiente.resucitado = personajeResucitado.UserIndex
    resucitacionPendiente.resucitador = personajeResucitador.UserIndex
    
    
    Set UserList(personajeResucitado.UserIndex).resucitacionPendiente = resucitacionPendiente
    Set UserList(personajeResucitador.UserIndex).resucitacionPendiente = resucitacionPendiente
    
    ' Agregamos a la lista
   Call resucitacionesPendientes.Add(resucitacionPendiente)
End Sub

Public Sub procesarResucitacionesPendientes(ahora As Long)
    Dim loopC As Integer
    Dim resucitacionPendiente As resucitacionPendiente
    
    Do While resucitacionesPendientes.Count > 0
        
        Set resucitacionPendiente = resucitacionesPendientes(1)
    
        If resucitacionPendiente.fecha > ahora Then
            Exit Do
        End If
        
        ' Removemos la estructura
        resucitacionesPendientes.Remove 1
    
        Set UserList(resucitacionPendiente.resucitado).resucitacionPendiente = Nothing
        Set UserList(resucitacionPendiente.resucitador).resucitacionPendiente = Nothing
            
        ' Revivimos
        Call modPersonaje.RevivirUsuario(UserList(resucitacionPendiente.resucitado), UserList(resucitacionPendiente.resucitado).Stats.MaxHP * 0.25)
        
        ' Penalizamos a la persona que revivo
        Call penalizarResucitador(UserList(resucitacionPendiente.resucitador))
                
        '¿El tipo esta en un evento?. Le aviso al evento que le dieron resu1
        If Not UserList(resucitacionPendiente.resucitado).evento Is Nothing Then
            Call UserList(resucitacionPendiente.resucitado).evento.usuarioRevive(UserList(resucitacionPendiente.resucitado).UserIndex, UserList(resucitacionPendiente.resucitador).UserIndex)
        End If
    Loop
  
End Sub

Private Sub penalizarResucitador(ByRef personaje As User)

    personaje.Stats.minHP = personaje.Stats.minHP - personaje.Stats.minHP * 0.4
    
    EnviarPaquete Paquetes.MensajeFight, "Resucitar a otro reduce tu salud.", personaje.UserIndex

    If personaje.Stats.minHP <= 0 Then
        personaje.Stats.minHP = 0
        Call UserDie(personaje.UserIndex, False)
    Else
        Call SendUserVida(personaje.UserIndex)
    End If

End Sub



