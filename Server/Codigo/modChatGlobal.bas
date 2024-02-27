Attribute VB_Name = "modChatGlobal"
Option Explicit

Public charlageneral As Boolean
Public UltimoMensajecharla As String

Public Sub activarChatGlobal()
    charlageneral = True
    EnviarPaquete Paquetes.MensajeServer, "Chat GLOBAL activado.", 0, ToAll
End Sub

Public Sub desactivarChatGlobal()
    charlageneral = False
    EnviarPaquete Paquetes.MensajeServer, "Chat GLOGBAL desactivado.", 0, ToAll
End Sub

Public Sub enviarMensaje(Usuario As User, mensaje As String)

    Dim loopUser As Integer
    
    mensaje = LCase$(Right(mensaje, Len(mensaje) - 1)) 'Paso el mensaje a minuscula y le saco el punto del principio
    
    If UltimoMensajecharla = mensaje Then Exit Sub Else UltimoMensajecharla = mensaje 'Antiflodeo
    
    mensaje = Usuario.Name & "> " & mensaje
    
    For loopUser = 1 To LastUser
        ' ¿Hay alguien ahi?
        If Not UserList(loopUser).ConnID = INVALID_SOCKET And UserList(loopUser).flags.UserLogged Then
            ' ¿Tiene el global activado?
            If UserList(loopUser).Stats.GlobAl = 2 Then
                EnviarPaquete Paquetes.mensajeGlobal, mensaje, loopUser, ToIndex
            End If
        End If
    Next
End Sub
