Attribute VB_Name = "modPersonaje_TCP"
Option Explicit

Public Sub ActualizarMeditacion(personaje As User)
    EnviarPaquete Paquetes.Meditando, "", personaje.UserIndex, ToIndex
    
    If personaje.Char.FX = 0 Then
        EnviarPaquete Paquetes.HechizoFX, ITS(personaje.Char.charIndex) & ByteToString(0) & ITS(0), personaje.UserIndex, ToMap, personaje.pos.map
    Else
        EnviarPaquete Paquetes.AuraFx, ITS(personaje.Char.charIndex) & Codify(personaje.Char.FX), personaje.UserIndex, ToMap
    End If
End Sub

Public Sub actualizarExperiencia(ByRef personaje As User)
   EnviarPaquete Paquetes.EnviarEXP, WriteString(FormatNumber(personaje.Stats.Exp, 0, vbTrue, vbFalse, vbFalse)), personaje.UserIndex, ToIndex
End Sub
' Crea la Representacion del usuario en un mapa para determinados (sndRoute, destinoIndex) usuario/s.
Public Sub MakeUserChar(ByRef personaje As User, destinoIndex As Integer, sndRoute As Byte)

Dim charIndex As Integer
Dim nombre As String
Dim PrivL As Byte
Dim bcr As Long
Dim tieneClan As Byte

' ¿Posicion valida?
If Not SV_PosicionesValidas.esPosicionJugable(personaje.pos.x, personaje.pos.y) Then Exit Sub

' ¿El personaje es totalmente nuevo?
If personaje.Char.charIndex = 0 Then
    ' Obtenems un charindex libre
    charIndex = NextOpenCharIndex
    
    ' Establecemos las relaciones
    personaje.Char.charIndex = charIndex
    CharList(charIndex) = personaje.UserIndex
End If
      
' Lo ponemos en la posicion del amap
MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).UserIndex = personaje.UserIndex
           
If personaje.flags.Mimetizado = 1 Then

    bcr = personaje.Mimetizado.alineacion
    nombre = personaje.Mimetizado.nombre
    PrivL = personaje.Mimetizado.Privilegios
Else

    bcr = personaje.faccion.alineacion
    
    If personaje.GuildInfo.id > 0 Then
        nombre = personaje.Name & "<" & personaje.GuildInfo.GuildName & ">"
        tieneClan = 1
    Else
        nombre = personaje.Name
        tieneClan = 0
    End If
      
    If personaje.flags.PertAlCons = 1 Or personaje.flags.PertAlConsCaos = 1 Then
        PrivL = PRIV_USUARIOS_CONSEJO
    End If
    
    If personaje.flags.Privilegios > 0 Then
        PrivL = personaje.flags.Privilegios
    End If
End If

' Creamos el char
EnviarPaquete CrearChar, ITS(personaje.Char.charIndex) _
                        & ByteToString(personaje.Char.FX) _
                        & ITS(personaje.Char.Body) _
                        & ITS(personaje.Char.Head) _
                        & ByteToString(personaje.Char.heading) _
                        & ITS(personaje.pos.x) & ITS(personaje.pos.y) _
                        & ByteToString(personaje.Char.WeaponAnim) _
                        & ByteToString(personaje.Char.ShieldAnim) _
                        & ByteToString(personaje.Char.CascoAnim) _
                        & ByteToString(bcr) _
                        & ByteToString(PrivL) _
                        & ByteToString(tieneClan) _
                        & nombre, destinoIndex, sndRoute, personaje.pos.map

End Sub

Public Sub enviarPosicion(ByRef personaje As User)
    EnviarPaquete Paquetes.EnviarPos, ITS(personaje.pos.x) & ITS(personaje.pos.y), personaje.UserIndex, ToIndex
End Sub

Public Sub enviarParalizado(ByRef personaje As User)
    EnviarPaquete Paquetes.Paralizado2, ITS(personaje.pos.x) & ITS(personaje.pos.y), personaje.UserIndex
End Sub

Public Sub actualizarNick(ByRef personaje As User)

Dim bcr As String
Dim nombre As String
Dim PrivL As Byte

PrivL = 0

If personaje.flags.Mimetizado = 1 Then
    nombre = personaje.Mimetizado.nombre
    PrivL = personaje.Mimetizado.Privilegios
    bcr = personaje.Mimetizado.alineacion
Else
    ' Nombre
    If personaje.GuildInfo.id > 0 Then
        nombre = personaje.Name & "<" & personaje.GuildInfo.GuildName & ">"
    Else
        nombre = personaje.Name
    End If
    
    bcr = personaje.faccion.alineacion
    
    ' Privilegios
    If personaje.flags.PertAlCons = 1 Or personaje.flags.PertAlConsCaos = 1 Then
        PrivL = PRIV_USUARIOS_CONSEJO
    End If
    
    If personaje.flags.Privilegios > 0 Then
        PrivL = personaje.flags.Privilegios
    End If
End If

EnviarPaquete Paquetes.ActualizaNick, ITS(personaje.Char.charIndex) & ByteToString(bcr) & ByteToString(PrivL) & nombre, personaje.UserIndex, ToMap

End Sub
' Envía la información a los clientes para que actualicen la estetica del personaje
Public Sub ActualizarEstetica(ByRef personaje As User)



' Personaje
EnviarPaquete Paquetes.pChangeUserChar, ITS(personaje.Char.charIndex) _
                                        & ITS(personaje.Char.Body) & ITS(personaje.Char.Head) _
                                        & ByteToString(personaje.Char.heading) _
                                        & ByteToString(personaje.Char.WeaponAnim) _
                                        & ByteToString(personaje.Char.ShieldAnim) _
                                        & ByteToString(personaje.Char.FX) _
                                        & ITS(personaje.Char.loops) _
                                        & ByteToString(personaje.Char.CascoAnim), personaje.UserIndex, ToArea


' Nombre
Call actualizarNick(personaje)

End Sub

