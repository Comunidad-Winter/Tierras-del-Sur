Attribute VB_Name = "modMimetismo"
Option Explicit

Public Sub DoMimetizarConCriatura(ByRef personaje As User, ByRef Fuente As npc)

    ' Obtenemos la apareciencia que va a tener cuando este mimetizado
    If personaje.flags.TargetObj = 147 Or personaje.flags.TargetObj = 148 Then
        personaje.Mimetizado.Apareciencia.Body = 25
        personaje.Mimetizado.Apareciencia.Head = 0
        personaje.Mimetizado.Apareciencia.CascoAnim = NingunCasco
        personaje.Mimetizado.Apareciencia.ShieldAnim = NingunEscudo
        personaje.Mimetizado.Apareciencia.WeaponAnim = NingunArma
    ElseIf Fuente.Char.Body <> 0 Then
        personaje.Mimetizado.Apareciencia.Body = Fuente.Char.Body
        personaje.Mimetizado.Apareciencia.Head = Fuente.Char.Head
        
        If Fuente.Char.CascoAnim = 0 Then
            personaje.Mimetizado.Apareciencia.CascoAnim = NingunCasco
        Else
            personaje.Mimetizado.Apareciencia.CascoAnim = Fuente.Char.CascoAnim
        End If
        
        If Fuente.Char.ShieldAnim = 0 Then
            personaje.Mimetizado.Apareciencia.ShieldAnim = NingunEscudo
        Else
            personaje.Mimetizado.Apareciencia.ShieldAnim = Fuente.Char.ShieldAnim
        End If
        
        If Fuente.Char.WeaponAnim = 0 Then
            personaje.Mimetizado.Apareciencia.WeaponAnim = NingunArma
        Else
            personaje.Mimetizado.Apareciencia.WeaponAnim = Fuente.Char.WeaponAnim
        End If
    Else
        ' Nada
        Exit Sub
    End If
    
    personaje.Mimetizado.Privilegios = 0
    personaje.Mimetizado.alineacion = eAlineaciones.indefinido
    personaje.Mimetizado.nombre = ""
    
    ' Seteamos el flag
    personaje.flags.Mimetizado = 1
    personaje.Counters.Mimetismo = 0

    ' Actualizamos la apareciencia segun corresponda
    Call modPersonaje.DarAparienciaCorrespondiente(personaje)
    
    ' Lo refljeamos en los clientes
    Call modPersonaje_TCP.ActualizarEstetica(personaje)
End Sub


Public Sub DoMimetizarConPersonaje(ByRef personaje As User, ByRef Fuente As User)
    ' Marcamos al personaje como mimetizado
    personaje.flags.Mimetizado = 1
    personaje.Counters.Mimetismo = 0
    
    ' Guardamos la información con la cual se mimetiza
    personaje.Mimetizado.Apareciencia.Body = Fuente.Char.Body
    personaje.Mimetizado.Apareciencia.Head = Fuente.Char.Head
    personaje.Mimetizado.Apareciencia.CascoAnim = Fuente.Char.CascoAnim
    personaje.Mimetizado.Apareciencia.ShieldAnim = Fuente.Char.ShieldAnim
    personaje.Mimetizado.Apareciencia.WeaponAnim = Fuente.Char.WeaponAnim
    
    ' ¿El personaje con el cual me quiero mimetizar, esta mimetizado?.
    If Fuente.flags.Mimetizado = 1 Then
        ' Si lo esta, le pongo el nombre del personaje mimetizado
        personaje.Mimetizado.nombre = Fuente.Mimetizado.nombre
    Else
        If Fuente.GuildInfo.id = 0 Then
            personaje.Mimetizado.nombre = Fuente.Name
        Else
            personaje.Mimetizado.nombre = Fuente.Name & "<" & Fuente.GuildInfo.GuildName & ">"
        End If
    End If
    
    personaje.Mimetizado.alineacion = Fuente.faccion.alineacion
    
    If Fuente.flags.PertAlCons = 1 Or Fuente.flags.PertAlConsCaos = 1 Then
        personaje.Mimetizado.Privilegios = PRIV_USUARIOS_CONSEJO
    Else
        personaje.Mimetizado.Privilegios = Fuente.flags.Privilegios
    End If
   ' Actualizamos la apareciencia
   Call modPersonaje.DarAparienciaCorrespondiente(personaje)
   
   ' Lo reflejamos en los clientes
   Call modPersonaje_TCP.ActualizarEstetica(personaje)
End Sub

Public Sub finalizarEfecto(ByRef personaje As User)

    ' Lo marco
    personaje.Counters.Mimetismo = 0
    personaje.flags.Mimetizado = 0
    
    personaje.Mimetizado.nombre = ""
    personaje.Mimetizado.alineacion = eAlineaciones.indefinido
    personaje.Mimetizado.Privilegios = 0

    ' Generamos la apariencia original para el personaje
    Call modPersonaje.DarAparienciaCorrespondiente(personaje)
    
    ' Si el personaje NO esta navegando, actualizamos su estetica inmediatamente
    If personaje.flags.Navegando = 0 Then
        Call modPersonaje_TCP.ActualizarEstetica(personaje)
        EnviarPaquete Paquetes.mensajeinfo, "Recuperas tu apariencia normal.", personaje.UserIndex
    Else
        EnviarPaquete Paquetes.mensajeinfo, "Se te ha ido el efecto del mimetismo.", personaje.UserIndex
    End If
        
End Sub

