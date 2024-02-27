Attribute VB_Name = "modUtilitarios"
Option Explicit

'Prepara a un usuaro para poder combatir

Public Sub Preparando(UserIndex As Integer)

With UserList(UserIndex)

    If .flags.Muerto = 1 Then
        'Lo revivo
        Call RevivirUsuarioEnREeto(UserIndex)
    Else
        'Si esta paralizado lo reuevo
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
            EnviarPaquete Paquetes.NoParalizado2, "", UserIndex, ToIndex
        End If

        'Si esta oculto o invisible le quito la invisibilidad
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            EnviarPaquete Paquetes.Desocultar, ITS(.Char.charIndex), UserIndex, ToMap
            EnviarPaquete Paquetes.MensajeSimple, Chr$(23), UserIndex
        ElseIf .flags.Invisible = 1 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(23), UserIndex
            .Counters.Invisibilidad = 0
            .flags.Invisible = 0
            .flags.Oculto = 0
            EnviarPaquete Paquetes.Visible, ITS(.Char.charIndex), UserIndex, ToMap, .pos.map
        End If
        
        'Le doy toda la vida, mana y sta
        If .Stats.minHP <> .Stats.MaxHP Or .Stats.MinMAN <> .Stats.MaxMAN Or .Stats.MinSta <> .Stats.MaxSta Then
                UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.MaxHP
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
                Call SendUserStatsBox(val(UserIndex))
        End If
        
    End If
    
End With

End Sub
