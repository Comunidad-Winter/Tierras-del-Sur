Attribute VB_Name = "modCombate"
Option Explicit


Public Sub usuarioPegar(personaje As User, timeStamp As Single)

    ' ¿Muerto?
    If personaje.flags.Muerto = 1 Then Exit Sub

    ' Los consejeros no pueden pegar
    If personaje.flags.Privilegios = PRIV_CONSEJERO Then Exit Sub
 
    'Si esta el contador no puede realizar acciones
    If personaje.Counters.combateRegresiva > 0 Then Exit Sub

    ' Esta resucitando a alguien?
    If Not personaje.resucitacionPendiente Is Nothing Then
        Call modResucitar.cancelarResucitacion(personaje.resucitacionPendiente)
    End If
    
    ' ¿Tiene un arma equipada?
    If personaje.Invent.WeaponEqpObjIndex > 0 Then
        ' ¿Es un arma que lanza proyectiles?
        If ObjData(personaje.Invent.WeaponEqpObjIndex).proyectil = 1 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(219), personaje.UserIndex
            Exit Sub
        End If
    End If
        
    ' --- ANTICHEAT ---
    personaje.controlCheat.VecesAtack = personaje.controlCheat.VecesAtack + 1
    
    Call anticheat.chequeoIntervaloCliente(personaje, personaje.Counters.ultimoTickPegar, personaje.intervalos.Golpe, timeStamp, "pegar")
    ' -----------------

    Call UsuarioAtaca(personaje.UserIndex)

    ' ¿Se mantiene oculto pese a atacar?
    If personaje.clase = eClases.Cazador And personaje.flags.Oculto > 0 And personaje.Stats.UserSkills(Ocultarse) > 90 Then
        If personaje.Invent.ArmourEqpObjIndex = ARMADURA_DE_CAZADOR Then
            Exit Sub
        End If
    End If
    
    ' No tiene la propiedad de mantenerse oculto
    If personaje.flags.Oculto > 0 Then
        personaje.flags.Oculto = 0
        personaje.flags.Invisible = 0
        EnviarPaquete Paquetes.Desocultar, ITS(personaje.Char.charIndex), personaje.UserIndex, ToMap
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(288 - 255), personaje.UserIndex, ToIndex
    End If

End Sub
