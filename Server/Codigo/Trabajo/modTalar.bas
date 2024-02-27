Attribute VB_Name = "modTalar"
Option Explicit

Public Const EsfuerzoTalarGeneral = 4
Public Const EsfuerzoTalarLeñador = 2

Public Function calcularRangoExtraccionTalar(personaje As User) As tRango

    If personaje.clase = eClases.Leñador Then
    
        Select Case personaje.Stats.UserSkills(eSkills.Talar)
            Case 0:
                calcularRangoExtraccionTalar.minimo = 0
                calcularRangoExtraccionTalar.maximo = 0
            Case 1 To 30:
                calcularRangoExtraccionTalar.minimo = 0
                calcularRangoExtraccionTalar.maximo = 1
            Case 31 To 60:
                calcularRangoExtraccionTalar.minimo = 0
                calcularRangoExtraccionTalar.maximo = 2
            Case 61 To 90:
                calcularRangoExtraccionTalar.minimo = 1
                calcularRangoExtraccionTalar.maximo = 2
            Case 91 To 99:
                calcularRangoExtraccionTalar.minimo = 1
                calcularRangoExtraccionTalar.maximo = 3
            Case 100:
                calcularRangoExtraccionTalar.minimo = 2
                calcularRangoExtraccionTalar.maximo = 4
        End Select
    
    Else
    
        Select Case personaje.Stats.UserSkills(eSkills.Talar)
            Case 0:
                calcularRangoExtraccionTalar.minimo = 0
                calcularRangoExtraccionTalar.maximo = 0
            Case 1 To 99:
                calcularRangoExtraccionTalar.minimo = 0
                calcularRangoExtraccionTalar.maximo = 1
            Case 100:
                calcularRangoExtraccionTalar.minimo = 1
                calcularRangoExtraccionTalar.maximo = 1
        End Select
    
    End If

    
End Function

Public Function calcularModificadorTalar(personaje As User) As Integer

    If personaje.clase = eClases.Leñador Then
        calcularModificadorTalar = 100
    Else
        Select Case personaje.Stats.UserSkills(eSkills.Talar)
            Case 0 To 30:
                calcularModificadorTalar = 400
            Case 31 To 60:
                calcularModificadorTalar = 300
            Case 61 To 90:
                calcularModificadorTalar = 200
            Case 91 To 100:
                calcularModificadorTalar = 100
        End Select
    End If
    
End Function

Public Sub DoTalar(personaje As User)
    Dim tieneEnergia As Boolean
    Dim MiObj As obj
    
    If personaje.clase = eClases.Leñador Then
        tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoTalarLeñador)
    Else
        tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoTalarGeneral)
    End If
    
    If Not tieneEnergia Then
        ' Le avisamos que esta cansado
        EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado cansado. Esperá un poco antes de seguir trabajando.", personaje.UserIndex, ToIndex
        ' Dejamos de trabajar
        Call modPersonaje.DejarDeTrabajar(personaje)
        Exit Sub
    End If
    
    ' Tratamos de subir Skill Siempre
    Call SubirSkill(personaje.UserIndex, eSkills.Talar)
    
    If personaje.Trabajo.modificador > 100 Then
        Dim res As Integer
        res = RandomNumber(1, personaje.Trabajo.modificador)
        
        ' ¿Tengo suerte de extaer?
        If res > 100 Then
            Exit Sub
        End If
    End If
    
    ' Creo el objeto que dependera de donde esta minando
    MiObj.ObjIndex = personaje.Trabajo.modo
    
    MiObj.Amount = RandomNumberInt(personaje.Trabajo.rangoGeneracion.minimo, personaje.Trabajo.rangoGeneracion.maximo)
    
    If MiObj.Amount = 0 Then
        Exit Sub
    End If
    
    ' ¿Tiene lugar para el objeto?
    If Not InvUsuario.tieneLugar(personaje, MiObj) Then
        ' Avisamos
        EnviarPaquete Paquetes.mensajeinfo, "No tienes más lugar para guardar leña.", personaje.UserIndex, ToIndex
        ' Dejamos de trabajar
        Call DejarDeTrabajar(personaje)
        ' Salimos
        Exit Sub
    End If
    
    ' Metemos el objeto en el inventario
    Call MeterItemEnInventario(personaje.UserIndex, MiObj)
    
    ' Enviamos el mensaje
    If Not personaje.flags.UltimoMensaje = 33 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(124), personaje.UserIndex
        personaje.flags.UltimoMensaje = 33
    End If
    
    ' Energia
    Call SendUserEsta(personaje.UserIndex)
    
    ' Efecto
    EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_TALAR), personaje.UserIndex, ToPCArea, personaje.pos.map

End Sub
