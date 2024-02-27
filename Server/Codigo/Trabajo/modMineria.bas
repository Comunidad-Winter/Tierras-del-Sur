Attribute VB_Name = "modMineria"
Option Explicit

Private Const EsfuerzoExcavarMinero = 2
Private Const EsfuerzoExcavarGeneral = 5


Public Function calcularRangoExtraccionMineria(personaje As User) As tRango

    If personaje.clase = eClases.Minero Then
    
        Select Case personaje.Stats.UserSkills(eSkills.Mineria)
            Case 0:
                calcularRangoExtraccionMineria.minimo = 0
                calcularRangoExtraccionMineria.maximo = 0
            Case 1 To 30:
                calcularRangoExtraccionMineria.minimo = 0
                calcularRangoExtraccionMineria.maximo = 1
            Case 31 To 60:
                calcularRangoExtraccionMineria.minimo = 0
                calcularRangoExtraccionMineria.maximo = 2
            Case 61 To 90:
                calcularRangoExtraccionMineria.minimo = 1
                calcularRangoExtraccionMineria.maximo = 2
            Case 91 To 99:
                calcularRangoExtraccionMineria.minimo = 1
                calcularRangoExtraccionMineria.maximo = 3
            Case 100:
                calcularRangoExtraccionMineria.minimo = 2
                calcularRangoExtraccionMineria.maximo = 4
        End Select
    
    Else
    
        Select Case personaje.Stats.UserSkills(eSkills.Mineria)
            Case 0:
                calcularRangoExtraccionMineria.minimo = 0
                calcularRangoExtraccionMineria.maximo = 0
            Case 1 To 99:
                calcularRangoExtraccionMineria.minimo = 0
                calcularRangoExtraccionMineria.maximo = 1
            Case 100:
                calcularRangoExtraccionMineria.minimo = 1
                calcularRangoExtraccionMineria.maximo = 1
        End Select
    
    End If

    
End Function

Public Function calcularModificadorMineria(personaje As User) As Integer

    If personaje.clase = eClases.Minero Then
        calcularModificadorMineria = 100
    Else
        Select Case personaje.Stats.UserSkills(eSkills.Mineria)
            Case 0 To 30:
                calcularModificadorMineria = 400
            Case 31 To 60:
                calcularModificadorMineria = 300
            Case 61 To 90:
                calcularModificadorMineria = 200
            Case 91 To 100:
                calcularModificadorMineria = 100
        End Select
    End If
    
End Function

    
    
    
Public Sub DoMineria(personaje As User)

Dim tieneEnergia As Boolean
Dim MiObj As obj
    
' Sacamos energia que consume esta accion
If personaje.clase = eClases.Minero Then
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoExcavarMinero)
Else
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoExcavarGeneral)
End If

' ¿Le pudimos sacar la energia?
If Not tieneEnergia Then
    ' Le avisamos que esta cansado
    EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado cansado. Esperá un poco antes de seguir trabajando.", personaje.UserIndex, ToIndex
    ' Dejamos de trabajar
    Call modPersonaje.DejarDeTrabajar(personaje)
    Exit Sub
End If

' Subimos Skills
Call SubirSkill(personaje.UserIndex, eSkills.Mineria)

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
    EnviarPaquete Paquetes.mensajeinfo, "No tienes más lugar para guardar minerales.", personaje.UserIndex, ToIndex
    ' Dejamos de trabajar
    Call DejarDeTrabajar(personaje)
    ' Salimos
    Exit Sub
End If
    
' Metemos el objeto
Call MeterItemEnInventario(personaje.UserIndex, MiObj)

' Avisamos
If Not personaje.flags.UltimoMensaje = 91 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(126), personaje.UserIndex
    personaje.flags.UltimoMensaje = 91
End If

' Energia
Call SendUserEsta(personaje.UserIndex)

' Efecto
EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_MINERO), personaje.UserIndex, ToPCArea
         
End Sub
