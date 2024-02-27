Attribute VB_Name = "modPesca"
Option Explicit

Public Sub DoPescar_Cana(personaje As User)

Dim tieneEnergia As Boolean
Dim numeroPez As Byte
Dim Suerte As Integer
Dim MiObj As obj

'Energia
If personaje.clase = eClases.Pescador Then
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoPescarPescador)
Else
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoPescarGeneral)
End If

If Not tieneEnergia Then
    ' Le avisamos que esta cansado
    EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado cansado. Esperá un poco antes de seguir trabajando.", personaje.UserIndex, ToIndex
    ' Dejamos de trabajar
    Call modPersonaje.DejarDeTrabajar(personaje)
    Exit Sub
End If

' Siempre sube Skill
Call SubirSkill(personaje.UserIndex, eSkills.Pesca)

' Tiramos la suerte
Suerte = RandomNumber(1, personaje.Trabajo.modificador)

If Suerte > 58 Then Exit Sub ' No ha tenido suerte

' Cuantos peces saca?
If personaje.clase = eClases.Pescador Then
    If Suerte < 3 And personaje.flags.Navegando = 1 And personaje.Invent.BarcoObjIndex = 475 Then
        numeroPez = 4
    ElseIf Suerte < 13 And personaje.flags.Navegando = 1 Then
        numeroPez = 3
    ElseIf Suerte < 19 Then
        numeroPez = 2
    Else
        numeroPez = 1
    End If
Else
    numeroPez = 1
End If

' Creamos los peces
Do While numeroPez > 0
   
    MiObj.Amount = 1 ' Siempre saca 1
    
    ' ¿Qué pez le toca?
    If numeroPez = 1 Then
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO1
    ElseIf numeroPez = 2 Then
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO2
    ElseIf numeroPez = 3 Then
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO3
    Else
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO4
    End If
    
    ' Lo agregamos
    If Not InvUsuario.tieneLugar(personaje, MiObj) Then
        ' Avisamos
        EnviarPaquete Paquetes.mensajeinfo, "No tienes más lugar para guardar más pesces.", personaje.UserIndex, ToIndex
        ' Dejamos de trabajar
        Call DejarDeTrabajar(personaje)
        ' Salimos
        Exit Sub
    Else
        ' Agregamos
        Call InvUsuario.MeterItemEnInventario(personaje.UserIndex, MiObj)
    End If

    ' Siguiente pez
    numeroPez = numeroPez - 1
Loop

' Efectos y Mensaje
EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_PESCAR), personaje.UserIndex, ToPCArea

' Energia
Call SendUserEsta(personaje.UserIndex)

If Not personaje.flags.UltimoMensaje = 6 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(117), personaje.UserIndex
    personaje.flags.UltimoMensaje = 6
End If

End Sub

Public Sub DoPescar(personaje As User)

    If personaje.Trabajo.modo = OBJTYPE_CAÑA Then
        Call DoPescar_Cana(personaje)
    Else
        Call DoPescar_Red(personaje)
    End If
    
End Sub
Public Sub DoPescar_Red(personaje As User)
Dim tieneEnergia As Boolean
Dim MiObj As obj
Dim numeroPez As Byte
Dim Suerte As Single

'Energia
If personaje.clase = eClases.Pescador Then
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoPescarPescador)
Else
    tieneEnergia = modPersonaje.QuitarEnergia(personaje, EsfuerzoPescarGeneral)
End If

If Not tieneEnergia Then
    ' Le avisamos que esta cansado
    EnviarPaquete Paquetes.mensajeinfo, "Estás demasiado cansado. Esperá un poco antes de seguir trabajando.", personaje.UserIndex, ToIndex
    ' Dejamos de trabajar
    Call modPersonaje.DejarDeTrabajar(personaje)
    Exit Sub
End If

' Siempre sube Skill
Call SubirSkill(personaje.UserIndex, eSkills.Pesca)

Suerte = RandomNumber(1, 100)
    
If Suerte > 58.33 Then Exit Sub ' Nada!

' ¿Cuantos peces saca?
If Suerte < 2 Then
    numeroPez = 5
ElseIf Suerte < 3.22 Then
    numeroPez = 4
ElseIf Suerte < 13.31 Then
    numeroPez = 3
ElseIf Suerte < 19.44 Then
    numeroPez = 2
Else
    numeroPez = 1
End If

' Creamos los peces
Do While numeroPez > 0
   
    MiObj.Amount = 1 ' Siempre saca 1
    
    ' ¿Qué pez le toca?
    If numeroPez = 1 Then
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO1
    ElseIf numeroPez = 2 Then
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO2
    ElseIf numeroPez = 3 Then
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO3
    ElseIf numeroPez = 4 Then
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO4
    Else
        MiObj.ObjIndex = PECES_POSIBLES.PESCADO5
    End If
    
    ' Lo agregamos
    If Not InvUsuario.tieneLugar(personaje, MiObj) Then
        ' Avisamos
        EnviarPaquete Paquetes.mensajeinfo, "No tienes más lugar para guardar más pesces.", personaje.UserIndex, ToIndex
        ' Dejamos de trabajar
        Call DejarDeTrabajar(personaje)
        ' Salimos
        Exit Sub
    Else
        ' Agregamos
        Call InvUsuario.MeterItemEnInventario(personaje.UserIndex, MiObj)
    End If

    ' Siguiente pez
    numeroPez = numeroPez - 1
Loop

' Sonido de Pesca
EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_PESCAR), personaje.UserIndex, ToPCArea

' Energia
Call SendUserEsta(personaje.UserIndex)

' Mensaje
If Not personaje.flags.UltimoMensaje = 118 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(118), personaje.UserIndex
    personaje.flags.UltimoMensaje = 118
End If
 
End Sub
