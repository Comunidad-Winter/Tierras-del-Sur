Attribute VB_Name = "GamePLay"
Option Explicit

Private Const TIEMPO_COMIENZO_MEDITAR_MS = 0

Public Sub Meditar(personaje As User)

    If personaje.Stats.MaxHP = 0 Then Exit Sub
    
    ' Cambiamos el estado
    personaje.flags.Meditando = Not personaje.flags.Meditando
    
    ' Si dejo de meditar le quito los efectos
    If personaje.flags.Meditando = False Then
        personaje.Char.FX = 0
        personaje.Char.loops = 0
        
        Call modPersonaje_TCP.ActualizarMeditacion(personaje)
        Exit Sub
    End If
    
    ' Si comienza le asigno la meditacion
    Select Case personaje.Stats.ELV
        Case 1 To 14
            personaje.Char.FX = Efectos_Constantes.FXMEDITARCHICO
        Case 15 To 29
            personaje.Char.FX = Efectos_Constantes.FXMEDITARMEDIANO
        Case 30 To 44
            personaje.Char.FX = Efectos_Constantes.FXMEDITARGRANDE
        Case 45 To 49
            personaje.Char.FX = Efectos_Constantes.FXMEDITARGIGANTE
        Case 50
            personaje.Char.FX = Efectos_Constantes.FXMEDITAR_4
    End Select
    
    Call modPersonaje_TCP.ActualizarMeditacion(personaje)
    
    ' Arranca con una penalización de dos segundos.
    personaje.Counters.Meditacion = TIEMPO_COMIENZO_MEDITAR_MS * -1
            
End Sub

