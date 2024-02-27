Attribute VB_Name = "ME_Climas"
Option Explicit

Private Type tClimaDisponible
    tipo As Tipos_Clima
    nombre As String
End Type

Private climasDisponibles() As tClimaDisponible
Private Const CANTIDAD_CLIMAS = 7

Public Sub cargarClimasDisponibles()

ReDim climasDisponibles(1 To 7) As tClimaDisponible

climasDisponibles(1).nombre = "Normal"
climasDisponibles(1).tipo = Tipos_Clima.ClimaNinguno

climasDisponibles(2).nombre = "Lluvioso"
climasDisponibles(2).tipo = Tipos_Clima.ClimaLluvia

climasDisponibles(3).nombre = "Neblina"
climasDisponibles(3).tipo = Tipos_Clima.ClimaNeblina

climasDisponibles(4).nombre = "Tormenta de arena"
climasDisponibles(4).tipo = Tipos_Clima.ClimaTormenta_de_arena

climasDisponibles(5).nombre = "Nublado"
climasDisponibles(5).tipo = Tipos_Clima.ClimaNublado

climasDisponibles(6).nombre = "Nevando"
climasDisponibles(6).tipo = Tipos_Clima.ClimaNieve

climasDisponibles(7).nombre = "Soleado"
climasDisponibles(7).tipo = Tipos_Clima.ClimaRayos_de_luz

End Sub

Public Sub cargarClimasDisponiblesEnCombo(combo As ComboBox)

Dim loopClima As Byte

combo.Clear

For loopClima = 1 To CANTIDAD_CLIMAS
    Call combo.AddItem(climasDisponibles(loopClima).nombre)
Next loopClima

End Sub

Public Function obtenerTipoClima(nombre As String) As Tipos_Clima

Dim loopClima As Byte

obtenerTipoClima = Tipos_Clima.ClimaNinguno

For loopClima = 1 To CANTIDAD_CLIMAS
    If climasDisponibles(loopClima).nombre = nombre Then
       obtenerTipoClima = climasDisponibles(loopClima).tipo
       Exit Function
    End If
Next loopClima

End Function
