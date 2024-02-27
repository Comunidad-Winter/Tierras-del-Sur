Attribute VB_Name = "modClima"
Option Explicit

Public Enum eClimas
    climalluvia = 2
    ClimaNeblina = 1
    ClimaNiebla = 4
    ClimaTormenta_de_arena = 8
    ClimaNublado = 16
    ClimaNieve = 32
    ClimaRayos_de_luz = 64
End Enum


Public bRain        As Boolean 'está raineando?
Public bSnow        As Boolean 'está nevando?


Public Sub setClima(clima As Integer)


    Call Cambiar_estado_climatico(clima)

End Sub


