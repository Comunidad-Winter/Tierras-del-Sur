Attribute VB_Name = "Engine_ComLib"
Option Explicit

Public LucesPermitidas As Byte 'Luces permitidas por el hardware

Public Enum TiposLuces
    Cuadradas = 16
    ConBrillo = 1
    Redondas = 2
    todas = 255
End Enum

Public Sub Instanciar_Engine()
    LucesPermitidas = TiposLuces.todas
End Sub

