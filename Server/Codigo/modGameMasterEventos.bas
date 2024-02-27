Attribute VB_Name = "modGameMasterEventos"
Option Explicit

Public Const MAPA_DESCANSO_GMS As Integer = 337

Public Function esMapaDeEvento(mapa As Integer)
    If mapa = MAPA_DESCANSO_GMS Or (mapa >= 370 And mapa <= 385) Then
        esMapaDeEvento = True
    Else
        esMapaDeEvento = False
    End If
End Function

