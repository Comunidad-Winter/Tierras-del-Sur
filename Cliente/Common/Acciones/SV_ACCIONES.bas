Attribute VB_Name = "Sv_Acciones"
Option Explicit

Public Enum tipo_accion
    COMPUESTA = 0
    EXIT_COMUN = 1
    EXIT_NORTE = 2
    EXIT_ESTE = 3
    EXIT_SUR = 4
    EXIT_OESTE = 5
End Enum


Public Function obtenerAccion(ByVal tipo As tipo_accion) As iAccion

    Select Case tipo

        Case tipo_accion.EXIT_COMUN
            Set obtenerAccion = New cAccionExit
        Case tipo_accion.EXIT_NORTE
            Set obtenerAccion = New cAccionExitNorte
        Case tipo_accion.EXIT_ESTE
            Set obtenerAccion = New cAccionExitEste
        Case tipo_accion.EXIT_OESTE
            Set obtenerAccion = New cAccionExitOeste
        Case tipo_accion.EXIT_SUR
            Set obtenerAccion = New cAccionExitSur
        Case tipo_accion.COMPUESTA
            Set obtenerAccion = New cAccionCompuesta
    End Select

End Function
