Attribute VB_Name = "SV_Bloqueos"
Option Explicit

'Devuelve true si se puede ingresar al tile
' con el heading actual
'El heading con el que se entra es el contraio al bloqueo que se evalua.
Public Function sePuedeIngresarTile(Trigger As Long, Optional ByVal headingConElQueSeEntra As eHeading = eHeading.Ninguno) As Boolean

    Select Case headingConElQueSeEntra
        Case eHeading.EAST
            sePuedeIngresarTile = Not (Trigger And eTriggers.BloqueoOeste)
        Case eHeading.WEST
            sePuedeIngresarTile = Not (Trigger And eTriggers.BloqueoEste)
        Case eHeading.NORTH
            sePuedeIngresarTile = Not (Trigger And eTriggers.BloqueoSur)
        Case eHeading.SOUTH
            sePuedeIngresarTile = Not (Trigger And eTriggers.BloqueoNorte)
        Case eHeading.Ninguno
            sePuedeIngresarTile = Not ((Trigger And eTriggers.TodosBordesBloqueados) = eTriggers.TodosBordesBloqueados)
    End Select
    
End Function

Public Function isTileBloqueado(pos As MapBlock) As Boolean
    isTileBloqueado = ((pos.Trigger And eTriggers.TodosBordesBloqueados) = eTriggers.TodosBordesBloqueados)
End Function



Public Function obtenerBordeContrario(borde As eHeading)

Select Case borde

    Case eHeading.EAST
        obtenerBordeContrario = eHeading.WEST
    Case eHeading.WEST
        obtenerBordeContrario = eHeading.EAST
    Case eHeading.NORTH
        obtenerBordeContrario = eHeading.SOUTH
    Case eHeading.SOUTH
        obtenerBordeContrario = eHeading.NORTH
    Case eHeading.Ninguno
        obtenerBordeContrario = eHeading.Ninguno
End Select

End Function
