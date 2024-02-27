Attribute VB_Name = "modCasteoTorneos"
Option Explicit

Public Sub setRing(objeto As iModoTorneo, ring As tRing)
    Select Case TypeName(objeto)
        Case "iModoTorneo_DeathMach"
            Dim deathmatch As iModoTorneo_DeathMach
            Set deathmatch = objeto
            Call deathmatch.iModoTorneo_setRing(ring)
            Exit Sub
        Case "iModoTorneo_PlayOff"
            Dim PlayOff As iModoTorneo_PlayOff
            Set PlayOff = objeto
            Call PlayOff.iModoTorneo_setRing(ring)
            Exit Sub
        Case "iModoTorneo_Liga"
            Dim Liga As iModoTorneo_Liga
            Set Liga = objeto
            Call Liga.iModoTorneo_setRing(ring)
            Exit Sub
    End Select
End Sub

Public Sub setRings(objeto As iModoTorneo, rings() As tRing)
    Select Case TypeName(objeto)
        Case "iModoTorneo_DeathMach"
            Dim deathmatch As iModoTorneo_DeathMach
            Set deathmatch = objeto
            Call deathmatch.iModoTorneo_setRings(rings)
            Exit Sub
        Case "iModoTorneo_PlayOff"
            Dim PlayOff As iModoTorneo_PlayOff
            Set PlayOff = objeto
            Call PlayOff.iModoTorneo_setRings(rings)
            Exit Sub
        Case "iModoTorneo_Liga"
            Dim Liga As iModoTorneo_Liga
            Set Liga = objeto
            Call Liga.iModoTorneo_setRings(rings)
            Exit Sub
    End Select
End Sub

Public Sub setTablaEquipos(objeto As iModoTorneo, tablaEquipos() As tEquipoTablaTorneo)
    Select Case TypeName(objeto)
        Case "iModoTorneo_DeathMach"
            Dim deathmatch As iModoTorneo_DeathMach
            Set deathmatch = objeto
            Call deathmatch.iModoTorneo_setTablaEquipos(tablaEquipos)
            Exit Sub
        Case "iModoTorneo_PlayOff"
            Dim PlayOff As iModoTorneo_PlayOff
            Set PlayOff = objeto
            Call PlayOff.iModoTorneo_setTablaEquipos(tablaEquipos)
            Exit Sub
        Case "iModoTorneo_Liga"
            Dim Liga As iModoTorneo_Liga
            Set Liga = objeto
            Call Liga.iModoTorneo_setTablaEquipos(tablaEquipos)
            Exit Sub
    End Select
End Sub

Public Function obtenerTabla(objeto As iModoTorneo) As tEquipoTablaTorneo()
    Select Case TypeName(objeto)
        Case "iModoTorneo_DeathMach"
            Dim deathmatch As iModoTorneo_DeathMach
            Set deathmatch = objeto
            obtenerTabla = deathmatch.iModoTorneo_obtenerTabla
            Exit Function
        Case "iModoTorneo_PlayOff"
            Dim PlayOff As iModoTorneo_PlayOff
            Set PlayOff = objeto
            obtenerTabla = PlayOff.iModoTorneo_obtenerTabla
            Exit Function
        Case "iModoTorneo_Liga"
            Dim Liga As iModoTorneo_Liga
            Set Liga = objeto
            obtenerTabla = Liga.iModoTorneo_obtenerTabla
            Exit Function
    End Select
End Function

Public Function obtenerEquipo(objeto As iModoTorneo, idEquipo As Byte) As tEquipoTablaTorneo
    Select Case TypeName(objeto)
        Case "iModoTorneo_DeathMach"
            Dim deathmatch As iModoTorneo_DeathMach
            Set deathmatch = objeto
            obtenerEquipo = deathmatch.iModoTorneo_obtenerEquipo(idEquipo)
            Exit Function
        Case "iModoTorneo_PlayOff"
            Dim PlayOff As iModoTorneo_PlayOff
            Set PlayOff = objeto
            obtenerEquipo = PlayOff.iModoTorneo_obtenerEquipo(idEquipo)
            Exit Function
        Case "iModoTorneo_Liga"
            Dim Liga As iModoTorneo_Liga
            Set Liga = objeto
            obtenerEquipo = Liga.iModoTorneo_obtenerEquipo(idEquipo)
            Exit Function
    End Select
End Function

