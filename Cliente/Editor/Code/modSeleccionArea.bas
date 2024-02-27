Attribute VB_Name = "modSeleccionArea"
Option Explicit


Public Type tAreaSeleccionada
        arriba As Integer
        abajo As Integer
        izquierda As Integer
        derecha As Integer
        invertidoHorizontal As Boolean
        invertidoVertical As Boolean
End Type

    



Public Sub puntoArea(ByRef area As tAreaSeleccionada, X As Integer, Y As Integer)
    area.abajo = Y
    area.arriba = Y
    area.derecha = X
    area.izquierda = X
    area.invertidoHorizontal = False
    area.invertidoVertical = False
End Sub

Public Sub actualizarArea(ByRef area As tAreaSeleccionada, nuevoX As Integer, nuevoY As Integer)

    If nuevoY = area.arriba Then
        If area.invertidoVertical And area.abajo <> area.arriba Then
            area.arriba = nuevoY
        Else
            area.abajo = nuevoY
            area.invertidoVertical = False
        End If
    ElseIf nuevoY > area.arriba Then
        If area.invertidoVertical Then
            area.arriba = nuevoY
        Else
            area.abajo = nuevoY
        End If
    Else
        area.invertidoVertical = True
        area.abajo = maxl(area.arriba, area.abajo)
        area.arriba = nuevoY
    End If
            
    If nuevoX = area.izquierda Then
        If area.invertidoHorizontal And area.izquierda <> area.derecha Then
            area.izquierda = nuevoX
        Else
            area.derecha = nuevoX
            area.invertidoHorizontal = False
        End If
    ElseIf nuevoX > area.izquierda Then
        If area.invertidoHorizontal Then
            area.izquierda = nuevoX
        Else
            area.derecha = nuevoX
        End If
    Else
        area.invertidoHorizontal = True
        area.derecha = maxl(area.izquierda, area.derecha)
        area.izquierda = nuevoX

    End If
         
End Sub

Public Sub reiniciarArea(ByRef area As tAreaSeleccionada)
    area.abajo = 0
    area.arriba = 0
    area.izquierda = 0
    area.derecha = 0
    
    area.invertidoHorizontal = False
    area.invertidoVertical = False
End Sub
