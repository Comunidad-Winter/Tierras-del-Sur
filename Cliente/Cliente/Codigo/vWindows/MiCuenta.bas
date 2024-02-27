Attribute VB_Name = "MiCuenta"
Option Explicit

Public cuenta As cuenta
Public personajes As collection


Public Sub cerrarSesion()
    Set cuenta = Nothing
    Set personajes = Nothing
End Sub

