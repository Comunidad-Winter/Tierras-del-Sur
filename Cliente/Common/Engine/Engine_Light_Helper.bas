Attribute VB_Name = "Engine_Light_Helper"
Option Explicit

Public Enum eFlagsLuces
    luzTieneBrillo = 1
    
    luzEsLlama = 8 ' Efecto de fuego
    
    luzEsCuadrada = 16
End Enum

Public Function EsLuzValida(ByVal radio As Byte, ByVal brillo As Byte, ByVal tipo As eFlagsLuces) As Boolean
    If tipo And luzTieneBrillo Then
        If radio < 3 Then Exit Function
        If brillo = 0 Then Exit Function
    Else
        If Not (tipo And luzEsCuadrada) And radio < 3 Then Exit Function
    End If
    
    EsLuzValida = True
End Function

Public Function obtener_hora_fraccion(ByVal fraccion As Byte) As String
    Dim Minutos As Integer
    Dim Hora As Integer
    Minutos = fraccion - 1
    Minutos = Minutos * 15
    
    Hora = Minutos \ 60
    Minutos = Minutos Mod 60
    
    obtener_hora_fraccion = IIf(Hora < 10, "0", "") & Hora & ":" & IIf(Minutos < 10, "0", "") & Minutos
End Function

