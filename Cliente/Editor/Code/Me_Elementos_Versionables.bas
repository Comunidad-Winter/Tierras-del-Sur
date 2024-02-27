Attribute VB_Name = "Me_Elementos_Versionables"
Option Explicit

Private Type tElementoVersionable
    nombre As String
    archivo As String
End Type

Private elementosVersionables(1 To 1) As tElementoVersionable

Public Type datosElemento
    eliminados() As Integer
    modificados() As Integer
    creados() As Integer
End Type

Public Sub iniciar()

    elementosVersionables(1).nombre = "PREDEFINIDOS"
    elementosVersionables(1).archivo = DatosPath & "presets.ini"

End Sub


Public Function obtenerDatos(idelemento As Byte) As datosElemento


    'Si es un .ini
    'Cargo el archivo. Recorro cada uno de los elementos.
        'Si version = 0 y Owner > 0
            'Lo agrego a CREADOS.
        'Si version = 0 y Owner = -1
            'Lo agrego a ELIMINADOS.
        'Si version = 1 y Owner > 0
            'Lo agrego a modificados.
        ' Si version = 0

End Function
