Attribute VB_Name = "CLI_EfectosPisadas"
Option Explicit

Public Type tEfectoPisada
    sonido_derecha As Integer
    sonido_izquierda As Integer
End Type

Private Const archivo_compilado = "EfectosPisada.ind"

Public EfectosPisadas() As tEfectoPisada


Public Function Cargar_EfectosPisadas()

    Dim handle  As Integer
    Dim loopEfecto As Integer
    Dim cantidadEfectos As Integer
         
    handle = FreeFile()
    
    Open IniPath & archivo_compilado For Binary Access Read As handle
        
    Get handle, , cantidadEfectos
    
    ReDim EfectosPisadas(1 To cantidadEfectos) As tEfectoPisada
    
    For loopEfecto = 1 To cantidadEfectos
        Get handle, , EfectosPisadas(loopEfecto).sonido_derecha
        Get handle, , EfectosPisadas(loopEfecto).sonido_izquierda
    Next loopEfecto
    
    Close handle
End Function
