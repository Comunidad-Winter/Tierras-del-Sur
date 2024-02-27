Attribute VB_Name = "ME_modComandos"
Option Explicit

Private pilaReHacer As cPila
Private pilaDesHacer As cPila


Private Const MAXIMO_ELEMENTOS_RECORDADOS_REHACER As Byte = 128
'La pila es mas chica porque rara vez se quiere re hacer
Private Const MAXIMO_ELEMENTOS_RECORDADOS_HACER As Byte = 64

Public Sub iniciarDesHacerReHacer()
        Set pilaReHacer = New cPila
        Set pilaDesHacer = New cPila
        
        Call pilaReHacer.iniciar(MAXIMO_ELEMENTOS_RECORDADOS_HACER)
        Call pilaDesHacer.iniciar(MAXIMO_ELEMENTOS_RECORDADOS_REHACER)
End Sub

Public Sub agregarComandoADesHacer(comando As iComando)
    
    Call pilaDesHacer.Push(comando)
    
    If Not pilaReHacer.estaVacia Then
        Call pilaReHacer.vaciar
    End If
    
End Sub

Public Function hayComandosReHacer() As Boolean
    hayComandosReHacer = Not pilaReHacer.estaVacia
End Function

Public Function hayComandosDesHacer() As Boolean
    hayComandosDesHacer = Not pilaDesHacer.estaVacia
End Function

Public Function obtenerDescripcionComandoReHacer() As String
    Dim comandoAuxiliar As iComando
    
    Set comandoAuxiliar = pilaReHacer.Pop
    
    obtenerDescripcionComandoReHacer = comandoAuxiliar.obtenerNombre
    
    Call pilaReHacer.Push(comandoAuxiliar)
End Function
Public Function obtenerDescripcionComandoDeshacer() As String
    Dim comandoAuxiliar As iComando
    
    Set comandoAuxiliar = pilaDesHacer.Pop
    
    obtenerDescripcionComandoDeshacer = comandoAuxiliar.obtenerNombre
    
    Call pilaDesHacer.Push(comandoAuxiliar)
End Function

Public Sub reHacerSiguienteComando()
    Dim comando As iComando
    
    'Lo saco de la pila de cosas para re hacer
    If Not pilaReHacer.estaVacia Then
        Set comando = pilaReHacer.Pop
        
        'Lo re hago
        Call comando.hacer
        
        Call pilaDesHacer.Push(comando)
        miniMap_Redraw
    Else
        Beep
    End If
End Sub

Public Sub desHacerAnteriorComando()
    Dim comando As iComando
    
    'Lo saco de la pila de cosas para deshacer
    If Not pilaDesHacer.estaVacia Then
        Set comando = pilaDesHacer.Pop
        
        'Lo desHago
        Call comando.desHacer
        
        Call pilaReHacer.Push(comando)
        miniMap_Redraw
    Else
        Beep
    End If
End Sub
