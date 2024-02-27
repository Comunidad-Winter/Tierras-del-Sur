VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPila"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Data() As Variant
Private punteroActual As Integer 'Marca la posicion libre donde se va a poner el elemento en el push
Private punteroPisoPila As Integer
Private cantidadMaximaElementos As Integer
Private vueltaCompleta As Boolean 'Es verdadero si punteroActual le dio una vuelta al buffer


Public Sub iniciar(capacidadElementos As Integer)

    punteroActual = 0
    punteroPisoPila = 1
    vueltaCompleta = False
    cantidadMaximaElementos = capacidadElementos
    
    ReDim Data(cantidadMaximaElementos)
End Sub
Public Function Pop()

    If estaVacia Then
          Set Pop = Nothing
    Else
          Set Pop = Data(punteroActual) 'Tomo el elemento
          punteroActual = punteroActual - 1
      
          If punteroActual = 0 And vueltaCompleta Then
              punteroActual = UBound(Data) 'Vamos hacia el final del vector
              vueltaCompleta = False
          End If
    End If

End Function

Public Sub Push(Value)

    punteroActual = punteroActual + 1
    
    If punteroActual > UBound(Data) Then
        punteroActual = 1
        vueltaCompleta = True
    End If
    
    Set Data(punteroActual) = Value
    
    If vueltaCompleta And punteroActual = punteroPisoPila Then
        'No tengo mas espacio virgen para poner este elemento.
        'Me comi el mas viejo
        punteroPisoPila = punteroPisoPila + 1
    End If
        
End Sub

Public Sub vaciar()

    Call iniciar(cantidadMaximaElementos)
    
End Sub

Public Function estaVacia() As Boolean
    
    estaVacia = (Not vueltaCompleta And punteroActual < punteroPisoPila) Or (vueltaCompleta And punteroActual > punteroPisoPila)

End Function
