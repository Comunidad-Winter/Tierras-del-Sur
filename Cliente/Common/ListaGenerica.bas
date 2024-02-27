Attribute VB_Name = "ListaGenerica"
'No usar. Sólo para tener de base para otros códigos.
'Menduz.


Option Explicit

Private Type udtListaGenerica
    active          As Byte
    ID              As Long 'hash?
End Type

Private Lista()     As udtListaGenerica
Private ListaMax    As Integer
Private ListaCount  As Integer
Private ListaLast   As Integer

Private Sub Lista_Iniciar(ByVal Maximo As Integer)
    ListaMax = Maximo
    Lista_ReIniciar
End Sub

Private Sub Lista_ReIniciar()
    ReDim Lista(ListaMax)
    ListaCount = 0
    ListaLast = 0
End Sub

Private Function Lista_Agregar(ByVal ID As Long) As Integer
    Dim i As Integer
    Lista_Agregar = Lista_ObtenerLibre
    
    If Lista_Agregar <> -1 Then
        With Lista(Lista_Agregar)
            .active = 1
            .ID = ID
        End With
        If Lista_Agregar > ListaLast Then ListaLast = Lista_Agregar
        ListaCount = ListaCount + 1
    End If
End Function

Private Function Lista_ObtenerLibre() As Integer
    Dim i As Integer
    
    If ListaCount < ListaMax Then 'nos aseguramos de que haya espacio
        For i = 0 To ListaMax
            If Lista(i).active = 0 Then
                Lista_ObtenerLibre = i
                Exit Function
            End If
        Next i
    End If
    
    Lista_ObtenerLibre = -1
End Function

Private Function Lista_Remover(ByRef Index As Integer) As Boolean 'True cuando se cambia el active de verdadero a falso
    Dim i As Integer
    If ListaCount Then
        If Index <= ListaLast Then
            Lista_Remover = Lista(Index).active
            Lista(Index).active = 0
            ListaCount = ListaCount - 1
            If Index = ListaLast Then
                For i = ListaLast To 0 Step -1
                    If Lista(i).active Then
                        ListaLast = i
                        Exit For
                    End If
                Next i
                If ListaLast = Index Then ListaLast = 0
            End If
        End If
    End If
End Function

Private Function Lista_Buscar(ByVal ID As Long) As Integer

    Dim i As Integer
    If ListaCount Then
        For i = 0 To ListaLast
            If Lista(i).ID = ID Then
                Lista_Buscar = i
                Exit Function
            End If
        Next i
    End If
    
    Lista_Buscar = -1
End Function

Private Function CercaniaBusqueda(ByVal Index As Integer) As Single
'f(x)<0 cuando tiene que buscar de el max a cero
'f(x)>0 cuando tiene que buscar de cero a max
'f(x)=0 cuando x==ListaLast
'              Index
'              / | \
'             /  |  \
'            /   |   \
'           /    |    \
'Lista[ Cero ListaLast ListaMax ]

    If Index = ListaLast Then
        'nada f(x)=0
    ElseIf Index < ListaMax - Index Then
        'el index está más cerca del cero que de el centro
    Else
        'el index estmá más cerca de el final de la lista que del principio
    End If

End Function

