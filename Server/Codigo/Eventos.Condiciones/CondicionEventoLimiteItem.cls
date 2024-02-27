VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoLimiteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento
               
Private objetosLimites() As tEventoObjetoRestringido
Private ContadorTemp() As Long

Private ExclusividadDeItems As Boolean
Private listaItemsVacia As Boolean
Private NoPermitirOroEnBille As Boolean


'http://stackoverflow.com/questions/183353/how-do-i-determine-if-an-array-is-initialized-in-vb6
'CSEH: Nada
Private Function arrayEstaIniciado(vector() As modEvento.tEventoObjetoRestringido) As Boolean

    On Error GoTo ProcError
    Dim lTmp As Long

    arrayEstaIniciado = False

    lTmp = UBound(vector) ' Acá puede saltar el error

    arrayEstaIniciado = (lTmp > -1)
    
    Exit Function
ProcError:
    arrayEstaIniciado = False
    'El error sera  "Subscript 'out of range", caso contrario el error es por otra cosa (no es un array)
    'If Not Err.Number = 9 Then Err.Raise (Err.Number)

End Function

Friend Sub setParametros(objetosLimites_() As modEvento.tEventoObjetoRestringido, NoPermitirOtrosItems_ As Boolean, NoPermitirOroEnBille_ As Boolean)
    
    'Exclusividad de estos items?
    ExclusividadDeItems = NoPermitirOtrosItems_
    
    '¿Sin oro en la bille?
    NoPermitirOroEnBille = NoPermitirOroEnBille_
    
    'Detectamos si el Array esta iniciado
    If arrayEstaIniciado(objetosLimites_) Then
        listaItemsVacia = False
        objetosLimites = objetosLimites_
        
        ReDim ContadorTemp(1 To UBound(objetosLimites_))
    Else
        listaItemsVacia = True
    End If
    
End Sub

Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String

    Dim loopObjeto As Byte
    
    'Opciones
    ' ExclusividadDeItems = Si, listaItemsVacia (NO hay una lista de objetos) = Si:
            ' Personajes desnudos
            
    ' ExclusividadDeItems = Si, listaItemsVacia = NO
            ' Solo puede tener los items que estan en la lista cumpliendo los maximos y minimos
            
    ' ExclusividadDeItems = No, listaItemsVacia = Si
            ' No tiene sentido. Si no hay una resitrcción a los objetos, no tiene sentido haber creado esta condicion.
            
    ' ExclusividadDeItems = No, listaItemsVacia = No
            ' Puede tener todos los items que quiera pero deben cumplir con los minimos y maximos
            
    'Hay exclusivdad de items?
    If ExclusividadDeItems = True Then
        ' Detallamos si esta desnudo
        If listaItemsVacia = True Then
            iCondicionEvento_obtenerDescripcion = "No se permite ningún item. Personajes desnudos."
        Else
            iCondicionEvento_obtenerDescripcion = "Solo se permiten los items restringidos."
        End If
    End If
    
    ' ¿Oro en la billetera?
    If NoPermitirOroEnBille = True Then
        iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " Los personajes no pueden tener oro en la billetera."
    End If
    
    ' Hay una lista de restricciones sobre los objetos
    If listaItemsVacia = False Then
        iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & "Restricciones de items:"
        
        ' Listamos los items y sus limitaciones
        For loopObjeto = LBound(objetosLimites) To UBound(objetosLimites)
            If objetosLimites(loopObjeto).id > 0 Then
                If objetosLimites(loopObjeto).cantidad = 0 Then
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & vbCrLf & _
                    ObjData(objetosLimites(loopObjeto).id).Name & ": no permitido su uso."
                Else
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & vbCrLf
                    
                    'Esta cantidad es como mínimo o maximo?
                    If objetosLimites(loopObjeto).tipo = eRangoLimite.minimo Then
                        iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & "-- " & ObjData(objetosLimites(loopObjeto).id).Name & ": mínimo " & objetosLimites(loopObjeto).cantidad & " unidades."
                    Else
                        iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & "-- " & ObjData(objetosLimites(loopObjeto).id).Name & ": hasta " & objetosLimites(loopObjeto).cantidad & " unidades."
                    End If
                End If
            End If
         Next loopObjeto
       
    End If
    
End Function
Private Sub limpiarContador()
    Dim loopItemLimite As Byte
    
     If listaItemsVacia = False Then
        For loopItemLimite = 1 To UBound(objetosLimites)
            ContadorTemp(loopItemLimite) = 0
        Next loopItemLimite
    End If
End Sub
Private Function obtenerPosicionDeItemsEnVectorLimitado(ObjIndex As Integer) As Byte
    Dim loopItemLimite As Byte
    
    If listaItemsVacia Then
        obtenerPosicionDeItemsEnVectorLimitado = 0
    Else
        For loopItemLimite = 1 To UBound(objetosLimites)
            If objetosLimites(loopItemLimite).id = ObjIndex Then
                'Lo encontre
                obtenerPosicionDeItemsEnVectorLimitado = loopItemLimite
                Exit Function
            End If
        Next loopItemLimite
    End If
    'El item no esta en la lista de item limitados
    obtenerPosicionDeItemsEnVectorLimitado = 0
End Function
Public Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean
    
Dim loopIntegrante As Byte
Dim loopItem As Byte
Dim Posicion As Byte

iCondicionEvento_puedeIngresarEquipo = False

For loopIntegrante = 1 To UBound(integrantesIndex)

    Call limpiarContador ' O(L)
    
    'Recorro el inventario del usuario
    If NoPermitirOroEnBille Then
        If UserList(integrantesIndex(loopIntegrante)).Stats.GLD > 0 Then
            ' Tiene oro cuando no esta permitido
            Exit Function
        End If
    End If
        
        
    'O(i x l)
    With UserList(integrantesIndex(loopIntegrante)).Invent
    
        For loopItem = 1 To UserList(integrantesIndex(loopIntegrante)).Stats.MaxItems
            
            If .Object(loopItem).ObjIndex > 0 Then
                Posicion = obtenerPosicionDeItemsEnVectorLimitado(.Object(loopItem).ObjIndex)
    
                If Posicion = 0 And ExclusividadDeItems Then
                    'Tiene un item que no puede traer
                    Exit Function
                ElseIf Posicion > 0 Then ' El item esta limitado? Entonces lo contamos
                    ContadorTemp(Posicion) = ContadorTemp(Posicion) + .Object(loopItem).Amount
                End If
            End If
       Next loopItem
    End With
    
    ' O(L)
    ' Chequeamos que la cantidad de items que tenga esta de acuerdo a las restricciones
    If listaItemsVacia = False Then
        For loopItem = 1 To UBound(objetosLimites)
            If objetosLimites(loopItem).tipo = eRangoLimite.minimo Then
                If ContadorTemp(loopItem) < objetosLimites(loopItem).cantidad Then
                    'Pepito tiene menos items de los que deberia
                    Exit Function
                End If
            Else
                If ContadorTemp(loopItem) > objetosLimites(loopItem).cantidad Then
                    'Pepito tiene mas items de los que deberia
                    Exit Function
                End If
            End If
        Next loopItem
    End If
Next loopIntegrante

iCondicionEvento_puedeIngresarEquipo = True

End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantesid() As Long)
End Sub

Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
End Sub

Private Function contarCantidadItems(inventario As inventario, IndexItem As Integer) As Long
    
    contarCantidadItems = 0
    
    Dim loopItemInventario As Byte
    
    For loopItemInventario = 1 To UBound(inventario.Object)
            If inventario.Object(loopItemInventario).ObjIndex = IndexItem Then
                contarCantidadItems = contarCantidadItems + inventario.Object(loopItemInventario).Amount
            End If
    Next loopItemInventario
    
End Function
