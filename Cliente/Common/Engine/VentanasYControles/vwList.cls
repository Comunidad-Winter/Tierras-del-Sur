VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Option Explicit


' Eventos
Public Event Change(SelectedIndex As Integer)
Public Event HoverChange(HoverIndex As Integer)

'Propiedades
Private CX%, cY%, cH%, cW%                            ' Posición X, Y del formulario y su ancho y alto

Private ControlRect                 As RECT

Private items                       As collection     ' Lista de Items



Private CantidadFilasAMostrar       As Integer        ' Cantidad de filas que como máximo se van a mostrar


Private cEnabled                    As Boolean        ' Control habilitado?

Private m_SelectedIndex             As Integer        ' De 1 a N, pero retorno de 0 a N - 1, INDEX seleccionado.
Private pHoverItem                  As Integer        ' De 1 a 1, pero retorno de 0 a N - 1 INDEX sobre cual esta el mouse.

Private OrdenSeleccionado           As Integer        ' De 1 a N
Private OrdenMuestroMinimo          As Integer        ' La primer final que se está mostrnado. De 1 a N
Private OrdenMaximo                 As Integer        ' El maximo elemento visible. N


Private PosScrollbar                As New clsAlpha

Private Sub Class_Initialize()
    ' Lista de items
    Set items = New collection

    cEnabled = True
    cW = 128
    cH = 64
    ReDoRECT
    PosScrollbar.Speed = 1000
    PosScrollbar.InitialValue = 0
    
    pHoverItem = 0
    m_SelectedIndex = 0
End Sub

Private Sub ReDoRECT()

    ' Espacio que ocupa
    With ControlRect
        .top = cY
        .bottom = cY + cH
        .left = CX
        .right = CX + cW
    End With
    
    ' Cantidad de items máximo que se van a mostrar
    CantidadFilasAMostrar = (cH / 16) - 1
End Sub


' Vacia la Lista
Public Sub Clear()
    Dim i As Integer
    Dim cantidad As Integer
    
    m_SelectedIndex = 0
    OrdenSeleccionado = 0
    OrdenMuestroMinimo = 1
    
    cantidad = items.count
    
    For i = cantidad To 1 Step -1
        Call items.Remove(i)
    Next i
    
End Sub

' Aggrega un elemento a la Lista
Public Function AddItem(texto As String) As Integer

    Dim nuevoItem As ControlListItem
    
    ' Creamos el item
    Set nuevoItem = New ControlListItem
    
    ' Le establecemos el texto
    OrdenMaximo = OrdenMaximo + 1
    
    nuevoItem.texto = texto
    nuevoItem.Orden = OrdenMaximo
    
    ' Lo agregamos
    Call items.Add(nuevoItem)
    
    ' Devolvemos el index de este elemento
    AddItem = items.count - 1
    
    ' Regeneramos el Orden de los elementos visibles
    Call RegenerarOrden
End Function

' Devolvemos la cantidad de elementos que hay en la lista
Public Property Get count() As Integer
    count = items.count
End Property

' Devuelve el Index del Elemento Seleccionado
Public Property Get SelectedIndex() As Integer
    SelectedIndex = m_SelectedIndex - 1
End Property

' Selecciona un Index
Public Property Let SelectedIndex(val As Integer)
    
    If m_SelectedIndex > 0 Then items(m_SelectedIndex).seleccionado = False

    m_SelectedIndex = maxl(minl(items.count, val + 1), 0)
    
    If m_SelectedIndex > 0 Then items(m_SelectedIndex).seleccionado = True
    
End Property

' Guarda información util del elemento
Public Property Get itemData(ByVal Index As Long) As String
    itemData = items(Index + 1).Data
End Property

Public Property Let itemData(ByVal Index As Long, Data As String)
    items(Index + 1).Data = Data
End Property

' Cambia la visibilidad de un elemento
Public Property Get ItemVisible(ByVal Index As Long) As Boolean
    ItemVisible = items(Index + 1).Visible
End Property

Public Property Let ItemVisible(ByVal Index As Long, Visible As Boolean)
    items(Index + 1).Visible = Visible
    
    Call RegenerarOrden
End Property

' Cambia el texto que representa al item
Public Property Let ItemTexto(ByVal Index As Long, texto As String)
    items(Index + 1).texto = texto
End Property

Public Property Get ItemTexto(ByVal Index As Long) As String
    ItemTexto = items(Index + 1).texto
End Property

' ****************************************************************************'
' Dibujo del CONTROL
Public Sub Render(Optional ByVal focus As Boolean = False)
    Dim loopElementoVisible As Integer
    Dim posicionAcumuladaY As Integer
    Dim restantes As Integer
    Dim item As ControlListItem
    Dim alto_barrita As Single
    Dim posBarrita As Single
    
    ' Color del Borde de la Lista
    If focus Then
        Grh_Render_Solid &HFF88AA88, CX - 2, cY - 2, cW + 4, cH + 4
        Grh_Render_Solid &HFF010101, CX - 1, cY - 1, cW + 2, cH + 2
    Else
        Grh_Render_Solid &HFF000000, CX - 2, cY - 2, cW + 4, cH + 4
        Grh_Render_Solid &HFF0F0F0F, CX - 1, cY - 1, cW + 2, cH + 2
    End If

    ' Dibujamos cada una de las Lineas visibles
    posicionAcumuladaY = cY ' A partir de donde comienzo a dibujar los elementos
    restantes = CantidadFilasAMostrar ' Cuantas filas debo mostrar
    
    loopElementoVisible = 0 ' Contador de Elementos Visible hasta el momento
    
    ' Recorremos todos los items
    For Each item In items
        
        
        If item.Orden >= OrdenMuestroMinimo Then
           If item.Visible Then
                ' ¿Esta selecionada?
                If item.seleccionado Then Grh_Render_Solid &HFF555555, CX, posicionAcumuladaY, cW - 13, 16
        
                ' Lo ponemos en la posicion correspondiente y lo dibujamos
                Call item.clsGUIText.SetPos(CX, cY + posicionAcumuladaY)
                Call item.clsGUIText.Render
                    
                ' Actualizo acumuladores
                restantes = restantes - 1 ' Uno menos que dibujar
                posicionAcumuladaY = posicionAcumuladaY + 16
    
                If restantes = 0 Then Exit For
            End If
        End If
    Next
          
    ' La barrra de Scroll
    Grh_Render_Solid &HFF333333, CX + cW - 12, cY, 12, cH
    
    ' ¿Es necesaria?
   If OrdenMaximo > CantidadFilasAMostrar Then
   
        ' Calculamos el alto de la barra
        alto_barrita = CantidadFilasAMostrar * cH / OrdenMaximo
        
        If focus Then
            Grh_Render_Solid &HFF33AAAA, CX + cW - 11, cY + Round(PosScrollbar) + 1, 10, alto_barrita - 2
        Else
            Grh_Render_Solid &HFFAAAAAA, CX + cW - 11, cY + Round(PosScrollbar) + 1, 10, alto_barrita - 2
        End If
        
        Dim a As Single
        a = OrdenMuestroMinimo + (CantidadFilasAMostrar / 2)
        posBarrita = (cH - 1) * (a / OrdenMaximo) - (alto_barrita / 2)
        posBarrita = maxl(posBarrita, 0)
        posBarrita = minl(posBarrita, (cH - 1) - alto_barrita)
        PosScrollbar = posBarrita

    End If
    
End Sub

' Calculamos el Orden de cada Elemento.
' Esto quiere decir el número dentro de los elementos visibles
Private Sub RegenerarOrden()

    Dim loopElementoVisible As Integer
    Dim item As ControlListItem
           
    loopElementoVisible = 0 ' Contador de Elementos Visible hasta el momento
    
    ' Recorremos todos los items
    For Each item In items
        If item.Visible Then
            loopElementoVisible = loopElementoVisible + 1
            
            item.Orden = loopElementoVisible
        Else
            item.Orden = 0
        End If
    Next
    OrdenMaximo = loopElementoVisible
End Sub

Private Sub Bajar()
End Sub

Private Sub Subir()
End Sub


Private Function IControl_Click(ByVal x As Integer, ByVal y As Integer) As Boolean
End Function

Private Property Get IControl_IsFocuseable() As Boolean
    IControl_IsFocuseable = True
End Property

Private Property Get IControl_IsMouseFocus() As Boolean
    IControl_IsMouseFocus = True
End Property

Private Function IControl_KeyDown(KeyCode As Integer, Shift As Integer) As Boolean
End Function

Private Function IControl_KeyPress(KeyAscii As Integer) As Boolean
End Function

Private Function IControl_KeyUp(KeyCode As Integer, Shift As Integer) As Boolean
    If KeyCode = vbKeyUp Then
        Subir
    ElseIf KeyCode = vbKeyDown Then
        Bajar
    End If
End Function

Private Function obtenerIndexPorOrden(Orden As Integer) As Integer

    Dim item As ControlListItem
    Dim Index As Integer
    
    obtenerIndexPorOrden = 0
    Index = 0
    For Each item In items
        Index = Index + 1
        If item.Orden = Orden Then
            obtenerIndexPorOrden = Index
            Exit Function
        End If
    Next
    
End Function
Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim tx As Integer
    Dim ty As Integer
    Dim nuevoOrdenSeleccionado As Integer
    Dim nuevoIndexSeleccionado As Integer
    
    If cEnabled Then IControl_MouseDown = ColisionRect(x, y, ControlRect)

    ' ¿Estoy sobre algo?
    If IControl_MouseDown Then

        tx = x - CX
        ty = y - cY
        
        ' ¿Estoy Sobre Algun Item?
        If tx < (cW - 12) Then
        
            nuevoOrdenSeleccionado = ty \ 16 + OrdenMuestroMinimo
                
            If nuevoOrdenSeleccionado <= OrdenMaximo Then
                OrdenSeleccionado = nuevoOrdenSeleccionado
                nuevoIndexSeleccionado = obtenerIndexPorOrden(OrdenSeleccionado)
                
                If m_SelectedIndex > 0 Then
                    items(m_SelectedIndex).seleccionado = False
                End If
                
                m_SelectedIndex = nuevoIndexSeleccionado
                
                items(nuevoIndexSeleccionado).seleccionado = True
                      
                RaiseEvent Change(nuevoIndexSeleccionado - 1)
            End If
                
        Else ' Estoy Sobre la Barra?
        
        
            ' Calculo el elemento inicial desde el cual debo empezar a mostrar
            OrdenMuestroMinimo = ((cH - ty) / cH) * (OrdenMaximo - CantidadFilasAMostrar)
            OrdenMuestroMinimo = maxl(OrdenMuestroMinimo, 1)
        End If
    End If
End Function


Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Boolean
    Dim tx As Integer
    Dim ty As Integer
    Dim nuevoIndexHover As Integer
    Dim nuevoOrdenSeleccionado As Integer
    
    If cEnabled Then IControl_MouseMove = ColisionRect(x, y, ControlRect)
    
    If IControl_MouseMove Then

        tx = x - CX
        ty = y - cY
        
        ' ¿Hice clic?
        If Button Then
        
            ' Click en el scrollbar
            If tx > (cW - 12) Then
                OrdenMuestroMinimo = ((ty - cH / 2) / cH) * OrdenMaximo + 1
                OrdenMuestroMinimo = maxl(OrdenMuestroMinimo, 0)
            End If
            HoverItem = 0
        Else
            ' Paso el Mouse por un Item
            If tx < (cW - 12) Then
                ' Obtenemos el elemento visible sobre el cual esta
                nuevoOrdenSeleccionado = (ty \ 16) + OrdenMuestroMinimo
                
                ' No es valido?
                If nuevoOrdenSeleccionado > OrdenMaximo Then
                    HoverItem = 0
                Else
                    ' Es el que ya tenia? Evito re mandar el evento
                    nuevoIndexHover = obtenerIndexPorOrden(nuevoOrdenSeleccionado)
  
                    If pHoverItem <> nuevoIndexHover Then
                        OrdenSeleccionado = nuevoOrdenSeleccionado
                        HoverItem = nuevoIndexHover
                    End If
                    
                End If
            End If
        End If
        
        If z = 1 Then
            If OrdenMuestroMinimo > 1 Then OrdenMuestroMinimo = OrdenMuestroMinimo - 1
        ElseIf z = -1 Then
            If OrdenMuestroMinimo + CantidadFilasAMostrar - 1 < OrdenMaximo Then OrdenMuestroMinimo = OrdenMuestroMinimo + 1
        End If
    Else
        If pHoverItem > 0 Then HoverItem = 0
    End If
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If cEnabled Then
        IControl_MouseUp = ColisionRect(x, y, ControlRect)
    End If
End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    Render focus
End Sub

Private Sub IControl_SetPos(ByVal x As Integer, ByVal y As Integer)
    SetPos x, y
End Sub

Private Sub IControl_SetSize(ByVal w As Integer, ByVal h As Integer)
    SetSize w, h
End Sub

Public Sub SetPos(ByVal x As Integer, ByVal y As Integer)
    CX = x
    cY = y
    ReDoRECT
End Sub

Public Sub SetSize(ByVal w As Integer, ByVal h As Integer)
    cW = w
    cH = h
    ReDoRECT
End Sub

Public Function item(Index As Integer) As String
Attribute item.VB_UserMemId = 0
    If Index > 1 And Index <= items.count Then
        item = items(Index + 1).texto
    End If
End Function

Public Property Get HoverItem() As Integer
    HoverItem = pHoverItem
End Property

Public Property Let HoverItem(ByVal v As Integer)
    ' Si tenia alguno marcado, lo desmarcó
    If pHoverItem > 0 And pHoverItem <= items.count Then
        items.item(pHoverItem).clsGUIText.Color = -1
    End If
        
    pHoverItem = v
    ' Marco al nuevo
    If pHoverItem > 0 Then
         items.item(pHoverItem).clsGUIText.Color = &HFFAAEEFF
    End If
    
    RaiseEvent HoverChange(pHoverItem - 1)
End Property
