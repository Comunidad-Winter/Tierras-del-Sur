VERSION 5.00
Begin VB.UserControl ListaConBuscador 
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2265
   ScaleWidth      =   2550
   Begin VB.Timer timerRetardador 
      Interval        =   500
      Left            =   600
      Top             =   1200
   End
   Begin VB.TextBox txtBuscador 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin VB.ListBox lstListaContenido 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   0
      TabIndex        =   0
      Top             =   280
      Width           =   2535
   End
End
Attribute VB_Name = "ListaConBuscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Cantidad de slots en donde se pueden guardar datos
Private cantidadAlocada As String
' Cantidad de datos guardados
Private cantidadElementos As Integer
' Nuemro del ultimo slot de almacenamiento utilizado
Private ultimoSlotUtilizado As Integer
Private pAdmitirElementoNulo As Boolean

' Aca se guarda la información de cada elemento
Private Type tElemento
     Nombre As String
     id As Integer
End Type

' Ultimo string buscado en la lista
Private ultimaBusqueda As String

' Tag del elemento
Private tagInfo As String

' Elementos almacenados
Private elementos() As tElemento

' Eventos
Public Event Change(valor As String, id As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DblClic()
Public Event Desactivado()

Private Const SLOT_LIBRE_ID = -2
Private Const ELEMENTO_NULO = -1

Private Sub lstListaContenido_Click()
    RaiseEvent Change(Me.obtenerValor, Me.obtenerIDValor)
End Sub

Private Sub lstListaContenido_DblClick()
   Call seleccionOk
End Sub

Private Sub lstListaContenido_GotFocus()
    txtBuscador.BackColor = vbGreen
End Sub

Private Sub lstListaContenido_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        RaiseEvent Desactivado
    End If
    
End Sub

Private Sub lstListaContenido_LostFocus()
    txtBuscador.BackColor = vbWhite
End Sub

Private Sub timerRetardador_Timer()
    Call cargarLista(txtBuscador)
    
    ultimaBusqueda = txtBuscador
    
    RaiseEvent Change(Me.obtenerValor(), Me.obtenerIDValor)
    
    timerRetardador.Enabled = False
End Sub

Private Sub txtBuscador_Change()
    
    timerRetardador.Enabled = False 'Lo deshabilito
    timerRetardador.Enabled = True 'Lo habilito
    timerRetardador.Interval = 300

End Sub

Private Sub cargarLista(Filtro As String)
    Dim i As Integer
    Dim cantidadvalidos As Integer
    
    ' Vamos a cargar en el List solo los elementos cuyo nombre matchee con Filtro
    lstListaContenido.Clear
    cantidadvalidos = 0
    
    Filtro = quitarTildes(Filtro)
    
    If pAdmitirElementoNulo Then
        lstListaContenido.AddItem ("")
        lstListaContenido.itemData(lstListaContenido.NewIndex) = ELEMENTO_NULO
    End If
    
    For i = 1 To ultimoSlotUtilizado
    
        '¿ Coincide?
        If InStr(1, quitarTildes(elementos(i).Nombre), Filtro, vbTextCompare) Then
        
            ' Agrego el elemento
            Call lstListaContenido.AddItem(elementos(i).Nombre)
            
            ' Relacionamos el ID
            lstListaContenido.itemData(lstListaContenido.NewIndex) = elementos(i).id
            
            cantidadvalidos = cantidadvalidos + 1
            
            ' Si hay 15 o más, refresco la lista para mostrar rápido resultados
            If cantidadvalidos = 15 Then lstListaContenido.Refresh

        End If
    Next
    
    ' Seleccioo el primero
    If pAdmitirElementoNulo = False Then
        If cantidadvalidos >= 1 Then lstListaContenido.Selected(0) = True
    Else
        If cantidadvalidos >= 2 Then lstListaContenido.Selected(1) = True
    End If
    

End Sub

Private Function quitarTildes(texto As String) As String
    quitarTildes = Replace$(texto, "á", "a", 1, -1, vbTextCompare)
    quitarTildes = Replace$(quitarTildes, "é", "e", 1, -1, vbTextCompare)
    quitarTildes = Replace$(quitarTildes, "í", "i", 1, -1, vbTextCompare)
    quitarTildes = Replace$(quitarTildes, "ó", "o", 1, -1, vbTextCompare)
    quitarTildes = Replace$(quitarTildes, "ú", "u", 1, -1, vbTextCompare)
End Function

Public Sub vaciar()
    Dim loopElemento As Integer
    
    ' Reiniciamos variables
    cantidadAlocada = 10
    cantidadElementos = 0
    ultimoSlotUtilizado = 0
    ultimaBusqueda = ""
    
    ' Obtenemos memoria
    ReDim elementos(1 To cantidadAlocada)

    ' Inicializo
    For loopElemento = 1 To cantidadAlocada
        elementos(loopElemento).id = SLOT_LIBRE_ID
        elementos(loopElemento).Nombre = vbNullString
    Next
    
    ' Limpiamos la lista
    lstListaContenido.Clear
    
    'Timer retardado
    timerRetardador.Enabled = False
End Sub

Private Sub txtBuscador_GotFocus()
    txtBuscador.BackColor = vbGreen
End Sub

Private Sub seleccionOk()
    txtBuscador.text = ""
    RaiseEvent DblClic
End Sub

Private Sub txtBuscador_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        RaiseEvent Desactivado
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call seleccionOk
    End If
    
End Sub

Private Sub txtBuscador_LostFocus()
    txtBuscador.BackColor = vbWhite
End Sub

Private Sub UserControl_GotFocus()
    Debug.Print "GotFocus"
End Sub

Private Sub UserControl_Initialize()
    Call vaciar
    ultimaBusqueda = ""
End Sub

Private Sub filtrarLista(texto As String)
    Dim i As Integer
    Dim cantidadActual As Integer
    Dim cantidadvalidos As Integer
    
    ' De los elementos actuales que hay en la lista, eliminamos los que no
    ' coincidan con el filtro
    i = 0
    cantidadActual = lstListaContenido.ListCount
    
    Do While (i <= cantidadActual - 1)
        
        If InStr(1, lstListaContenido.list(i), texto, vbTextCompare) = 0 Then
            ' Sino coincide, bye
            Call lstListaContenido.RemoveItem(i)
            cantidadActual = cantidadActual - 1
        Else
            i = i + 1
            cantidadvalidos = cantidadvalidos + 1
            
            If cantidadvalidos = 15 Then
                lstListaContenido.Refresh
                lstListaContenido.Selected(0) = True
            End If
        End If
    Loop

    ' Selecciono automaticamente el primero
    If cantidadActual > 0 Then lstListaContenido.Selected(0) = True



End Sub

' Actualiza el nombre de un elemento
Public Sub cambiarNombre(ByVal id As Integer, texto As String)
        
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To ultimoSlotUtilizado
    
        If elementos(i).id = id Then
            elementos(i).Nombre = texto
            
            ' Busco el elemento en la lista y actualizo lo que se muestra
            For j = 0 To lstListaContenido.ListCount - 1
                If lstListaContenido.itemData(j) = id Then
                    lstListaContenido.list(j) = texto
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i
End Sub

' Actualiza el nombre de un elemento
Public Function existe(ByVal id As Integer) As Boolean
    Dim i As Integer
    
    For i = 1 To ultimoSlotUtilizado
        If elementos(i).id = id Then existe = True
    Next i
End Function

' Devuelve el caption del elemento seleccionado de la lista
Public Function obtenerValor() As String
    If lstListaContenido.ListIndex > -1 Then
        obtenerValor = lstListaContenido.list(lstListaContenido.ListIndex)
    Else
        obtenerValor = ""
    End If
End Function

' Devuelve el identificador del elemento seleccionado
Public Function obtenerIDValor() As Integer
    If lstListaContenido.ListIndex > -1 Then
        obtenerIDValor = lstListaContenido.itemData(lstListaContenido.ListIndex)
    Else
        obtenerIDValor = 0
    End If
End Function


' Agrega un nuevo elemento, con determinado identificador
Public Sub addString(ByVal id As Integer, Contenido As String)
    
    Dim loopElemento As Integer
    Dim Slot As Integer
    
    ' Si habia aplicado un filtro, lo desaplico cargando toda la lista de nuevo
    If ultimaBusqueda <> "" Then Call cargarLista("")
    
    ' Aumento la cantidad de elementos
    cantidadElementos = cantidadElementos + 1
    
    '
    If cantidadElementos > ultimoSlotUtilizado Then
    
        If cantidadElementos > cantidadAlocada Then
            cantidadAlocada = cantidadAlocada * 2
            ReDim Preserve elementos(1 To cantidadAlocada) As tElemento
        End If
        
        Slot = cantidadElementos
        ultimoSlotUtilizado = Slot
    Else
        ' Busco un slot sin utilizar
        For loopElemento = 1 To ultimoSlotUtilizado
            If elementos(loopElemento).id = SLOT_LIBRE_ID Then
                Slot = loopElemento
            End If
        Next
    
    End If
    
    ' Almaceno la informacion
    elementos(Slot).Nombre = Contenido
    elementos(Slot).id = id
    
    ' Agrega a lo ultmo de la lista, con la ID
    Call lstListaContenido.AddItem(Contenido)
    lstListaContenido.itemData(lstListaContenido.NewIndex) = id
End Sub

Public Sub deseleccionar()
    lstListaContenido.ListIndex = -1
End Sub
Public Function seleccionarID(ByVal id As Integer) As Boolean
    Dim loopElemento As Integer

    ' Si hizo una busqueda, recargo la lista entera
    If ultimaBusqueda <> "" Then Call cargarLista("")
    
    ' Busco el elemento en la lista por su ID
    For loopElemento = 0 To lstListaContenido.ListCount - 1
        If lstListaContenido.itemData(loopElemento) = id Then
            lstListaContenido.ListIndex = loopElemento
            seleccionarID = True
            Exit Function
        End If
    Next
    seleccionarID = False
End Function

Public Sub seleccionarIndex(ByVal Index As Integer)
    ' Selecciono en base al index de la lista
    If Index > -1 And Index < lstListaContenido.ListCount Then
        Call cargarLista("")
        lstListaContenido.ListIndex = Index
    End If
End Sub

Public Sub desseleccionar()
    Call cargarLista("")
    lstListaContenido.ListIndex = -1
    txtBuscador.text = ""
End Sub

Public Sub eliminar(ByVal id As Integer)
    Dim loopElemento As Integer
    
    ' Si hizo una busqueda, recargo la lista entera
    If ultimaBusqueda <> "" Then Call cargarLista("")
        
    For loopElemento = 1 To ultimoSlotUtilizado
        If elementos(loopElemento).id = id Then
        
            ' Reseteo
            elementos(loopElemento).id = SLOT_LIBRE_ID
            elementos(loopElemento).Nombre = vbNullString
        
            ' Voy a remover el elemento de la lista
            Call lstListaContenido.RemoveItem(obtenerIndexID(id))
            
            ' Reduzco la cantidad de elementos
            cantidadElementos = cantidadElementos - 1
            
            ' ¿Era el ultimo Slot?
            If loopElemento = ultimoSlotUtilizado Then ultimoSlotUtilizado = ultimoSlotUtilizado - 1

            Exit For
        End If
    Next
End Sub

' Obtengo el index de la lista en base al id
Private Function obtenerIndexID(ByVal id As Integer) As Integer
    Dim loopElemento As Integer
    Dim Index As Integer
    
    obtenerIndexID = SLOT_LIBRE_ID
    
    For loopElemento = 0 To lstListaContenido.ListCount - 1
    
        If lstListaContenido.itemData(loopElemento) = id Then
            obtenerIndexID = loopElemento
            Exit For
        End If
        
    Next
    
End Function
Private Sub UserControl_Resize()
    lstListaContenido.Width = UserControl.Width
    lstListaContenido.Height = UserControl.Height - txtBuscador.Height
    txtBuscador.Width = UserControl.Width
End Sub

Public Function getCantidadElementos() As Long
    getCantidadElementos = cantidadElementos
End Function
Public Property Get Enabled() As Boolean
   Enabled = txtBuscador.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    txtBuscador.Enabled = vNewValue
    lstListaContenido = vNewValue
End Property

Public Property Get tag() As String
    tag = tagInfo
End Property

Public Property Let tag(ByVal vNewValue As String)
    tagInfo = vNewValue
End Property

Public Property Get AdmitirElementoNulo() As Boolean
    pAdmitirElementoNulo = tagInfo
End Property

Public Property Let AdmitirElementoNulo(ByVal vNewValue As Boolean)

    If pAdmitirElementoNulo = False Then
        Call lstListaContenido.AddItem("", 0)
        lstListaContenido.itemData(lstListaContenido.NewIndex) = ELEMENTO_NULO
    End If
    
    pAdmitirElementoNulo = vNewValue
End Property

