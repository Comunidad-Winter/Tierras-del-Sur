VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditorAccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Acción"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7785
   Icon            =   "EditorAccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
   End
   Begin VB.ListBox listaDisponibles 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar Nodo"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _Version        =   393217
      Indentation     =   118
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblEjecForzada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"EditorAccion.frx":1CCA
      Height          =   390
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   7380
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEscribirParaEditar 
      BackStyle       =   0  'Transparent
      Caption         =   "- Para cambiarle el nombre a un nodo padre, seleccionarlo y presionar BACKSPACE"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   7335
   End
   Begin VB.Label cmdAyuda 
      Caption         =   "- Doble Click para editar."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "Accionr Actual"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Acciones disponibles"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditorAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private test As Collection
Private accionPadre As cAccionCompuestaEditor
Private nodoTildado As MSComctlLib.node


Public Sub Cargar(modificando As cAccionCompuestaEditor)
    Set accionPadre = modificando
    Me.caption = "Editando acción " & UCase$(accionPadre.iAccionEditor_getNombre)
End Sub
Private Sub cmdAgregar_Click()

Dim nodo As node
Dim accion As iAccionEditor
Dim nuevaAccion As cAccionTileEditor
Dim auxAccionCompuesta As cAccionCompuestaEditor


If Me.listaDisponibles.ListIndex = -1 Then
    MsgBox "Debe seleccionar una Acción de la lista de acciones disponibles y luego hacer clic en aceptar.", vbInformation, "Editor de acciones"
    Exit Sub
End If

Set nodo = Me.TreeView.SelectedItem

If nodo Is Nothing Then
    MsgBox "Debe seleccionar alguna Accion de la lista de Accion Actual en donde se va a agregar esta accion.", vbInformation, "Editor de acciones"
    Exit Sub
End If


Set accion = test.item(CInt(mid(nodo.key, 2)))

Set nuevaAccion = listaAccionTileEditor.item(Me.listaDisponibles.ListIndex + 1).Clonar


Call frmModificarAccion.Cargar(nuevaAccion)

If frmModificarAccion.edicion(Me) Then
    If accion.getTIPO = 1 Then
        Set auxAccionCompuesta = accion
        Call auxAccionCompuesta.agregarHijo(nuevaAccion)
    Else
        
        Set auxAccionCompuesta = New cAccionCompuestaEditor
        Call auxAccionCompuesta.iAccionEditor_crear(InputBox("Ingrese un nombre descriptivo"), "")
        
        Call auxAccionCompuesta.agregarHijo(accion)
        Call auxAccionCompuesta.agregarHijo(nuevaAccion)
        
        Dim accionPadre As cAccionCompuestaEditor
        
        Set accionPadre = test.item(CInt(mid$(nodo.parent.key, 2)))
        Call accionPadre.cambiar(accion, auxAccionCompuesta)
    End If
    
    Call refrescarArbol(Me.TreeView)
    MsgBox "Acción agregada", , "Editor de acciones"
End If

End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Function obtenerCantidadHijos(nodo As node) As Integer
    obtenerCantidadHijos = 0
    
    If nodo.FirstSibling = nodo.LastSibling Then
        obtenerCantidadHijos = 0
    Else
        obtenerCantidadHijos = 1
    End If
End Function

Private Sub eliminarNodo(nodo As node)
    Dim accionHijo As iAccionEditor
    Dim accionPadre As cAccionCompuestaEditor
    
    Set accionHijo = test.item(CInt(mid(nodo.key, 2)))
    
    If nodo.parent Is Nothing Then
        Call ME_modAccionEditor.eliminarAccion(accionHijo.getID)
        Unload Me
    Else
        Set accionPadre = test.item(CInt(mid(nodo.parent.key, 2)))
        
        Call accionPadre.eliminarHijo(accionHijo)
        
        If obtenerCantidadHijos(nodo) = 0 Then
            Call eliminarNodo(nodo.parent)
        End If
   End If
End Sub
Private Sub cmdEliminar_Click()

Dim nodo As node

Set nodo = Me.TreeView.SelectedItem

Dim resultado As VbMsgBoxResult

resultado = MsgBox("¿Esta seguro que desea eliminar esta acción?", vbYesNo + vbInformation, "Editor de Acciones")

If resultado = vbYes Then
    Call eliminarNodo(nodo)
End If

Call refrescarArbol(Me.TreeView)
End Sub

Private Sub Form_Load()
    Call ME_modAccionEditor.refrescarListaDisponibles(Me.listaDisponibles)
    Call refrescarArbol(Me.TreeView)
End Sub

Private Sub refrescarArbol(destino As TreeView)
    destino.Nodes.Clear
    Call cargarArbol(destino, accionPadre)

    If Me.TreeView.Nodes.Count > 0 Then
        Me.TreeView.Nodes.item(1).Expanded = True
    End If
    
    destino.Refresh
End Sub
Private Sub cargarArbol(destino As TreeView, contenido As cAccionCompuestaEditor)
    
    Dim numeroElemento As Integer
      
    destino.Nodes.Add , , "N" & 1, contenido.iAccionEditor_getNombre
    destino.Nodes.item(1).Checked = True
    
    numeroElemento = 2
    Set test = New Collection
    Call test.Add(contenido)
    
    Call cargarNodo(destino.Nodes, contenido.obtenerHijos, contenido, 1, numeroElemento)
End Sub
Private Sub cargarNodo(destino As Nodes, contenido As Collection, padre As cAccionCompuestaEditor, numeroPadre As Integer, ByRef numeroElemento As Integer)
    
    Dim accion As iAccionEditor
    Dim accionCompuesta As cAccionCompuestaEditor
    Dim numeroRelativoElemento As Byte
    Dim chequeado As Boolean
    
    numeroRelativoElemento = 1
    
    For Each accion In contenido
    
        test.Add accion
        
        chequeado = padre.seEjecutaSiempre(numeroRelativoElemento)

        If accion.getTIPO = 0 Then
            destino.Add "N" & numeroPadre, tvwChild, "N" & numeroElemento, accion.getNombreExtendido
            
            destino.item(numeroElemento).Checked = chequeado
            
            numeroElemento = numeroElemento + 1
        
        Else
            destino.Add "N" & numeroPadre, tvwChild, "N" & numeroElemento, accion.GetNombre
            destino.item(numeroElemento).Checked = chequeado
            
            numeroElemento = numeroElemento + 1
            
            Set accionCompuesta = accion
            Call cargarNodo(destino, accionCompuesta.obtenerHijos, accionCompuesta, numeroElemento - 1, numeroElemento)
        End If
        

        numeroRelativoElemento = numeroRelativoElemento + 1
    Next
    
End Sub

Private Sub TreeView_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim posicionMinimaEdicion As Integer
Dim accion As iAccionEditor
    
Set accion = test.item(CInt(mid(Me.TreeView.SelectedItem.key, 2)))
    
If accion.getTIPO = 1 Then
    Call accion.SetNombre(Trim(NewString))
Else
    MsgBox "No se le puede cambiar el nombre a una acción atomica."
End If


End Sub

Private Sub TreeView_DblClick()
    Dim nodo As node
    Set nodo = Me.TreeView.SelectedItem
    
    Dim accion As iAccionEditor
    
    Set accion = test.item(CInt(mid(nodo.key, 2)))
    
    If accion.getTIPO = 0 Then
        Call frmModificarAccion.Cargar(accion)
        frmModificarAccion.Show vbModal, Me
        nodo.Text = accion.getNombreExtendido
    End If
End Sub


Private Sub TreeView_KeyDown(KeyCode As Integer, Shift As Integer)

Dim accion As iAccionEditor

If KeyCode = vbKeyBack Then

    Set accion = test.item(CInt(mid(Me.TreeView.SelectedItem.key, 2)))
    
    If Not accion.getTIPO = 0 Then
        Me.TreeView.StartLabelEdit
    End If
End If
End Sub

Private Function obtenerNumeroHijo(padre As node, hijo As node) As Byte
    Dim posibleHijo As node
    
    obtenerNumeroHijo = 1
    
    Set posibleHijo = padre.Child.FirstSibling
    
    Do While ((Not posibleHijo Is hijo) And (Not posibleHijo Is padre.Child.LastSibling))
        Set posibleHijo = posibleHijo.Next
        obtenerNumeroHijo = obtenerNumeroHijo + 1
    Loop
    
    
End Function

Private Sub TreeView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not nodoTildado Is Nothing Then
        nodoTildado.Checked = True
    End If
End Sub

Private Sub TreeView_NodeCheck(ByVal node As MSComctlLib.node)
Dim accion As cAccionCompuestaEditor
Dim numeroHijo As Byte

Set nodoTildado = Nothing

If node.parent Is Nothing And node.Checked = False Then
    Set nodoTildado = node
    MsgBox "La acción raiz siempre se ejecuta. No se puede cancelar. Si desea que esta accion no haga nada saquela del tile donde la puso.", vbInformation, "Editor de acciones"
    Exit Sub
End If

numeroHijo = obtenerNumeroHijo(node.parent, node)

If numeroHijo = 1 Then
    Set nodoTildado = node
    MsgBox "La primer accion de un conjunto siempre se ejecuta y es la que marca la pauta para la ejecución de las siguientes acciones.", vbInformation, "Editor de acciones"
    Exit Sub
End If

Set accion = test.item(CInt(mid(node.parent.key, 2)))

Call accion.establecerAccionar(numeroHijo, node.Checked)

End Sub

