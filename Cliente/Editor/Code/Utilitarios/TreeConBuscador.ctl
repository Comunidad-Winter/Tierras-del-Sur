VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TreeConBuscador 
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2625
   ScaleWidth      =   2535
   Begin MSComctlLib.TreeView lstListaContenido 
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Timer timerRetardador 
      Enabled         =   0   'False
      Left            =   720
      Top             =   1800
   End
   Begin VB.TextBox txtBuscador 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   150
      Width           =   2535
   End
   Begin MSComctlLib.TreeView arbolBackups 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   430
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   106
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "TreeConBuscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event change(valor As String, id As Integer, esPadre As Boolean)
Private tagInfo As String

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Private Sub arbolBackups_GotFocus()
    txtBuscador.BackColor = vbGreen
End Sub

Public Function list(indexNodo As Integer) As String
     list = arbolBackups.Nodes.item(indexNodo).text
End Function

Public Function estaChequeado(indexNodo As Integer) As Boolean
     estaChequeado = arbolBackups.Nodes.item(indexNodo).checked
End Function


Property Get cantidadChecked(Optional ByVal soloPadres As Boolean = False) As Long
    Dim loopElemento As Integer
    
    cantidadChecked = 0
    
    With arbolBackups.Nodes
        For loopElemento = .count To 1 Step -1
            If .item(loopElemento).checked And Not (soloPadres = True And Not .item(loopElemento).parent Is Nothing) Then
                cantidadChecked = cantidadChecked + 1
            End If
        Next loopElemento
            
    End With

End Property

Private Sub arbolBackups_LostFocus()
    txtBuscador.BackColor = vbWhite
End Sub

Private Sub arbolBackups_NodeClick(ByVal node As MSComctlLib.node)
    RaiseEvent change(Me.obtenerValor, Me.obtenerIDValor, node.Children > 0)
End Sub

Private Sub timerRetardador_Timer()
    Call cargarLista(txtBuscador)
        
    RaiseEvent change(Me.obtenerValor(), Me.obtenerIDValor, left$(Me.obtenerValor, 1) = "P")
    
    timerRetardador.Enabled = False
End Sub

Private Sub txtBuscador_Change()
    
    timerRetardador.Enabled = False 'Lo deshabilito
    timerRetardador.Enabled = True 'Lo habilito
    timerRetardador.Interval = 700


End Sub

Public Sub cambiarNombre(ByVal idNodo As Long, ByVal texto As String)
    Dim nodo As node
    
    Set nodo = buscarNodo(lstListaContenido, idNodo)
    
    If Not nodo Is Nothing Then
        nodo.text = texto
    End If
    
    Set nodo = buscarNodo(arbolBackups, idNodo)
    
    If Not nodo Is Nothing Then
        nodo.text = texto
    End If
    
End Sub

Public Function existe(ByVal idNodo As Long) As Boolean

    Dim nodo As node
    
    Set nodo = buscarNodo(lstListaContenido, idNodo)
        
    If Not nodo Is Nothing Then existe = True Else existe = False
    
End Function

Public Function esPadre(ByVal idNodo As Long) As Boolean
Dim nodo As node

Set nodo = buscarNodo(lstListaContenido, idNodo)

esPadre = (nodo.parent Is Nothing)

End Function
Public Sub eliminarHijos(idNodo As Long)
    Dim nodo As node
    
    Set nodo = buscarNodo(lstListaContenido, idNodo)
    
    If Not nodo Is Nothing Then
        Do While Not nodo.Child Is Nothing
        lstListaContenido.Nodes.Remove (nodo.Child.Index)
        Loop
    End If
    
    Set nodo = buscarNodo(arbolBackups, idNodo)
    
    If Not nodo Is Nothing Then
        Do While Not nodo.Child Is Nothing
            Call arbolBackups.Nodes.Remove(nodo.Child.Index)
        Loop
    End If
    
End Sub

Public Function obtenerCantidadPadres() As Integer
    obtenerCantidadPadres = lstListaContenido.Nodes.count
End Function

Public Function obtenerCantidadVisible() As Integer
    obtenerCantidadVisible = arbolBackups.Nodes.count
End Function

Public Sub eliminarElemento(ByVal idNodo As Long)
    Dim nodo As node
    
    Set nodo = buscarNodo(lstListaContenido, idNodo)
    
    If Not nodo Is Nothing Then
        Call lstListaContenido.Nodes.Remove(nodo.Index)
    End If
    
    Set nodo = buscarNodo(arbolBackups, idNodo)
    
    If Not nodo Is Nothing Then
        Call arbolBackups.Nodes.Remove(nodo.Index)
    End If
End Sub

Private Function buscarNodo(arbol As TreeView, ByVal idNodo As Long) As node
    
    For Each buscarNodo In arbol.Nodes
        If buscarNodo.key = "P" & idNodo Then
            Exit Function
        End If
    Next
    Set buscarNodo = Nothing
End Function
Private Sub borrarLista()
Dim x As Integer

With arbolBackups.Nodes
    For x = .count To 1 Step -1
     .Remove (x)
    Next x
End With

End Sub

Private Sub borrarDatos()
Dim x As Integer

With lstListaContenido.Nodes
    For x = .count To 1 Step -1
     .Remove (x)
    Next x
End With

End Sub

Private Sub cargarLista(Filtro As String)
    Dim cantidadvalidos As Integer
    Dim nodo As node
    Dim hijo As node
    
    lblEstado.caption = "Buscando..."
    lblEstado.FontItalic = True
    
    'No permitimos que se actualice el arbol, si estamos posicionados
    'en más dela primer hoja del arbol y borramos, se laguea
    LockWindowUpdate arbolBackups.hwnd
    
    borrarLista
    
    cantidadvalidos = 0
    
    Filtro = quitarTildes(Filtro)
    
    If lstListaContenido.Nodes.count = 0 Then
        Set nodo = Nothing
    Else
        Set nodo = lstListaContenido.Nodes(1)
    End If
    
    Do While Not nodo Is Nothing

        If InStr(1, quitarTildes(nodo.text), Filtro, vbTextCompare) Then

            arbolBackups.Nodes.Add , , nodo.key, nodo.text
            
            Set hijo = nodo.Child
            
            Do While Not hijo Is Nothing
                arbolBackups.Nodes.Add nodo.key, tvwChild, hijo.key, hijo.text
                Set hijo = hijo.Next
            Loop
                            
            cantidadvalidos = cantidadvalidos + 1
            
        ElseIf nodo.Children > 0 Then
                Set hijo = nodo.Child
                
                Do While Not hijo Is Nothing

                    If InStr(1, quitarTildes(hijo.text), Filtro, vbTextCompare) Then
                        
                            arbolBackups.Nodes.Add , , nodo.key, nodo.text
                            Set hijo = nodo.Child
                            
                            Do While Not hijo Is Nothing
                                 arbolBackups.Nodes.Add nodo.key, tvwChild, hijo.key, hijo.text
                                 Set hijo = hijo.Next
                            Loop
                        Exit Do
                    End If
                    
                    Set hijo = hijo.Next
                Loop
        End If
        
        Set nodo = nodo.Next
    Loop
    
    If cantidadvalidos > 0 Then
        Set arbolBackups.SelectedItem = arbolBackups.Nodes(1)
        lblEstado.caption = ""
    Else
        lblEstado.caption = "Sin resultados"
        lblEstado.FontItalic = False
    End If
    
    
    LockWindowUpdate False
End Sub

Private Function quitarTildes(texto As String) As String
    quitarTildes = Replace$(texto, "á", "a")
    quitarTildes = Replace$(quitarTildes, "é", "e")
    quitarTildes = Replace$(quitarTildes, "í", "i")
    quitarTildes = Replace$(quitarTildes, "ó", "o")
    quitarTildes = Replace$(quitarTildes, "ú", "u")
End Function

Public Sub vaciar()
    borrarLista
    
    borrarDatos
    'Timer retardado
    timerRetardador.Enabled = False
End Sub

Private Sub txtBuscador_GotFocus()
    txtBuscador.BackColor = vbGreen
End Sub

Private Sub txtBuscador_LostFocus()
    txtBuscador.BackColor = vbWhite
End Sub

Private Sub UserControl_Initialize()
    Call vaciar
    lblEstado.caption = ""
End Sub

Public Function obtenerValor() As String
    If Not arbolBackups.SelectedItem Is Nothing Then
        obtenerValor = arbolBackups.SelectedItem.text
    Else
        obtenerValor = ""
    End If

End Function

Public Function obtenerIDValor() As Integer
    If Not arbolBackups.SelectedItem Is Nothing Then
        obtenerIDValor = val(mid(arbolBackups.SelectedItem.key, 2))
    Else
        obtenerIDValor = 0
    End If
End Function

Public Function obtenerIDSeleccionados() As Long()
    Dim aux() As Long
    Dim nodo As node
    Dim loopElemento As Long
    
    ReDim aux(0 To Me.getCantidadSeleccionada - 1) As Long
        
    Set nodo = arbolBackups.Nodes(1)
    
    loopElemento = 0
    
    Do While Not nodo Is Nothing

        If nodo.checked Then
            aux(loopElemento) = val(mid$(nodo.key, 2))
            loopElemento = loopElemento + 1
        End If
        
        Set nodo = nodo.Next
    Loop
        
    
    obtenerIDSeleccionados = aux
    
End Function
Public Sub addString(ByVal id As Long, ByVal Contenido As String, ByVal idPadre As Long)
    On Error Resume Next
    If idPadre > 0 Then
        Call arbolBackups.Nodes.Add("P" & idPadre, tvwChild, "H" & id & "P" & idPadre, Contenido)
        Call lstListaContenido.Nodes.Add("P" & idPadre, tvwChild, "H" & id & "P" & idPadre, Contenido)
    Else
        Call arbolBackups.Nodes.Add(, , "P" & id, Contenido)
        Call lstListaContenido.Nodes.Add(, , "P" & id, Contenido)
    End If
End Sub

Public Sub deseleccionar()
    arbolBackups.SelectedItem = Nothing
End Sub

Public Function seleccionarElemento(ByVal id As Long) As Boolean
    Dim nodo As node
    Set nodo = buscarNodo(arbolBackups, id)
        
    If Not nodo Is Nothing Then
        Set arbolBackups.SelectedItem = nodo
        RaiseEvent change(Me.obtenerValor, Me.obtenerIDValor, True)
        arbolBackups.SelectedItem.EnsureVisible
        arbolBackups.SetFocus
        
        seleccionarElemento = True
    Else
        seleccionarElemento = False
    End If
End Function

Private Sub UserControl_Resize()
    arbolBackups.width = UserControl.width
    arbolBackups.height = UserControl.height - txtBuscador.height - lblEstado.height
    txtBuscador.width = UserControl.width
End Sub

Public Property Get tag() As String
    tag = tagInfo
End Property

Public Property Let tag(ByVal vNewValue As String)
    tagInfo = vNewValue
End Property

Property Let checked(valor As Boolean)
    lstListaContenido.Checkboxes = valor
    arbolBackups.Checkboxes = valor
End Property

Property Get checked() As Boolean
     checked = lstListaContenido.Checkboxes
End Property


Property Get getCantidadSeleccionada() As Integer

    Dim cantidad As Integer
    Dim nodo As node
    
    cantidad = 0
        
    If arbolBackups.Checkboxes = False Then
        getCantidadSeleccionada = 0
        Exit Function
    End If

    Set nodo = arbolBackups.Nodes(1)
    
    Do While Not nodo Is Nothing

        If nodo.checked Then cantidad = cantidad + 1

        Set nodo = nodo.Next
    Loop
    

    getCantidadSeleccionada = cantidad
End Property

