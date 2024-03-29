VERSION 5.00
Begin VB.UserControl GridTextConAutoCompletar 
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2445
   ScaleWidth      =   4215
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtCampo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   300
         Width           =   1455
      End
      Begin EditorTDS.ListaConBuscador listaConBuscador 
         Height          =   1695
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _extentx        =   2778
         _extenty        =   2990
      End
      Begin VB.Label lblHorizontal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ddddd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   3240
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Se vera de 0 a 100 de vida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label lblCampo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Campo 1"
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      LargeChange     =   50
      Left            =   3960
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Value           =   50
      Width           =   255
   End
End
Attribute VB_Name = "GridTextConAutoCompletar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Textbox que esta actualemnte seleccionado
Private campoSeleccionado As VB.TextBox
Private nombreCampoGenerico As String
Private controlesDinamicos() As Control
Private controlesDinamicosDescripciones() As String

Private tieneDinamicos As Boolean

Public Event CantidadElementoChange()
Public Event ElementoChange(Index As Integer)

Private tagInfo As String

'Eliminamos todos los elementos
Public Sub limpiar()
    Call borrar(txtCampo.count - 1)
End Sub

'Establece el nombre del campo que se muestra en la primera columna
'y va aumentando nombre_ X por cada campo que se agrega
Public Sub setNombreCampos(nombre_ As String)
    nombreCampoGenerico = nombre_
End Sub

'Retorna la cantidad de campos completos que hay
Public Function obtenerCantidadCampos() As Byte
    obtenerCantidadCampos = txtCampo.count - 1 'Menos 1 porque El ultimo esta vacio
End Function

'A�ade un elemento a la lista principal
Public Sub addString(ByVal id As Integer, ByVal Contenido As String)
    Call listaConBuscador.addString(id, Contenido)
End Sub

'Establece la descripcion para un campo numerico
Public Sub setDescripcion(numeroCampo As Byte, descripcion As String)
    lblDescripcion(numeroCampo).caption = descripcion
End Sub

Public Function obtenerID(ByVal elemento As Byte)
    obtenerID = txtCampo(elemento).tag
End Function

'Agrega un control dinamico, en base a un prototipo pasado como parametro, mas un nombre y una descripcion
Public Sub agregarControlDinamico(prototipo As Control, nombre As String, descripcion As String)

    If tieneDinamicos Then
        ReDim Preserve controlesDinamicos(0 To UBound(controlesDinamicos) + 1)
        ReDim Preserve controlesDinamicosDescripciones(0 To UBound(controlesDinamicosDescripciones) + 1)
    Else
        ReDim controlesDinamicos(0)
        ReDim controlesDinamicosDescripciones(0)
    End If

    Set controlesDinamicos(UBound(controlesDinamicos)) = clonarControl(prototipo, nombre & "_" & 0)
    controlesDinamicosDescripciones(UBound(controlesDinamicosDescripciones)) = descripcion
    
    tieneDinamicos = True
End Sub

Public Sub setValorDinamico(nombreControl As String, fila As Byte, valor As String)

    Dim c As Control
    
    Set c = Controls(nombreControl & "_" & fila)

   Call establecerValorAControl(c, valor)
End Sub

Private Function obtenerValorAControl(Control As Control) As String

On Error Resume Next

obtenerValorAControl = Control.text
If Err = 0 Then Exit Function Else Err.Clear

obtenerValorAControl = Control.value
If Err = 0 Then Exit Function Else Err.Clear

obtenerValorAControl = Control.caption

End Function

Private Sub establecerValorAControl(Control As Control, valor As String)

On Error Resume Next

Control.text = valor
If Err = 0 Then Exit Sub Else Err.Clear

Control.value = valor
If Err = 0 Then Exit Sub Else Err.Clear

Control.caption = valor
End Sub

'Obtiene el vaalor para del control dinamico nombreControl en determinada fila
Public Function getValorDinamico(nombreControl As String, fila As Byte) As String
    Dim c As Control
    
    Set c = Controls(nombreControl & "_" & fila)
    
    getValorDinamico = obtenerValorAControl(c)
End Function


'Propiedades tipicas de un control
Public Property Get Enabled() As Boolean
   Enabled = txtCampo(0).Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Dim Control As Control
    
    For Each Control In Controls
        Control.Enabled = vNewValue
    Next
End Property


'Borramos CANTIDAD de elementos, empezando desde el ultimo
Private Sub borrar(cantidad As Byte)
    Dim i As Byte
    Dim loopDinamico As Byte
    
    For i = 1 To cantidad
        Unload txtCampo(txtCampo.UBound)
        Unload lblCampo(lblCampo.UBound)
        Unload lblDescripcion(lblDescripcion.UBound)
        
        If tieneDinamicos Then
            For loopDinamico = LBound(controlesDinamicos) To UBound(controlesDinamicos)
                Call Controls.Remove(mid$(controlesDinamicos(loopDinamico).Name, 1, InStr(1, controlesDinamicos(loopDinamico).Name, "_")) & (txtCampo.UBound + 1))
           Next
        End If
        
    Next
    
    lblDescripcion(lblDescripcion.UBound).caption = ""

    Call Me.seleccionarID(lblCampo.UBound, -1)
End Sub
Private Sub redimensionarControles()
Dim i As Integer
Dim cantidad As Byte
Dim totalCampos As Byte

totalCampos = txtCampo.UBound - txtCampo.LBound + 1

'Cuanto la cantidad con ID  = 0
For i = txtCampo.UBound To txtCampo.LBound Step -1
    If val(txtCampo(i).tag) = -1 Then
        cantidad = cantidad + 1
    Else
        Exit For
    End If
Next i

'Si estan todos ocupados, creo uno nuevo.
If cantidad = 0 Then
       
       'Cargo el textbox y el label
        load txtCampo(totalCampos)
       
        With txtCampo(totalCampos)
            .Visible = True
            .left = txtCampo(totalCampos - 1).left
            .top = txtCampo(totalCampos - 1).top + txtCampo(0).height + lblDescripcion(0).height + 50
            .tag = -1
            .text = ""
        End With
       
        load lblCampo(totalCampos)
       
        With lblCampo(totalCampos)
            .caption = nombreCampoGenerico & " " & (totalCampos + 1)
            .left = lblCampo(totalCampos - 1).left
            .top = txtCampo(totalCampos).top + 10
            .Visible = True
        End With
       
        load lblDescripcion(totalCampos)
       
        With lblDescripcion(totalCampos)
            .caption = ""
            .left = lblDescripcion(totalCampos - 1).left
            .top = txtCampo(totalCampos).top - lblDescripcion(0).height - 1
            .Visible = True
        End With
       
        If tieneDinamicos Then
            Dim c As Control
            Dim loopDinamico As Byte
            Dim acumu As Integer
            Dim C2 As Control
            For loopDinamico = LBound(controlesDinamicos) To UBound(controlesDinamicos)
                Set c = clonarControl(controlesDinamicos(loopDinamico), mid$(controlesDinamicos(loopDinamico).Name, 1, InStr(1, controlesDinamicos(loopDinamico).Name, "_")) & totalCampos)
                Set C2 = Controls(mid$(controlesDinamicos(loopDinamico).Name, 1, InStr(1, controlesDinamicos(loopDinamico).Name, "_")) & totalCampos - 1)
                
                c.top = txtCampo(totalCampos).top
                c.left = C2.left
            Next
            
        End If
        
       'Aumento la cantidad de campos que estoy visualizando
       totalCampos = totalCampos + 1
ElseIf cantidad > 1 Then

    'Tengo m�s de uno libre en el final, tengo que eliminar los sombrantes
    Call borrar(cantidad - 1)
    
    'Actulizo la cantidad de campos que tenia
    totalCampos = totalCampos - (cantidad - 1)
Else
    Exit Sub
End If


' Actualizo el tama�o del frame
'El nuevo largo sera la cantidad de campos que estoy visualizando y para que me entre toda la lista
Frame1.height = (totalCampos * (txtCampo(0).height + lblDescripcion(0).height + 50)) + (listaConBuscador.height - txtCampo(0).height + 200)

' Actualizo el valor de la barra
Frame1.top = (Frame1.height - UserControl.height) * (VScroll1.value / 100) * -1

If Frame1.height > UserControl.height Then
    VScroll1.value = -100 * (Frame1.top / (Frame1.height - UserControl.height))
Else
    VScroll1.value = 0
End If

RaiseEvent CantidadElementoChange
End Sub
Private Sub OcultarLista()
    If Not campoSeleccionado Is Nothing Then
        listaConBuscador.Visible = False
        campoSeleccionado.Visible = True
        
        Set campoSeleccionado = Nothing
    End If
End Sub





Public Sub seleccionarID(ByVal elemento As Byte, ByVal id As Long)

    Set campoSeleccionado = txtCampo(elemento)
    
    If listaConBuscador.seleccionarID(CInt(id)) Then
        campoSeleccionado.tag = listaConBuscador.obtenerIDValor
        campoSeleccionado.text = listaConBuscador.obtenerValor
        
        Call redimensionarControles
    End If

End Sub

Public Sub iniciar()
    Dim anchoDinamicos As Integer
    Dim loopDinamico As Byte
    lblCampo(0).caption = nombreCampoGenerico & " 1"
    
    Call listaConBuscador.addString(-1, "")
    
    anchoDinamicos = 0

    If tieneDinamicos Then
         For loopDinamico = LBound(controlesDinamicos) To UBound(controlesDinamicos)
            anchoDinamicos = anchoDinamicos + controlesDinamicos(loopDinamico).width + 100
        Next
    End If
    
    lblCampo(0).width = Len(lblCampo(0).caption) * 80
    
    txtCampo(0).left = lblCampo(0).left + lblCampo(0).width + 100
    
    txtCampo(0).width = (Frame1.width - 50) - txtCampo(0).left - anchoDinamicos
    
    If tieneDinamicos Then
    
        anchoDinamicos = txtCampo(0).left + txtCampo(0).width + 100
        
        For loopDinamico = LBound(controlesDinamicos) To UBound(controlesDinamicos)
    
            If loopDinamico > 0 Then load lblHorizontal(loopDinamico)
            
            controlesDinamicos(loopDinamico).left = anchoDinamicos
            
            lblHorizontal(loopDinamico).caption = controlesDinamicosDescripciones(loopDinamico)
            lblHorizontal(loopDinamico).left = controlesDinamicos(loopDinamico).left
            lblHorizontal(loopDinamico).Visible = True
            
            anchoDinamicos = anchoDinamicos + controlesDinamicos(loopDinamico).width + 100
            
        Next

    End If
    
    
    listaConBuscador.left = txtCampo(0).left
    listaConBuscador.width = txtCampo(0).width
End Sub

Private Sub MostrarEn(text As TextBox)

    If Not campoSeleccionado Is Nothing Then
        campoSeleccionado.Visible = True
    End If
    
    Set campoSeleccionado = text
    
    'Posiciono la lista con buscador
    listaConBuscador.top = text.top
    listaConBuscador.left = text.left
    listaConBuscador.Visible = True
    campoSeleccionado.Visible = False
    'Establece el foco en la lista
    listaConBuscador.SetFocus
End Sub


Private Sub ListaConBuscador_DblClic()
    
    If Not campoSeleccionado Is Nothing Then
        
        If val(campoSeleccionado.tag) <> listaConBuscador.obtenerIDValor Then
            
            If Not (listaConBuscador.obtenerIDValor = -1 And campoSeleccionado.Index = 0) Then
                campoSeleccionado.text = listaConBuscador.obtenerValor
                campoSeleccionado.tag = listaConBuscador.obtenerIDValor
            Else
                'Error
            End If
        End If
    End If
    
    Call OcultarLista
    
    Call redimensionarControles
End Sub

Private Sub ListaConBuscador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call OcultarLista
    End If
End Sub

Private Sub ListaConBuscador_LostFocus()
    Call OcultarLista
End Sub

Private Sub txtCampo_Change(Index As Integer)
   RaiseEvent ElementoChange(Index)
End Sub

Private Sub txtCampo_GotFocus(Index As Integer)
    Call MostrarEn(txtCampo(Index))
End Sub

Private Sub UserControl_Resize()
    Dim i As Byte

    VScroll1.left = UserControl.width - VScroll1.width
    VScroll1.height = UserControl.height
    
    Frame1.width = VScroll1.left - Frame1.left - 1
    
    For i = txtCampo.LBound To txtCampo.UBound
        txtCampo(i).width = Frame1.width - txtCampo(i).left - 50
    Next i
    
    listaConBuscador.left = txtCampo(0).left
    listaConBuscador.width = txtCampo(0).width
End Sub

Private Sub actualizarVisibilidad()
    '�Es mas grande de lo que puedo ver?
    If (Frame1.height - UserControl.height) > 0 Then
        'El top va a estar entre 0 y (la altura del frame - lo altura del control, que es la parte que ve el usuario)
        'Depende el porcentaje es donde se ubica en ese intervalo
        Frame1.top = (Frame1.height - UserControl.height) * (VScroll1.value / 100) * -1
    Else
        Frame1.top = 0
    End If
End Sub
Private Sub VScroll1_Change()
    Call actualizarVisibilidad
End Sub

Private Sub VScroll1_Scroll()
    Call actualizarVisibilidad
End Sub

Private Function clonarControl(Control As Control, nombre As String) As Control
    Dim Tipo As String
    Dim c As Control
    
    Tipo = TypeName(Control)

    On Error Resume Next
    Set c = Controls.Add("VB." & Tipo, nombre)
    If Err > 0 Then
        Err.Clear
        Set c = Controls.Add("EditorTDS." & Tipo, nombre)
    End If
    
    On Error Resume Next
    'Copio la apariencia
    c.width = Control.width
    c.height = Control.height
    c.Appearance = Control.Appearance
    c.BackColor = Control.BackColor
    c.ToolTipText = Control.ToolTipText
    'Posicion
    c.Visible = True
    c.top = txtCampo(0).top
    
    Set c.Container = Frame1
End Function

Public Property Get tag() As String
    tag = tagInfo
End Property

Public Property Let tag(ByVal vNewValue As String)
    tagInfo = vNewValue
End Property

