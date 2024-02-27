VERSION 5.00
Begin VB.Form frmMisc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tierras del Sur - Más opciones"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMisc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmAspecto 
      Caption         =   "Aspecto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   4080
      TabIndex        =   21
      Top             =   3240
      Width           =   4215
      Begin VB.ComboBox cmbAspecto 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton cmdAspecto_Aplicar 
         Caption         =   "Aplicar aspecto"
         Height          =   360
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.Frame frmRemplazarPiso 
      Caption         =   "Remplazar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   3855
      Begin EditorTDS.TextConListaConBuscador txtRemplazarPiso_Viejo 
         Height          =   285
         Left            =   480
         TabIndex        =   14
         Top             =   315
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         CantidadLineasAMostrar=   6
      End
      Begin EditorTDS.TextConListaConBuscador txtRemplazarPiso_Nuevo 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         CantidadLineasAMostrar=   4
      End
      Begin VB.ComboBox cmbRemplazarPiso_Nuevo_Sec 
         Height          =   315
         ItemData        =   "frmMisc.frx":1CCA
         Left            =   960
         List            =   "frmMisc.frx":1CCC
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cmbRemplazarPiso_Viejo_Sec 
         Height          =   315
         ItemData        =   "frmMisc.frx":1CCE
         Left            =   960
         List            =   "frmMisc.frx":1CD0
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   630
         Width           =   2775
      End
      Begin VB.CommandButton cmdRemplazarPiso 
         Caption         =   "Remplazar"
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   3615
      End
      Begin VB.CommandButton txtRemplazarPiso_Recargar 
         Height          =   285
         Left            =   3480
         Picture         =   "frmMisc.frx":1CD2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Recargar lista de pisos"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label lblRemplazarPiso_Por 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "por"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame frmPersonaje 
      Caption         =   "Personaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   4215
      Begin VB.Timer trmMoverme 
         Enabled         =   0   'False
         Left            =   360
         Top             =   2280
      End
      Begin VB.CommandButton cmdCaminar 
         Caption         =   "Caminar Automatico"
         Height          =   480
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   3855
      End
      Begin VB.HScrollBar scrlVelocidadDelPersonaje 
         Height          =   255
         Left            =   120
         Max             =   16
         Min             =   1
         TabIndex        =   10
         Top             =   840
         Value           =   8
         Width           =   2175
      End
      Begin VB.Label lblVelocidadPersonaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad del personaje: 200,00 milisegundos por paso"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3930
      End
   End
   Begin VB.Frame frmRemplazarGraficos 
      Caption         =   "Remplazar Gráficos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "?"
         Height          =   360
         Left            =   3360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   120
         Width           =   390
      End
      Begin VB.TextBox txtRemplazarGraficos_Remplazar 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtRemplazarGraficos_Buscar 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdRemplazarGraficos 
         Caption         =   "Remplazar"
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   3615
      End
      Begin VB.ComboBox cmbRemplazarGraficos_Capa 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblRemplazarGraficos_Remplazar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y cambiarlo por el gráfico que tiene de nombre el nombre original remplazando lo buscador por:"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   3645
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRemplazarGraficos_Buscar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar gráficos que coincidan con"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblRemplazarGraficos_Capa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "En la capa:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAspecto_Aplicar_Click()
    Dim idPixel As Byte
    
    idPixel = val(Me.cmbAspecto.list(Me.cmbAspecto.listIndex))
    
    PixelShaderBump = CreateShaderFromCode(Replace$(PixelShaders(idPixel).codigo, ";", vbNewLine))
    
    If PixelShaderBump = 0 Then
        Call GUI_Alert("Falló al aplicar el aspecto. Es probable que este mal el código del aspecto.", "Aspecto")
    End If
End Sub

Private Sub cmdCaminar_Click()
    
    If Me.trmMoverme.Enabled Then
        Me.trmMoverme.Enabled = False
        Exit Sub
    End If
    
    Me.trmMoverme.Interval = (8 * 200 / Me.scrlVelocidadDelPersonaje.value) * 0.9
    Me.trmMoverme.Enabled = True
    
End Sub

Private Sub cmdRemplazarGraficos_Click()
    Dim respuesta As VbMsgBoxResult
    Dim buscar As String
    Dim remplazar As String
    Dim Capa As Integer
    Dim comando As cComandoRemplazarGrafico
    
    
    If Me.cmbRemplazarGraficos_Capa.listIndex = -1 Then
        MsgBox "Tenes que seleccionar la Capa en la cual se va a aplicar el remplazo de gráficos.", vbExclamation, Me.caption
        Exit Sub
    End If
    
    buscar = Me.txtRemplazarGraficos_Buscar
    remplazar = Me.txtRemplazarGraficos_Remplazar
    
    If Len(Trim$(buscar)) = 0 Or Len(Trim$(remplazar)) = 0 Then
        MsgBox "Tenes que escribir lo que queres remplazar y por que lo queres remplazar.", vbExclamation, Me.caption
        Exit Sub
    End If

    Capa = Me.cmbRemplazarGraficos_Capa.itemData(Me.cmbRemplazarGraficos_Capa.listIndex)
    
    respuesta = MsgBox("Se buscarán todos los gráficos en la capa " & Capa & " que tengan en su nombre la cadena '" & buscar & "' y lo remplazará por el gráfico que tenga el mismo nombre pero remplazando '" & buscar & "' por '" & remplazar & "'." & vbNewLine & vbNewLine & " Por ejemplo donde este el gráfico '" & buscar & "Puerta cerrada' se cambiará por el gráfico que tiene el nombre '" & remplazar & "Puerta cerrada'" & vbNewLine & vbNewLine & "¿Estás seguro? Si te equivocas podes deshacerlo con CONTORL + Z.", vbQuestion + vbYesNo)
    
    If respuesta = vbYes Then
        
        Set comando = New cComandoRemplazarGrafico
    
        Call comando.crear(buscar, remplazar, Capa)
        
        Call ME_Tools.ejecutarComando(comando)
    End If
    
End Sub

Private Sub cmdRemplazarPiso_Click()
    Dim comando As cComandoRemplazarTileSet
    Dim respuesta As VbMsgBoxResult
    Dim viejo As Integer
    Dim nuevo As Integer
    Dim comienzoNueva As Integer
    Dim comienzoVieja As Integer
    
    
    If Me.cmbRemplazarPiso_Viejo_Sec.listIndex = -1 Then
        MsgBox "Tenes que informar la parte (o todo) del piso que estas buscando para remplazar.", vbExclamation
        Exit Sub
    End If
    
    If Me.cmbRemplazarPiso_Nuevo_Sec.listIndex = -1 Then
        MsgBox "Tenes que seleccionar que parte (o todo) del piso que vas a poner por el que buscas.", vbExclamation, Me.caption
        Exit Sub
    End If
    
    viejo = Me.txtRemplazarPiso_Viejo.obtenerIDValor
    nuevo = Me.txtRemplazarPiso_Nuevo.obtenerIDValor
    
    If viejo = 0 Then
        MsgBox "Tenes que seleccionar el piso que queres remplazar.", vbExclamation, Me.caption
        Exit Sub
    End If
    
    If nuevo = 0 Then
        MsgBox "Tenes que seleccionar el piso que queres poner.", vbExclamation, Me.caption
        Exit Sub
    End If
    
    If (comienzoNueva = 0 And comienzoVieja > 0) Or (comienzoNueva > 0 And comienzoVieja = 0) Then
        MsgBox "Esto esta mal. Si seleccionas remplazar una 'parte' de un piso tenes que seleccionar una 'parte' (cualquiera) del otro piso. Y si seleccionas que queres remplazar 'todo' el piso, también tenes que seleccionar 'todo' en el otro piso.", vbExclamation
        Exit Sub
    End If
    
    comienzoNueva = Me.cmbRemplazarPiso_Nuevo_Sec.itemData(Me.cmbRemplazarPiso_Nuevo_Sec.listIndex)
    comienzoVieja = Me.cmbRemplazarPiso_Viejo_Sec.itemData(Me.cmbRemplazarPiso_Viejo_Sec.listIndex)
    
    respuesta = MsgBox("Estás seguro que queres remplazar, en el mapa, el piso '" & Me.txtRemplazarPiso_Viejo.obtenerValor & "' por '" & Me.txtRemplazarPiso_Nuevo.obtenerValor & "'.", vbQuestion + vbYesNo)
    
    If respuesta = vbYes Then
        
        Set comando = New cComandoRemplazarTileSet
    
        Call comando.crear(viejo, comienzoVieja, nuevo, comienzoNueva)
        Call ME_Tools.ejecutarComando(comando)
    End If
End Sub

Private Sub Form_Load()

    Dim loopCapa As Byte
    Dim loopPixel As Integer
    
    ' Remplazar graficos
    Me.cmbRemplazarGraficos_Capa.Clear
    
    For loopCapa = 1 To CANTIDAD_CAPAS
        Me.cmbRemplazarGraficos_Capa.AddItem (loopCapa)
        Me.cmbRemplazarGraficos_Capa.itemData(Me.cmbRemplazarGraficos_Capa.NewIndex) = loopCapa
    Next
    
    ' Cambiar la velocidad del personaje
    Me.scrlVelocidadDelPersonaje.value = CharList(UserCharIndex).Velocidad.x
    
    Call actualizarCaptionVelocidadPersonaje
    
    ' Remplazar pisos
   Call cargarPisosRemplazarPisos
    
    
    ' Pixel Shaders
    Me.cmbAspecto.Clear
    
    For loopPixel = LBound(PixelShaders) To UBound(PixelShaders)
        If Not PixelShaders(loopPixel).codigo = "" Then
            Me.cmbAspecto.AddItem (loopPixel & " - " & PixelShaders(loopPixel).nombre)
        End If
    Next
End Sub


Private Sub actualizarCaptionVelocidadPersonaje()
    Me.lblVelocidadPersonaje.caption = "Velocidad del personaje: " & FormatNumber(200 * 8 / Me.scrlVelocidadDelPersonaje.value, 2) & " milisegundos por paso."
End Sub

Private Sub scrlVelocidadDelPersonaje_Change()
   
    Call actualizarCaptionVelocidadPersonaje

    CharList(UserCharIndex).Velocidad.x = Me.scrlVelocidadDelPersonaje.value
    CharList(UserCharIndex).Velocidad.y = Me.scrlVelocidadDelPersonaje.value
End Sub

Private Sub trmMoverme_Timer()

    UserDirection = CharList(UserCharIndex).heading

End Sub

Private Sub txtRemplazarPiso_Recargar_Click()
     Call cargarPisosRemplazarPisos
End Sub

Public Sub cargarPisosRemplazarPisos()

    Dim loopPiso As Integer
    Dim loopSeccion As Integer
    
    ' Limpiamos
    Me.txtRemplazarPiso_Nuevo.limpiarLista
    Me.txtRemplazarPiso_Viejo.limpiarLista
    Me.cmbRemplazarPiso_Nuevo_Sec.Clear
    Me.cmbRemplazarPiso_Viejo_Sec.Clear
    
    ' Cargamos las texturas disponibles
    For loopPiso = 1 To Tilesets_count
        If Me_indexar_Pisos.existe(loopPiso) Then
            Call Me.txtRemplazarPiso_Nuevo.addString(loopPiso, loopPiso & " - " & Engine_Tilesets.Tilesets(loopPiso).nombre)
            Call Me.txtRemplazarPiso_Viejo.addString(loopPiso, loopPiso & " - " & Engine_Tilesets.Tilesets(loopPiso).nombre)
        End If
    Next
    
    ' Cargamos las secciones
    With Me.cmbRemplazarPiso_Nuevo_Sec
    
        .AddItem ("Todo")
        .itemData(.NewIndex) = -1
        
        .AddItem ("1 - Superior izquierdo")
        .itemData(.NewIndex) = 0
        
        .AddItem ("2 - Superior derecho")
        .itemData(.NewIndex) = 8
    
        .AddItem ("3 - Inferior derecho")
        .itemData(.NewIndex) = 136
    
        .AddItem ("4 - Inferior izquierdo")
        .itemData(.NewIndex) = 128
        
    End With
    
    For loopSeccion = 0 To 4
        Me.cmbRemplazarPiso_Viejo_Sec.AddItem (Me.cmbRemplazarPiso_Nuevo_Sec.list(loopSeccion))
        Me.cmbRemplazarPiso_Viejo_Sec.itemData(loopSeccion) = Me.cmbRemplazarPiso_Nuevo_Sec.itemData(loopSeccion)
    Next
    
End Sub
