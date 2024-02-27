VERSION 5.00
Begin VB.Form frmConfigurarPisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Pisos"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12450
   Icon            =   "frmTilesets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   12450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmPropiedades 
      Caption         =   "Propiedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   8655
      Begin VB.Frame frmConfigTipo 
         Caption         =   "Configuración del tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4200
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   4410
         Begin VB.Frame frmConfigTipo_CostaParte2 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   4215
            Begin VB.ComboBox cmbConfigTipo_CostaParte1 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   240
               Width           =   3255
            End
            Begin VB.Label lblConfigTipo_CostaParte1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1ª Parte"
               Height          =   195
               Left            =   0
               TabIndex        =   35
               Top             =   300
               Width           =   570
            End
         End
         Begin VB.Frame frmConfigTipo_CostaParte1 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   4215
            Begin VB.ComboBox cmbConfigTipo_Agua 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   240
               Width           =   3375
            End
            Begin VB.Label lblAgua 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Agua:"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   300
               Width           =   420
            End
         End
         Begin VB.Frame frmConfigTipo_CaminoGrandeParte2 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   3975
            Begin VB.ComboBox cmbConfigTipo_CaminoParte1 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   180
               Width           =   3255
            End
            Begin VB.Label lblConfigTipo_CaminoParte1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1ª Parte:"
               Height          =   195
               Left            =   0
               TabIndex        =   29
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame frmConfigTipo_CaminoChicoGrandeP1 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   4095
            Begin VB.ComboBox cmbConfigTipo_TexturaSimple 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   0
               Width           =   3255
            End
            Begin VB.ComboBox cmbConfigTipo_Sector 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   360
               Width           =   3255
            End
            Begin VB.Label lblConfigTipo_TexturaSimple 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Textura:"
               Height          =   195
               Left            =   0
               TabIndex        =   26
               Top             =   100
               Width           =   585
            End
            Begin VB.Label lblConfigTipo_TexturaSector 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sector:"
               Height          =   195
               Left            =   0
               TabIndex        =   25
               Top             =   400
               Width           =   510
            End
         End
      End
      Begin VB.CommandButton cmdConfigurarEfectosPisadas 
         Caption         =   "Configurar Efectos de Sonido al Pisar"
         Height          =   360
         Left            =   5040
         TabIndex        =   20
         Top             =   2760
         Width           =   3495
      End
      Begin VB.ComboBox cmbTipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   3135
      End
      Begin EditorTDS.TextConListaConBuscador txtOlitas 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
      End
      Begin VB.TextBox nombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         ToolTipText     =   "Alto de Pixeles"
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox Index 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Nos indica el numero de GrhIndex"
         Top             =   480
         Width           =   735
      End
      Begin EditorTDS.UpDownText Speed 
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         MaxValue        =   10000
         MinValue        =   0
         Enabled         =   -1  'True
      End
      Begin EditorTDS.GridTextConAutoCompletar frames 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4260
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblImagen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imágen/es"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   1920
         Width           =   1740
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblMilisegundos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "milisegundos"
         Height          =   195
         Left            =   7560
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         ToolTipText     =   "1000 milisegundos = 1 segundo"
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lblOlitas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imágen de Olitas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         ToolTipText     =   "Número del recurso de imágen que corresponde a las olas que se verán en caso de que este suelo sea utilizado como aguatierra"
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         ToolTipText     =   "Si el piso tiene animación aquí se debe poner la velocidad entre cada imágen indicadas previamente"
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   9360
      TabIndex        =   5
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdEliminar_Pisos 
      Caption         =   "Eliminar"
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin EditorTDS.ListaConBuscador Listado 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8916
   End
   Begin VB.CommandButton cmdAplicar_Pisos 
      Caption         =   "Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdRestablecer_Pisos 
      Caption         =   "Reestablecer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdNuevo_Pisos 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "frmConfigurarPisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cambiosPendientes As Boolean

Private Sub configTipo_CargarSectoresTextura()
    Dim actual As Integer

    If Me.cmbConfigTipo_Sector.listIndex > -1 Then
        actual = Me.cmbConfigTipo_Sector.itemData(Me.cmbConfigTipo_Sector.listIndex)
    Else
        actual = -1
    End If
    
    Me.cmbConfigTipo_Sector.Clear
    
    Call Me.cmbConfigTipo_Sector.AddItem("1 - Superior izquierdo")
    Me.cmbConfigTipo_Sector.itemData(Me.cmbConfigTipo_Sector.NewIndex) = 0 ' Número de tile donde comienza
    
    Call Me.cmbConfigTipo_Sector.AddItem("2 - Superior derecho")
    Me.cmbConfigTipo_Sector.itemData(Me.cmbConfigTipo_Sector.NewIndex) = 8
    
    Call Me.cmbConfigTipo_Sector.AddItem("3 - Inferior derecho")
    Me.cmbConfigTipo_Sector.itemData(Me.cmbConfigTipo_Sector.NewIndex) = 136
    
    Call Me.cmbConfigTipo_Sector.AddItem("4 - Inferior izquierdo")
    Me.cmbConfigTipo_Sector.itemData(Me.cmbConfigTipo_Sector.NewIndex) = 128
    
    If actual > -1 Then Call seleccionarComboID(Me.cmbConfigTipo_Sector, actual)

End Sub
Private Sub configTipo_CargarTexturas(combo As ComboBox, Tipo As eFormatoTileSet)
    Dim loopPiso As Integer
    Dim actual As Integer

    If combo.listIndex > -1 Then
        actual = combo.itemData(combo.listIndex)
    Else
        actual = -1
    End If
    
    ' Limpiamos
    combo.Clear
    
    For loopPiso = 1 To Tilesets_count
        With Tilesets(loopPiso)
        
            If Me_indexar_Pisos.existe(loopPiso) Then
                
                If Tilesets(loopPiso).formato = Tipo Then
                    Call combo.AddItem(loopPiso & " - " & .nombre)
                    combo.itemData(combo.NewIndex) = loopPiso
                End If
                
            End If
        
        End With
    Next loopPiso

    If actual > -1 Then Call seleccionarComboID(combo, actual)
End Sub

Private Sub mostrarConfiguracionFormato(formato As eFormatoTileSet)

    ' Oculto todo
    Me.frmConfigTipo_CaminoChicoGrandeP1.Visible = False
    Me.frmConfigTipo_CaminoGrandeParte2.Visible = False
    Me.frmConfigTipo_CostaParte1.Visible = False
    Me.frmConfigTipo_CostaParte2.Visible = False

    Me.frmConfigTipo.Visible = False
        
    If formato = eFormatoTileSet.camino_chico Or formato = eFormatoTileSet.camino_grande_parte1 Then
        ' Muestro la opcion
        Me.frmConfigTipo.Visible = True
        Me.frmConfigTipo_CaminoChicoGrandeP1.Visible = True
            
        ' Reseteamos los campos
        Me.cmbConfigTipo_Sector.listIndex = -1
        Me.cmbConfigTipo_TexturaSimple.listIndex = -1
    ElseIf formato = eFormatoTileSet.camino_grande_parte2 Then
        ' Muestro la opción
        Me.frmConfigTipo.Visible = True
        Me.frmConfigTipo_CaminoGrandeParte2.Visible = True
            
        Me.cmbConfigTipo_CaminoParte1.listIndex = -1
    ElseIf formato = eFormatoTileSet.costa_tipo_1_parte1 Then
        
        Me.frmConfigTipo.Visible = True
        Me.frmConfigTipo_CostaParte1.Visible = True
        
        Me.cmbConfigTipo_Agua.listIndex = -1
                
    ElseIf formato = eFormatoTileSet.costa_tipo_1_parte2 Then
    
        
        Me.frmConfigTipo.Visible = True
        Me.frmConfigTipo_CostaParte2.Visible = True
        Me.cmbConfigTipo_CostaParte1.listIndex = -1
    
    End If
End Sub

Private Sub cmbConfigTipo_CaminoParte1_Click()
    Call eb
End Sub

Private Sub cmbConfigTipo_CostaParte1_Change()
    Call eb
End Sub

Private Sub cmbConfigTipo_Sector_Click()
    Call eb
End Sub

Private Sub cmbConfigTipo_TexturaSimple_Click()
    Call eb
End Sub

Private Sub cmbTipo_Click()
   ' Mostramos la pantalla de configuracion
    Call mostrarConfiguracionFormato(Me.cmbTipo.listIndex)
    
    Call eb
End Sub

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdAplicar_Pisos_Click()
    Call Aplicar
    
    tileset_actual = TmpTilesetsNum
    
    Call Me_Tools_TileSet.actualizarListaTileSetDeSeleccion
    
    Call ejecutarControlCambios
End Sub

Private Sub ejecutarControlCambios()
    If cambiosPendientes = False Then
        cambiosPendientes = True
    End If
End Sub

Private Sub cmdConfigurarEfectosPisadas_Click()
    load frmConfigEfectosPisadasEn
    
    frmConfigEfectosPisadasEn.caption = "Configurar Efectos de Pisadas para el piso " & Tilesets(TmpTilesetsNum).nombre & "."
    
    Call frmConfigEfectosPisadasEn.iniciar(16, 16, Tilesets(tileset_actual).EfectoPisada, Me)
    Call frmConfigEfectosPisadasEn.Show
    
    Me.Hide

End Sub

Private Sub cmdEliminar_Pisos_Click()

Dim confirma As VbMsgBoxResult
Dim idPiso As Integer

If Not Me.Listado.obtenerValor() = "" Then
    confirma = MsgBox("¿Está seguro de que desea eliminar el piso '" & Me.Listado.obtenerValor & "'?", vbYesNo + vbExclamation, Me.caption)
    
    If confirma = vbYes Then
    
        tileset_actual = 0
        
        idPiso = Me.Listado.obtenerIDValor
                
        'Guardo en el archivo
        Call Me_indexar_Pisos.eliminar(idPiso)
                
        ' Elimino de la lista
        Call Me.Listado.eliminar(idPiso)
        
        Call Me_Tools_TileSet.actualizarListaTileSetDeSeleccion
        
        Call ejecutarControlCambios
    End If
End If


End Sub

Private Sub Form_Load()
    ' Permisos
    If Not cerebro.Usuario.tienePermisos("CONFIG.PISOS", ePermisosCDM.lectura) Then End
    '
    Dim i As Integer
    
    If ME_ControlCambios.hayCambiosSinActualizarDe("Pisos") Then
        cambiosPendientes = True
    Else
        cambiosPendientes = False
    End If
    
    MOSTRAR_TILESET = True
    
    Listado.vaciar
    For i = 1 To Tilesets_count
        With Tilesets(i)
            If .anim > 0 Then
                Listado.addString i, i & " (ani) - " & .nombre
            Else
                If .stage_count Then
                    Listado.addString i, i & " - " & .nombre
                End If
            End If
        End With
    Next i

    Dim elementos() As modEnumerandosDinamicos.eEnumerado
    
    ' Imagenes que se muestra como ola cuando el personaje se mueve por un pozo de la textura
    elementos = modEnumerandosDinamicos.obtenerEnumeradosDinamicos("IMAGENES")
        
    Call Me.frames.limpiar

    For i = LBound(elementos) To UBound(elementos)
        Call Me.txtOlitas.addString(elementos(i).valor, elementos(i).valor & " - " & elementos(i).nombre)
    Next
    
    ' Imagenes posibles para una animacion
    For i = 1 To UBound(elementos)
         Call Me.frames.addString(elementos(i).valor, elementos(i).valor & " - " & elementos(i).nombre)
    Next
    Call Me.frames.addString(-1, "")
       
    Call Me.frames.setDescripcion(0, "")
    Call Me.frames.setNombreCampos("")
    
    Call Me.frames.iniciar

    ' Cargo tipo
    Call cargarTipoDePisos
    
    Call configTipo_CargarSectoresTextura
    
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_TexturaSimple, eFormatoTileSet.textura_simple)
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_CaminoParte1, eFormatoTileSet.camino_grande_parte1)
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_CostaParte1, eFormatoTileSet.costa_tipo_1_parte1)
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_Agua, eFormatoTileSet.textura_agua)
        
    check_enabled_guardar
End Sub

Private Sub cargarTipoDePisos()
    Call Me.cmbTipo.AddItem("Formato descontracturado")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.formato_viejo
     
    Call Me.cmbTipo.AddItem("Textura simple")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.textura_simple
    
    Call Me.cmbTipo.AddItem("Caminos chicos")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.camino_chico
    
    Call Me.cmbTipo.AddItem("Caminos grandes parte 1")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.camino_grande_parte1
    
    Call Me.cmbTipo.AddItem("Caminos grandes parte 2")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.camino_grande_parte2
    
    Call Me.cmbTipo.AddItem("Agua/Mar")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.textura_agua
    
    Call Me.cmbTipo.AddItem("Costas Parte 1")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.costa_tipo_1_parte1
    
    Call Me.cmbTipo.AddItem("Costas Parte 2")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.costa_tipo_1_parte2
    
    Call Me.cmbTipo.AddItem("Rocas acuaticas")
    Me.cmbTipo.itemData(Me.cmbTipo.NewIndex) = eFormatoTileSet.rocas_acuaticas
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MOSTRAR_TILESET = False
        
    If cambiosPendientes Then
        Call ME_ControlCambios.SetHayCambiosSinActualiar("Pisos")
    End If
End Sub

Private Sub Frames_Change()
eb
End Sub


Sub check_enabled_guardar()

End Sub

Sub eb()
    cmdAplicar_Pisos.Enabled = True
    cmdRestablecer_Pisos.Enabled = True
End Sub

Public Sub hola()
    MsgBox "hola"
End Sub
Private Sub frames_CantidadElementoChange()
    If Me.frames.obtenerCantidadCampos = 1 Then
        Me.Speed.Enabled = False
        Me.Speed.MinValue = 0
        Me.Speed.MaxValue = 0
    Else
        Me.Speed.MaxValue = 5000
        Me.Speed.MinValue = 400
        Me.Speed.Enabled = True
    End If
    
    eb
End Sub

Private Sub frames_ElementoChange(Index As Integer)
    eb
End Sub

Private Sub lblPieIzquierda_Click()
End Sub

Private Sub Listado_Change(valor As String, ID As Integer)
    TmpTilesetsNum = ID
    tileset_actual = TmpTilesetsNum
    reestablece
End Sub

Private Sub cmdNuevo_Pisos_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    error = False
    Me.cmdNuevo_Pisos.Enabled = False
    
    nuevo = Me_indexar_Pisos.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar un nuevo piso. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        
        If Listado.seleccionarID(nuevo) = False Then
            Call Listado.addString(nuevo, nuevo & " - (LIBRE)")
        End If
        
        Call Listado.seleccionarID(nuevo)
                
    End If
    
    Me.cmdNuevo_Pisos.Enabled = True
End Sub

Private Sub nombre_Change()
eb
End Sub


Private Sub cmdRestablecer_Pisos_Click()
reestablece
End Sub



Sub Aplicar()
On Error GoTo errorH
    Dim loopFrame As Byte
    Dim tf() As Long    'TempFrames
    Dim Tileset_BACKUP As TilesetStruct
    Dim sectorTextura As Byte
    
    Dim Velocidad As Single
    Dim CantidadFrames As Byte
    
    
    Tileset_BACKUP = Tilesets(TmpTilesetsNum)
    
    CantidadFrames = Me.frames.obtenerCantidadCampos
    Velocidad = CInt(Speed.value)
    
    If CantidadFrames = 0 Then
        MsgBox "Tenes que seleccionar al menos una imágen para el piso. ", vbExclamation, Me.caption
        Exit Sub
    End If
    
    If Velocidad = 0 And CantidadFrames > 1 Then
        MsgBox "Si el piso es animado, la velocidad de la animación no puede ser 0 milisegundos.", vbExclamation, Me.caption
        Exit Sub
    End If
                
       
    ' Obtengo los frames
    ReDim Preserve tf(1 To CantidadFrames)
        
    For loopFrame = 1 To CantidadFrames
        tf(loopFrame) = Me.frames.obtenerID(loopFrame - 1)
    Next loopFrame
    
    With Tilesets(TmpTilesetsNum)
        
        'Asigno los frames
        .stage_count = CantidadFrames
        
        ReDim .stages(1 To CantidadFrames)
        
        For loopFrame = 1 To .stage_count
            .stages(loopFrame) = tf(loopFrame)
        Next loopFrame
        
        ' Imagen principal
        .filenum = .stages(1)
        .nombre = nombre.text
        
        If CantidadFrames = 1 Then
            .anim = 0
        Else
            .anim = Velocidad
        End If
        
        .Olitas = txtOlitas.obtenerIDValor
        
        ' Especiales para el MapEditor. Nuevo formato de pisos
        .formato = Me.cmbTipo.itemData(Me.cmbTipo.listIndex)
        
        If .formato = eFormatoTileSet.textura_simple Or .formato = formato_viejo Or .formato = textura_agua Or .formato = rocas_acuaticas Then
            ' No guarda referencia a otro
            .referencia.numero = 0
            .referencia.textura = 0
        ElseIf .formato = eFormatoTileSet.camino_chico Or .formato = eFormatoTileSet.camino_grande_parte1 Then
            ' Guardan referencia a una textura simple
            sectorTextura = Me.cmbConfigTipo_Sector.itemData(Me.cmbConfigTipo_Sector.listIndex)
            
            .referencia.textura = Me.cmbConfigTipo_TexturaSimple.itemData(Me.cmbConfigTipo_TexturaSimple.listIndex)
            .referencia.numero = sectorTextura
            
            ' Generamos la matriz
            If .formato = eFormatoTileSet.camino_chico Then
                Call Me_indexar_Pisos.setearMatrizTranformacion(TmpTilesetsNum)
            End If
            
        ElseIf .formato = eFormatoTileSet.camino_grande_parte2 Then
            ' Guardamoa la referencia a la primera parte
            .referencia.textura = Me.cmbConfigTipo_CaminoParte1.itemData(Me.cmbConfigTipo_CaminoParte1.listIndex)
            .referencia.numero = 0
            
            ' Generamos la matriz
            Call Me_indexar_Pisos.setearMatrizTranformacion(TmpTilesetsNum)
            
        ElseIf .formato = eFormatoTileSet.costa_tipo_1_parte1 Then
            ' Guardamos la referencia al agua
            .referencia.textura = Me.cmbConfigTipo_Agua.itemData(Me.cmbConfigTipo_Agua.listIndex)
            .referencia.numero = 0
            
        ElseIf .formato = eFormatoTileSet.costa_tipo_1_parte2 Then
            ' Guardamos la referencia al agua
            .referencia.textura = Me.cmbConfigTipo_CostaParte1.itemData(Me.cmbConfigTipo_CostaParte1.listIndex)
            .referencia.numero = 0
            
            ' Generamos la matriz
            Call Me_indexar_Pisos.setearMatrizTranformacion(TmpTilesetsNum)
        Else
            Call MsgBox("Formato de piso desconocido. Consulte al Administrador del Sistema.", vbExclamation)
            Exit Sub
        End If
        
        
    End With
    
    ' Guardamos la modificacion en el .ini
    Call Me_indexar_Pisos.actualizarEnIni(TmpTilesetsNum)
        
    'Actualizamos el nombre en la lista
    Call Listado.cambiarNombre(CInt(TmpTilesetsNum), TmpTilesetsNum & " - " & nombre)
    
    ' Actualizamos
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_TexturaSimple, eFormatoTileSet.textura_simple)
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_CaminoParte1, eFormatoTileSet.camino_grande_parte1)
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_CostaParte1, eFormatoTileSet.costa_tipo_1_parte1)
    Call configTipo_CargarTexturas(Me.cmbConfigTipo_Agua, eFormatoTileSet.textura_agua)
    
    'Activamoslos botones
    cmdAplicar_Pisos.Enabled = False
    cmdRestablecer_Pisos.Enabled = False
    check_enabled_guardar
    
    Exit Sub
errorH:
Tilesets(TmpTilesetsNum) = Tileset_BACKUP
MsgBox "Error al aplicar los cambios, comprobá los datos."
End Sub

Sub reestablece()
    Dim n As Integer
    Dim sector As Byte
    
    With Tilesets(TmpTilesetsNum)
        Index.text = TmpTilesetsNum
        nombre.text = .nombre
        Speed.value = .anim
        
         
        Call txtOlitas.seleccionarID(.Olitas)
        
        Call frames.limpiar
        
        If .stage_count Then
            For n = 1 To .stage_count
                Call frames.seleccionarID(n - 1, .stages(n))
            Next n
        End If
        
        Me.cmbTipo.listIndex = .formato
        Call mostrarConfiguracionFormato(.formato)
        
        If .formato = eFormatoTileSet.camino_chico Or .formato = eFormatoTileSet.camino_grande_parte1 Then
            If .referencia.textura > 0 Then
                Call seleccionarComboID(Me.cmbConfigTipo_TexturaSimple, .referencia.textura)
                Call seleccionarComboID(Me.cmbConfigTipo_Sector, .referencia.numero)
            End If
        ElseIf .formato = camino_grande_parte2 Then
        
            If .referencia.textura > 0 Then Call seleccionarComboID(Me.cmbConfigTipo_CaminoParte1, .referencia.textura)
 
        ElseIf .formato = costa_tipo_1_parte1 Then
        
            If .referencia.textura > 0 Then Call seleccionarComboID(Me.cmbConfigTipo_Agua, .referencia.textura)

        ElseIf .formato = costa_tipo_1_parte2 Then
            
            If .referencia.textura > 0 Then Call seleccionarComboID(Me.cmbConfigTipo_CostaParte1, .referencia.textura)
        
        End If
        
        
    End With
        
    cmdAplicar_Pisos.Enabled = False
    cmdRestablecer_Pisos.Enabled = False
    
End Sub

Private Sub seleccionarComboID(combo As VB.ComboBox, ByVal ID As Long)

    Dim loopElemento As Integer
    
    For loopElemento = 0 To combo.ListCount - 1
        If combo.itemData(loopElemento) = ID Then
            combo.listIndex = loopElemento
        End If
    Next

End Sub

Private Sub Speed_Change(valor As Double)
    eb
End Sub

Private Sub txtOlitas_Change(valor As String, ID As Integer)
    eb
End Sub

Private Sub txtSonidos_Change(Index As Integer, valor As String, ID As Integer)
    eb
End Sub
