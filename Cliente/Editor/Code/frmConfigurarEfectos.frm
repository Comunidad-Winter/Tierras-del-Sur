VERSION 5.00
Begin VB.Form frmConfigurarEfectos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Efectos"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigurarEfectos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEliminar_Efecto 
      Caption         =   "Eliminar"
      Height          =   360
      Left            =   1680
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevo_Efecto 
      Caption         =   "Nuevo"
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame frmPropiedades 
      Caption         =   "Propiedades"
      Height          =   3945
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox chkRepetir 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   310
         TabIndex        =   20
         ToolTipText     =   "Repetir 10 veces el efecto"
         Top             =   3560
         Width           =   255
      End
      Begin EditorTDS.TextConListaConBuscador lstGraficos 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
      End
      Begin EditorTDS.TextConListaConBuscador lstParticulas 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
      End
      Begin EditorTDS.TextConListaConBuscador lstSonidos 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   1800
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "Aplicar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Guarda las modificaciones"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton cmdProbar 
         Caption         =   "Probar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Muestra el efecto en el Personaje del modo caminata"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   2160
         TabIndex        =   16
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmdRestablecer 
         Caption         =   "Restablecer"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2160
         TabIndex        =   19
         ToolTipText     =   "Restablece las propiedades del efecto previo a que se comience a modificar"
         Top             =   3000
         Width           =   1695
      End
      Begin EditorTDS.UpDownText txtPosicionY 
         Height          =   310
         Left            =   1080
         TabIndex        =   21
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MaxValue        =   1000
         MinValue        =   -1000
         Enabled         =   -1  'True
      End
      Begin EditorTDS.UpDownText txtPosicionX 
         Height          =   310
         Left            =   1080
         TabIndex        =   22
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MaxValue        =   1000
         MinValue        =   -1000
         Enabled         =   -1  'True
      End
      Begin VB.Label lblNumeroEfectoResultado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   360
         Width           =   2700
      End
      Begin VB.Label lblNumeroEfecto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblPosicionY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posición Y"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2570
         Width           =   705
      End
      Begin VB.Label lblPosicionX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posición X"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label lblSonido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblParticula 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particula"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblAnimacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Animación"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   555
      End
   End
   Begin EditorTDS.ListaConBuscador lstEfectos 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6165
   End
End
Attribute VB_Name = "frmConfigurarEfectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ignorarChangePropiedades As Boolean

Private fxbackup As tIndiceFx

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdAplicar_Click()
    Dim fxID As Integer
    
    fxID = CInt(val(Me.lblNumeroEfectoResultado))

    FxData(fxID).Animacion = Me.lstGraficos.obtenerIDValor
    FxData(fxID).particula = Me.lstParticulas.obtenerIDValor
    FxData(fxID).wav = Me.lstSonidos.obtenerIDValor
    
    FxData(fxID).offsetX = CInt(val(Me.txtPosicionX.value))
    FxData(fxID).offsetY = CInt(val(Me.txtPosicionY.value))

    FxData(fxID).nombre = Me.txtNombre
        
    Call Me.lstEfectos.cambiarNombre(fxID, fxID & " - " & FxData(fxID).nombre)
    
    Me.cmdAplicar.Enabled = False
    
    'Guardamos el cambio
    Call Me_indexar_Efectos.actualizarEnIni(fxID)
        
    'Inidicamos que hay algo sin actualizar en el cliente
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Efectos")
End Sub

Private Sub cmdEliminar_Efecto_Click()
    Dim confirma As VbMsgBoxResult
    Dim idElemento As Integer
    
    If Not Me.lstEfectos.obtenerValor = "" Then
        
        idElemento = Me.lstEfectos.obtenerIDValor
        
        confirma = MsgBox("¿Está seguro de que desea eliminar el efecto '" & Me.lstEfectos.obtenerValor & "'?", vbYesNo + vbExclamation, Me.caption)
        
        If confirma = vbYes Then
            Call Me_indexar_Efectos.eliminar(idElemento)
            'Lo borramos de la lista
            Call Me.lstEfectos.eliminar(CLng(idElemento))
            'Deshabilitamos los botones
            Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades, Me)
            Me.cmdEliminar_Efecto.Enabled = False
        End If
    End If
End Sub

Private Sub cmdNuevo_Efecto_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    error = False
    Me.cmdNuevo_Efecto.Enabled = False
    
    'Obtengo el nuevo id
    nuevo = Me_indexar_Efectos.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar un nuevo efecto. Por favor, intente más tarde o contacte a un administrador.", vbExclamation
    End If
    
    If Not error Then
        'Lo selecciono
        If Me.lstEfectos.seleccionarID(CLng(nuevo)) = False Then
            Call Me.lstEfectos.addString(nuevo, nuevo & " - ")
            Call Me.lstEfectos.seleccionarID(CLng(nuevo))
        End If
    End If
    
    Me.cmdNuevo_Efecto.Enabled = True
    
    'Cuando se haga clic en "Aplicar" se guarda
End Sub

Private Sub cmdProbar_Click()
    Call SetCharacterFx(UserCharIndex, Me.lstEfectos.obtenerIDValor, 10 * Me.chkRepetir.value)
End Sub

Private Sub cmdRestablecer_Click()
    FxData(Me.lstEfectos.obtenerIDValor) = fxbackup
    Call cargarEfectoEnEditor(Me.lstEfectos.obtenerIDValor)
    
    Me.cmdAplicar.Enabled = False
    Me.cmdRestablecer.Enabled = False
End Sub


Private Sub Form_Load()

Dim i As Long

'Sino esta el modo caminata activado lo activamos
If ME_Render.WalkMode = False Then
    Call ME_Render.ToggleWalkMode
End If

'Sino pude activarlo, salgo
If ME_Render.WalkMode = False Then Unload Me: Exit Sub

'Cargoslos efectos
For i = 1 To UBound(FxData)
    If Me_indexar_Efectos.existe(i) Then
        Call Me.lstEfectos.addString(CInt(i), i & " - " & FxData(i).nombre)
    End If
Next i

'Cargo los grh posibles que tiene el juego
Call Me.lstGraficos.addString(CInt(0), 0 & " - Sin Grafico")
For i = 1 To UBound(GrhData)
    If GrhData(i).NumFrames > 0 Then
        Call Me.lstGraficos.addString(i, i & " - " & GrhData(i).nombreGrafico)
    End If
Next

' Cargamos las particulas
Call Me.lstParticulas.addString(CInt(0), 0 & " - Sin Particulas")
'For i = 1 To UBound(GlobalParticleGroup)
'   Call Me.lstParticulas.addString(i, i & " - " & GlobalParticleGroup(i).GetNombre())
'Next

'Cargamos los sonidos
Call Me.lstSonidos.addString(CInt(0), 0 & " - Sin Sonido")
For i = 1 To UBound(Me_indexar_Sonidos.Sonidos)
    If Me_indexar_Sonidos.existe(i) Then
        Call Me.lstSonidos.addString(i, i & " - " & Sonidos(i).nombre)
    End If
Next

'Opciones de visualizacion
Me.lstGraficos.CantidadLineasAMostrar = 10
Me.lstParticulas.CantidadLineasAMostrar = 8
Me.lstSonidos.CantidadLineasAMostrar = 6

' Estado de los botones inicialmente
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades, Me)
Me.cmdAceptar.Enabled = True
Me.cmdEliminar_Efecto.Enabled = False
End Sub

Private Sub cargarEfectoEnEditor(ID As Integer)

    ignorarChangePropiedades = True

    Me.lblNumeroEfectoResultado = ID

    Me.txtNombre = FxData(ID).nombre
    Me.txtPosicionX.value = FxData(ID).offsetX
    Me.txtPosicionY.value = FxData(ID).offsetY
    
    Call Me.lstParticulas.seleccionarID(FxData(ID).particula)
    Call Me.lstGraficos.seleccionarID(FxData(ID).Animacion)
    Call Me.lstSonidos.seleccionarID(FxData(ID).wav)
    
    Me.txtPosicionX.Enabled = True
    Me.txtPosicionY.Enabled = True
    Me.txtNombre.Enabled = True
    
    'Botones
    Me.cmdProbar.Enabled = True
    Me.cmdAplicar.Enabled = False
    Me.cmdRestablecer.Enabled = False
    
    ignorarChangePropiedades = False
End Sub

Private Sub lstEfectos_Change(valor As String, ID As Integer)
    Call cargarEfectoEnEditor(ID)
    
    fxbackup = FxData(ID)
    
    'Botones
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades, Me)
    Me.cmdAceptar.Enabled = True
    Me.cmdEliminar_Efecto.Enabled = True
End Sub

Private Sub lstGraficos_Change(valor As String, ID As Integer)
   Call actualizarEfecto
End Sub

Private Sub actualizarEfecto()
    If Not ignorarChangePropiedades Then
        Me.cmdRestablecer.Enabled = True
        Me.cmdAplicar.Enabled = True
    End If
End Sub

Private Sub lstParticulas_Change(valor As String, ID As Integer)
    Call actualizarEfecto
End Sub

Private Sub lstSonidos_Change(valor As String, ID As Integer)
    Call actualizarEfecto
End Sub

Private Sub txtNombre_Change()
    Call actualizarEfecto
End Sub

Private Sub txtPosicionX_Change(valor As Double)
    Call actualizarEfecto
End Sub

Private Sub txtPosicionY_Change(valor As Double)
    Call actualizarEfecto
End Sub
