VERSION 5.00
Begin VB.Form frmConfigurarSonidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Sonidos"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigSonidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstAuxList 
      Height          =   1035
      Left            =   1080
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar_Sonidos 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   3240
      Width           =   1470
   End
   Begin VB.CommandButton cmdNuevo_Sonidos 
      Caption         =   "Nuevo"
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame frmPropiedades 
      Caption         =   "Propiedades"
      Height          =   3495
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton optTipoSonido 
         Appearance      =   0  'Flat
         Caption         =   "Es una Música"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   20
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton optTipoSonido 
         Appearance      =   0  'Flat
         Caption         =   "Es un efecto"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdPararReproducccion 
         Height          =   360
         Left            =   1800
         Picture         =   "frmConfigSonidos.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin EditorTDS.TextConListaConBuscador txtConAutoCompletar 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "Aplicar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdRestablecer 
         Caption         =   "Restablecer"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2400
         TabIndex        =   13
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   2400
         TabIndex        =   12
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton cmdProbar 
         Caption         =   "Probar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox chkReproducirInfinitamente 
         Appearance      =   0  'Flat
         Caption         =   "Probar con re producir infinitamente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         HelpContextID   =   -2147483633
         Left            =   1800
         TabIndex        =   16
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1800
         Picture         =   "frmConfigSonidos.frx":200C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Escuchar"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblAlerta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El número de sonido debe ser el mismo que el número del archivo de sonido."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSonidoRecurso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo de Sonido:"
         Height          =   255
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         ToolTipText     =   "Primero tenes que haber agregado el archivo de sonido desde las sección de Recursos"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   660
         Width           =   555
      End
      Begin VB.Label lblNumeroSonidoResultado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   60
      End
      Begin VB.Label lblNumeroSonido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   555
      End
   End
   Begin EditorTDS.ListaConBuscador lstSonidos 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5530
   End
End
Attribute VB_Name = "frmConfigurarSonidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sonidoBackup As tSonido
Private ignorarChangePropiedades As Boolean
Private sonidoIDSonando As Integer

Private Sub cmdAceptar_Click()
    Call cerrarFormulario
End Sub

Private Sub cmdAplicar_Click()
    Dim sonidoID As Integer
    
    sonidoID = CInt(val(Me.lblNumeroSonidoResultado))

    Sonidos(sonidoID).tipo = IIf(Me.optTipoSonido(0).value, 0, 1)
    Sonidos(sonidoID).nombre = Me.txtNombre
        
    Call Me.lstSonidos.cambiarNombre(sonidoID, sonidoID & " - " & Sonidos(sonidoID).nombre)
    
    'Botones
    Me.cmdAplicar.Enabled = False
    Me.cmdRestablecer.Enabled = False
    
    'Guardamos el cambio
    Call Me_indexar_Sonidos.actualizarEnIni(sonidoID)
End Sub

Private Sub cmdEliminar_Sonidos_Click()
    Dim confirma As VbMsgBoxResult
    Dim idElemento As Integer
    
    If Not Me.lstSonidos.obtenerValor = "" Then
        
        idElemento = Me.lstSonidos.obtenerIDValor
        
        confirma = MsgBox("¿Está seguro de que desea eliminar el sonido '" & Me.lstSonidos.obtenerValor & "'?", vbYesNo + vbExclamation, Me.caption)
        
        If confirma = vbYes Then
            Call Me_indexar_Sonidos.eliminar(idElemento)

            'Lo borramos de la lista
            Call Me.lstSonidos.eliminar(CLng(idElemento))
            
            'Deshabilitamos los botones
            Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades, Me)
        End If
    End If
End Sub

Private Sub cmdNuevo_Sonidos_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    error = False
    Me.cmdNuevo_Sonidos.Enabled = False

    'Obtengo el nuevo id
    nuevo = Me_indexar_Sonidos.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar un nuevo sonido. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        'Lo selecciono
        If Me.lstSonidos.seleccionarID(CLng(nuevo)) = False Then
            Call Me.lstSonidos.addString(nuevo, nuevo & " - ")
            Call Me.lstSonidos.seleccionarID(CLng(nuevo))
        End If
    End If
 
    Me.cmdNuevo_Sonidos.Enabled = True
End Sub

Private Sub cerrarFormulario()
    Call pararReproduccion
    Unload Me
End Sub
Private Sub pararReproduccion()
    If sonidoIDSonando > 0 Then
        Call Sonido_Stop(sonidoIDSonando)
    End If
    
    Me.cmdPararReproducccion.visible = False
End Sub
Private Sub cmdPararReproducccion_Click()
    Call pararReproduccion
End Sub

Private Sub cmdProbar_Click()
    Dim idSonidoActual As Integer
    If Me.lstSonidos.obtenerIDValor > 0 Then
        idSonidoActual = Me.lstSonidos.obtenerIDValor
        
        Call Engine_Sonido.Sonido_PlayEX(idSonidoActual, False)
    End If
End Sub

Private Sub cmdRestablecer_Click()
    Dim idSonido As Integer
    
    idSonido = CInt(val(Me.lblNumeroSonidoResultado))
    
    Sonidos(idSonido) = sonidoBackup
    Call cargarEnEditor(idSonido)
    
    Me.cmdAplicar.Enabled = False
    Me.cmdRestablecer.Enabled = False
End Sub

Private Sub Command1_Click()

If Me.txtConAutoCompletar.obtenerIDValor > 0 Then
    sonidoIDSonando = Me.txtConAutoCompletar.obtenerIDValor
    Call Engine_Sonido.Sonido_PlayEX(sonidoIDSonando, (Me.chkReproducirInfinitamente.value = 1))
End If

'Activamos la opcion para que pare el sonido
If Me.chkReproducirInfinitamente.value = 1 Then
    Me.cmdPararReproducccion.visible = True
End If

End Sub


Private Sub Form_Load()

Dim i As Integer

' Cargo los sonidos ya indexados
For i = 0 To UBound(Me_indexar_Sonidos.Sonidos)
    If Me_indexar_Sonidos.existe(i) Then
        Call Me.lstSonidos.addString(i, i & " - " & Sonidos(i).nombre)
    End If
Next

'Negrada pero no quiero modificar la clase de empaquetado
Call pakSonidos.Add_To_Listbox(Me.lstAuxList)

For i = 0 To Me.lstAuxList.ListCount
   Call Me.txtConAutoCompletar.addString(val(Me.lstAuxList.list(i)), Me.lstAuxList.list(i))
Next

Me.txtConAutoCompletar.CantidadLineasAMostrar = 8

'Inicialmente el formulario arranca desactivado
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades, Me)
End Sub

Private Sub actualizarSonido()
    If Not ignorarChangePropiedades Then
        Me.cmdRestablecer.Enabled = True
        Me.cmdAplicar.Enabled = True
    End If
End Sub
Private Sub cargarEnEditor(id As Integer)
    Dim idarchivo As Integer
    Dim loopOpcion As Byte
    
    idarchivo = id
    If Not Me.txtConAutoCompletar.seleccionarID(idarchivo) Then
        MsgBox "El archivo de sonido " & id & " no existe en el archivo de recursos de sonido. Primero tenes que agregar el sonido desde el menú Recursos > Agregar/Cambiar Recursos", vbExclamation, Me.caption
        Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades, Me)
        Exit Sub
    End If
    
    ignorarChangePropiedades = True
    
    Me.lblNumeroSonidoResultado.caption = id
    
    Me.txtNombre = Sonidos(id).nombre
    
    For loopOpcion = Me.optTipoSonido.LBound To Me.optTipoSonido.UBound
        Me.optTipoSonido.item(loopOpcion).value = (loopOpcion = Sonidos(id).tipo)
    Next

    Me.txtConAutoCompletar.Enabled = False
    
    'Guardo una copia por si quiere establecer
    sonidoBackup = Sonidos(id)
    
    'Botones
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades, Me)
    Me.txtConAutoCompletar.Enabled = False
    Me.cmdEliminar_Sonidos.Enabled = True
    
    ignorarChangePropiedades = False
    
End Sub

Private Sub lstSonidos_Change(valor As String, id As Integer)
    Call cargarEnEditor(id)
End Sub

Private Sub optTipoSonido_Click(Index As Integer)
    Call actualizarSonido
End Sub

Private Sub txtConAutoCompletar_Change(valor As String, id As Integer)
    Call actualizarSonido
End Sub

Private Sub txtNombre_Change()
    Call actualizarSonido
End Sub
