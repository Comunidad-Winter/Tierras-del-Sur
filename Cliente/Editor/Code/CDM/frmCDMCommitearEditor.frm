VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmCDMCommitearEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tierras del Sur - Compartir novedades"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCDMCommitearEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmArchivosEspeciales 
      Caption         =   "Archivos especiales"
      Height          =   2655
      Left            =   5400
      TabIndex        =   10
      Top             =   2040
      Width           =   5055
      Begin VB.TextBox txtCarpetaRoot 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   310
         Width           =   3255
      End
      Begin EditorTDS.GridFileSelected GridFileSelected1 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2990
      End
      Begin VB.Label lblCarpetaRoot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta Root:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Timer frmIniciar 
      Interval        =   1
      Left            =   720
      Top             =   4680
   End
   Begin VB.Frame frmMensaje 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   4935
      Begin MSComctlLib.ProgressBar prgCompartir 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4635
      End
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      Top             =   4800
      Width           =   1350
   End
   Begin VB.CommandButton cmdCompartir 
      Caption         =   "Compartir"
      Height          =   360
      Left            =   3960
      TabIndex        =   0
      Top             =   4800
      Width           =   1350
   End
   Begin EditorTDS.TreeConBuscador lstCambios 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4895
   End
   Begin VB.Label lblAlerta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por el momento no es posible compartir elementos individuales, por ejemplo, seleccionar sólo un piso."
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5205
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblIngresaComentario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingresá un comentario acerca de las novedades que vas a compartir."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4980
   End
   Begin VB.Label lblSeleccione 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccioná los elementos modificados que queres compartir"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4230
   End
End
Attribute VB_Name = "frmCDMCommitearEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents repositorio As clsCDM ' Para recibir los eventos del repositorio
Attribute repositorio.VB_VarHelpID = -1

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub obtenerElementosACommitear(archivos() As String)
    Dim loopElemento As Integer
    Dim elemento As Integer
    Dim cantidadSeleccionados As Integer
    
    cantidadSeleccionados = Me.lstCambios.cantidadChecked(True)
    loopElemento = 0
    elemento = 0
    
    If cantidadSeleccionados > 0 Then
    'Obtenemos los archivos que desea modificar
    ReDim archivos(1 To cantidadSeleccionados)
    
    Do While elemento < cantidadSeleccionados
        
        If Me.lstCambios.estaChequeado(loopElemento + 1) Then
            elemento = elemento + 1
            archivos(elemento) = Trim$(mid$(Me.lstCambios.list(loopElemento + 1), 1, InStr(1, Me.lstCambios.list(loopElemento + 1), "(") - 1))
        End If
        
        loopElemento = loopElemento + 1
            
    Loop
    End If
    
End Sub
Private Sub cmdCompartir_Click()
   
    Dim archivos() As String
    Dim comentarioCommit As String
    Dim loopArchivo As Integer
    Dim cantidadSeleccionadaLista As Integer
    Dim cantidadSeleccionadaEspeciales As Integer
    Dim archivoEspecial As String
    
    cantidadSeleccionadaLista = Me.lstCambios.cantidadChecked(True)
    cantidadSeleccionadaEspeciales = Me.GridFileSelected1.cantidad
    
    ' ¿Selecciono algo?
    If cantidadSeleccionadaLista + cantidadSeleccionadaEspeciales = 0 Then
        MsgBox "No has seleccionado ninguna novedad para compartir.", vbInformation, Me.caption
        Exit Sub
    End If
    
    ' Tiene que poner un comentario
    comentarioCommit = Trim$(Me.txtComentario)
    
    If Len(comentarioCommit) = 0 Then
        MsgBox "Por favor, ingresá un comentario de lo que estas compartiendo.", vbInformation, Me.caption
        Exit Sub
    End If
        
    ' De la lista seleccionada
    Call obtenerElementosACommitear(archivos)
    
    ReDim Preserve archivos(1 To cantidadSeleccionadaLista + cantidadSeleccionadaEspeciales) As String
    
    ' De los archivos especiales
    For loopArchivo = cantidadSeleccionadaLista To cantidadSeleccionadaLista + cantidadSeleccionadaEspeciales - 1
        archivoEspecial = Me.GridFileSelected1.text(loopArchivo - cantidadSeleccionadaLista)
        If Not right$(archivoEspecial, 4) = ".exe" Then
            archivos(loopArchivo + 1) = archivoEspecial
        Else
            Call MsgBox("No se pueden compartir archivos '.exe'. Tenes que cambiarle la extensión y poner '.exe_'", vbExclamation)
            Exit Sub
        End If
    Next
    
    Me.lblMensaje.caption = "Compartiendo..."
        
    ' Desactivo el boton
    Me.cmdCompartir.Enabled = False
    
    ' Barra de progeso
    Me.prgCompartir.min = 0
    Me.prgCompartir.max = 10
    Me.prgCompartir.value = 0
    Me.frmMensaje.Visible = True
    
  
    'Lo enviamos
    Call repositorio.Repositorio_Compartir(archivos, comentarioCommit, Me.txtCarpetaRoot)
End Sub

Private Sub iniciar()
    Dim archivos() As versionador.tArchivoAlterado
    Dim total As Integer
    Dim loopArchivo As Integer
    Dim infoArchivo As Dictionary
    Dim id As Long
    Dim nombre As String
    
    ' Buscamos novedades
    Call versionador.obtenerArchivosAlterados(total, archivos)
         
    If total > 0 Then
        ' Hay novedades para compartir
        Me.lstCambios.vaciar
        Me.lstCambios.checked = True
        
        For loopArchivo = LBound(archivos) To UBound(archivos)
            Call Me.lstCambios.addString(loopArchivo, archivos(loopArchivo).Tipo & " (Creados: " & archivos(loopArchivo).creados.count & ", Modificados: " & archivos(loopArchivo).modificados.count & ", Eliminados: " & archivos(loopArchivo).eliminados.count & ")", 0)
        
            For Each infoArchivo In archivos(loopArchivo).info
             
                id = infoArchivo.item("id")
                nombre = infoArchivo.item("nombre")
                Call Me.lstCambios.addString(id, id & " - " & infoArchivo.item("accion") & IIf(Len(nombre) > 0, ": " & nombre, ""), loopArchivo)
            
            Next
        
        Next loopArchivo
        
        ' Chequeo que tenga el Editor actualizado
        
                
        Dim estoyActualizado As Boolean
        
        estoyActualizado = repositorio.estoyActualizado
    
        If Len(repositorio.ultimoError) > 0 Then
            MsgBox "Se produjo un error en la conexión con el Cerebro de Mono. Por favor, intentá más tarde. " & repositorio.ultimoError & ".", vbExclamation
            Me.cmdCompartir.Enabled = False
            Exit Sub
        End If
        
        If Not estoyActualizado Then
           ' Me.lstCambios.Enabled = False
           ' Me.cmdCompartir.Enabled = False
            MsgBox "El editor no está actualizado. Antes de compartir novedades tenes que actualizar el Editor.", vbExclamation, "Cuidado"
            Me.cmdCompartir.Enabled = False
            Exit Sub
        End If
    Else
        Me.lstCambios.vaciar
        Me.lstCambios.addString 0, "Sin novedades para compartir", 0
       ' Me.lstCambios.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
    Set repositorio = CDM.cerebro
    
    Me.lblMensaje.caption = "Recopilando info..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set repositorio = Nothing
    Unload Me
End Sub

Private Sub frmIniciar_Timer()
    frmIniciar.Enabled = False
    '*************************
    iniciar
    
    frmMensaje.Visible = False
End Sub

Private Sub repositorio_compartido(Version As Long)
    
    Me.frmMensaje.Visible = False
    
    If Version > 0 Then
        'Activo el boton
        Me.cmdCancelar.Enabled = False
                
        'Avisamos
        MsgBox "Novedades publicadas. Un pasito más cerca de la nueva version!. La versión interna es la " & Version & ".", vbInformation, Me.caption
                
       ' Desbloqueamos
        Unload Me
    Else
        Me.Enabled = True
        Me.cmdCompartir.Enabled = True
        MsgBox "Se ha producido un error al intentar compartir las novedades. Error: " & repositorio.ultimoError, vbExclamation, "Error al compartir"
    End If
End Sub

Private Sub repositorio_Progreso(actual As Single, Maximo As Single)
    Me.prgCompartir.max = Maximo
    Me.prgCompartir.value = actual
End Sub

