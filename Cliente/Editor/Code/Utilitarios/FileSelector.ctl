VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl FileSelector 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   375
   ScaleWidth      =   4485
   Begin MSComDlg.CommonDialog diagDialogo 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSeleccionar 
      Height          =   375
      Left            =   4100
      Picture         =   "FileSelector.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Clic para seleccionar un archivo"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtArchivo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FileSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event change(valor As String)

'Propiedades tipicas de un control
Public Property Get Enabled() As Boolean
   Enabled = txtArchivo.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Dim Control As Control
    
    For Each Control In Controls
        Control.Enabled = vNewValue
    Next
End Property

Private Sub cmdSeleccionar_Click()
    On Error GoTo BotonCancelar

    diagDialogo.CancelError = True 'Si la persona toca cancelar, se genera un error
    diagDialogo.flags = cdlOFNHideReadOnly
    diagDialogo.InitDir = app.Path
    diagDialogo.DialogTitle = "Seleccionar archivo"
    diagDialogo.ShowOpen
    
    txtArchivo.text = diagDialogo.FileName

Exit Sub
BotonCancelar:
    Err.Clear
    Exit Sub
End Sub

Public Function obtenerArchivo() As String
    obtenerArchivo = txtArchivo.text
End Function

Property Get text() As String
    text = txtArchivo
End Property

Property Let text(valor As String)
    txtArchivo = valor
End Property

Private Sub txtArchivo_Change()
    RaiseEvent change(txtArchivo.text)
End Sub

Private Sub UserControl_Resize()
    txtArchivo.Width = UserControl.Width - 375
    txtArchivo.Height = UserControl.Height - 15
    
    cmdSeleccionar.left = txtArchivo.Width
    cmdSeleccionar.Height = UserControl.Height - 1
End Sub
