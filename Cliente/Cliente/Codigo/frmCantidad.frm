VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   1680
   ClientTop       =   4455
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Todo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      MouseIcon       =   "frmCantidad.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1035
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A&ceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      MouseIcon       =   "frmCantidad.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1680
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   1
      Top             =   525
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba la cantidad:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   585
      TabIndex        =   0
      Top             =   165
      Width           =   2415
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmCantidad.Visible = False
frmCantidad.Text1.text = Int(val(frmCantidad.Text1.text))
If val(frmCantidad.Text1.text) <= 0 Then Exit Sub
    If Not DeAmuchos Or ItemDragued = 0 Then
        If itemElegido = FLAGORO Then itemElegido = 254
        EnviarPaquete Tirar, Chr$(itemElegido) & Codify(val(frmCantidad.Text1.text))
    Else
    If UserInventory(ItemDragued).Amount < val(frmCantidad.Text1.text) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "No tenés esa cantidad.", 2, 51, 223, 1, 1: Exit Sub
    If ItemDragued = 0 Then frmCantidad.Text1.text = "": Exit Sub
    If Not UserMeditar Then EnviarPaquete Paquetes.DIClick, Chr$(tx) & Chr$(ty) & Chr$(ItemDragued) & ITS(val(frmCantidad.Text1.text))
    ItemDragued = 0
    DeAmuchos = False
    End If
    
frmCantidad.Text1.text = ""
End Sub

Private Sub Command2_Click()
frmCantidad.Visible = False
If Not DeAmuchos Or ItemDragued = 0 Then
    If itemElegido = FLAGORO Then
        itemElegido = 254: If UserGLD > 100000 Then Exit Sub
        EnviarPaquete Tirar, Chr$(itemElegido) & Codify(val(UserGLD))
    Else
        EnviarPaquete Tirar, Chr$(itemElegido) & ITS(val(UserInventory(itemElegido).Amount))
    End If
Else
If UserInventory(ItemDragued).Amount < val(frmCantidad.Text1.text) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "No tenés esa cantidad.", 2, 51, 223, 1, 1: Exit Sub
If ItemDragued = 0 Then frmCantidad.Text1.text = "": Exit Sub
EnviarPaquete Paquetes.DIClick, Chr$(tx) & Chr$(ty) & Chr$(ItemDragued) & ITS(val(UserInventory(itemElegido).Amount))
ItemDragued = 0
DeAmuchos = False
End If
frmCantidad.Text1.text = ""
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmCantidad)
End Sub

Private Sub Text1_Change()
If val(Text1.text) < 0 Then
    Text1.text = MAX_INVENTORY_OBJS
End If
If val(Text1.text) > MAX_INVENTORY_OBJS And itemElegido <> FLAGORO Then
    Text1.text = 1
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim Index As Integer
If (KeyAscii <> 8) Then
    If (Index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub
