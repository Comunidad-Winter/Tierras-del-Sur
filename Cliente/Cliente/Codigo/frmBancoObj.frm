VERSION 5.00
Object = "{50CBA22D-9024-11D1-AD8F-8E94A5273767}#8.7#0"; "TranImg2.ocx"
Begin VB.Form frmBancoObj 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   3765
   ClientTop       =   2550
   ClientWidth     =   6165
   ControlBox      =   0   'False
   Icon            =   "frmBancoObj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DevPowerTransImg.TransImg itemimg 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      AutoSize        =   0   'False
      MaskColor       =   0
      Transparent     =   -1  'True
   End
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Text            =   "1"
      Top             =   3000
      Width           =   525
   End
   Begin VB.PictureBox Bovinventario 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   1920
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   2
      Top             =   3240
      Width           =   2880
   End
   Begin VB.PictureBox BovBoveda 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   225
      ScaleHeight     =   158
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   0
      Top             =   360
      Width           =   3840
   End
   Begin VB.Image Label1 
      Height          =   765
      Index           =   2
      Left            =   5535
      Top             =   2205
      Width           =   435
   End
   Begin VB.Image Label1 
      Height          =   480
      Index           =   0
      Left            =   4920
      Top             =   5070
      Width           =   1110
   End
   Begin VB.Image Label1 
      Height          =   735
      Index           =   1
      Left            =   5520
      Top             =   3300
      Width           =   435
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   11
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   10
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   9
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Top             =   600
      Width           =   1740
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   7
      Top             =   5040
      Width           =   1500
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   3600
      Width           =   1770
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BX As Integer
Private by As Integer
Private BintemElegido As Byte
Private BBitemElegido As Byte
Private DragMouse As Byte
Private ItemDragued As Byte
Private Selecionado As Byte

Private UserBancoInventory() As Inventory

Public Sub refrescar()
    Call DrawInvSimple(Me.BovBoveda, UserBancoInventory, BBitemElegido)
    Call DrawInvSimple(Me.Bovinventario, UserInventory, BintemElegido)
End Sub

Friend Sub setBoveda(ByRef inventario() As Inventory)
    UserBancoInventory = inventario
End Sub

Friend Sub setSlot(Slot As Byte, Inventory As Inventory)
    UserBancoInventory(Slot) = Inventory
End Sub


Private Sub BovBoveda_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If BBitemElegido > 0 And BBitemElegido <= MAX_BANCOINVENTORY_SLOTS Then
If Button = 2 And UserStats(SlotStats).UserEstado = 0 And UserBancoInventory(BBitemElegido).GrhIndex > 0 Then
Call CambiarCursor(Me, 1)
ItemDragued = BBitemElegido

Set itemimg.Picture = clsEnpaquetado_LeerIPicture(pakGraficos, GrhData(UserBancoInventory(BBitemElegido).GrhIndex).filenum)

DragMouse = 1
itemimg.Visible = True
End If
End If
End Sub

Private Sub BovBoveda_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragMouse = 1 Then
Call itemimg.Move(x, y)
End If
End Sub

Private Sub BovBoveda_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If x > 0 And y > 0 And x < Me.BovBoveda.ScaleWidth And y < Me.BovBoveda.ScaleHeight Then
    BX = x \ 33 + 1
    by = y \ 33 + 1
    BBitemElegido = (BX + (by - 1) * 8)
    If BBitemElegido <= MAX_BANCOINVENTORY_SLOTS Then
        If UserBancoInventory(BBitemElegido).GrhIndex > 0 Then
            'Call Dibujar(CInt(BBitemElegido), Me.BovBoveda, UserBancoInventory, 8)
            
            Me.Bovedalbl(0) = UserBancoInventory(BBitemElegido).Name
            Me.Bovedalbl(1) = UserBancoInventory(BBitemElegido).MinDef & "/" & UserBancoInventory(BBitemElegido).MaxDef
            Me.Bovedalbl(2) = UserBancoInventory(BBitemElegido).MaxHit
            Me.Bovedalbl(3) = UserBancoInventory(BBitemElegido).MinHit
            Me.Bovedalbl(0).ToolTipText = UserBancoInventory(BBitemElegido).Name
        End If
    End If
Else
If Not Shift = 100 Then Call BovInventario_MouseUp(Button, Shift, 1, 1): Exit Sub
End If

If Button = 2 Then
    If DragMouse = 1 Then
        If ItemDragued < 100 Then
            If Not ItemDragued = BBitemElegido Then
            EnviarPaquete ChangeItemsSlotboveda, Chr$(BBitemElegido) & Chr$(ItemDragued)
            End If
        Else
        ItemDragued = ItemDragued - 100
        EnviarPaquete Paquetes.Depositar, Chr$(ItemDragued) & Codify(cantidad.text)
        End If
    DragMouse = 0
    Call CambiarCursor(frmBancoObj)
    Me.itemimg.Visible = False
    End If
End If
End Sub




Private Sub BovInventario_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If BintemElegido > 0 And BintemElegido <= MAX_INVENTORY_SLOTS Then
If Button = 2 And UserStats(SlotStats).UserEstado = 0 And UserInventory(BintemElegido).GrhIndex > 0 Then
Call CambiarCursor(Me, 1)
ItemDragued = 100 + BintemElegido
'ExtractData App.Path & "\Graficos\Graficos.tds", GPdataBMP(GrhData(UserInventory(BintemElegido).GrhIndex).FileNum).Offset, GPdataBMP(GrhData(UserInventory(BintemElegido).GrhIndex).FileNum).FileSizeBMP

Set itemimg.Picture = clsEnpaquetado_LeerIPicture(pakGraficos, GrhData(UserInventory(BintemElegido).GrhIndex).filenum)

DragMouse = 1
itemimg.Visible = True
End If
End If
End Sub

Private Sub BovInventario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragMouse = 1 Then
Call itemimg.Move(x + 120, y + 190)
End If
End Sub

Private Sub BovInventario_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If x > 0 And y > 0 And x < Me.Bovinventario.ScaleWidth And y < Me.Bovinventario.ScaleHeight Then
    BX = x \ 33 + 1
    by = y \ 33 + 1
    BintemElegido = (BX + (by - 1) * 6)
    If BintemElegido <= MAX_INVENTORY_SLOTS Then
        If UserInventory(BintemElegido).GrhIndex > 0 Then
            'Call Dibujar(CInt(BintemElegido), Me.Bovinventario, UserInventory, 6)
            Me.Inventariolbl(0) = UserInventory(BintemElegido).Name
            Me.Inventariolbl(1) = UserInventory(BintemElegido).MinDef
            Me.Inventariolbl(2) = UserInventory(BintemElegido).MaxHit
            Me.Inventariolbl(3) = UserInventory(BintemElegido).MinHit
            Me.Inventariolbl(0).ToolTipText = UserInventory(BintemElegido).Name
        End If
    End If
Else
Call BovBoveda_MouseUp(Button, 100, 0, 0)
End If

If Button = 2 Then
    If DragMouse = 1 Then
        If ItemDragued > 100 Then
        ItemDragued = ItemDragued - 100
            If Not ItemDragued = BintemElegido Then
            EnviarPaquete ChangeItemsSlot, Chr$(BintemElegido) & Chr$(ItemDragued)
            End If
        Else
        EnviarPaquete Paquetes.Retirar, Chr(BBitemElegido) & Codify(cantidad.text)
        End If
    DragMouse = 0
    Call CambiarCursor(frmBancoObj)
    Me.itemimg.Visible = False
    End If
End If

End Sub

Private Sub cantidad_Change()
If val(cantidad.text) < 0 Then
    cantidad.text = 1
End If

If val(cantidad.text) > MAX_INVENTORY_OBJS Then
    cantidad.text = 1
End If
cantidad.text = val(cantidad.text)
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_Initialize()
    ReDim UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmBancoObj)

DameImagenForm Me, 93
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Label1(Selecionado).tag = "1" Then
    Label1(Selecionado).tag = "0"
    Label1(Selecionado).Picture = Nothing
End If

End Sub

Private Sub Label1_Click(Index As Integer)
cantidad.text = val(cantidad.text)
Select Case Index
    Case 0
    EnviarPaquete Paquetes.BancoOk, ""
    Case 1
    If cantidad.text = 0 Or BBitemElegido <= 0 Then Exit Sub
    EnviarPaquete Paquetes.Retirar, Chr$(BBitemElegido) & Codify(cantidad.text)
    Case 2
    If cantidad.text = 0 Or BintemElegido <= 0 Then Exit Sub
    EnviarPaquete Paquetes.Depositar, Chr$((BintemElegido)) & Codify(cantidad.text)
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selecionado <> Index Then
        Label1(Selecionado).tag = "0"
        Label1(Selecionado).Picture = Nothing
    End If
    
    If Label1(Index).tag <> "1" Then
        Label1(Index).tag = 1
        Call DameImagen(Label1(Index), Index + 1)
        Selecionado = Index
    End If
End Sub

