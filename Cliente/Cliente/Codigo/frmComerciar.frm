VERSION 5.00
Object = "{50CBA22D-9024-11D1-AD8F-8E94A5273767}#8.7#0"; "TranImg2.ocx"
Begin VB.Form frmComerciar 
   BackColor       =   &H001D4A78&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DevPowerTransImg.TransImg itemimg 
      Height          =   450
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   794
      BackColor       =   -2147483639
      MaskColor       =   0
      Transparent     =   -1  'True
   End
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Text            =   "1"
      Top             =   3120
      Width           =   720
   End
   Begin VB.PictureBox ComercioInventario 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   3435
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   3
      Top             =   375
      Width           =   2910
   End
   Begin VB.PictureBox NpcInventarioComercio 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   315
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   2
      Top             =   375
      Width           =   2910
   End
   Begin VB.Image Command2 
      Height          =   420
      Left            =   5670
      Top             =   3030
      Width           =   735
   End
   Begin VB.Image Boton 
      Height          =   420
      Index           =   1
      Left            =   4245
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Image Boton 
      Height          =   420
      Index           =   0
      Left            =   1125
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Haz click en un item para mas información."
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
      Left            =   360
      TabIndex        =   6
      Top             =   75
      Width           =   3675
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   2925
      Width           =   705
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemElegidoV As Byte
Private ItemElegidoC  As Byte
Private DragMouse As Byte

Private NPCInventory() As Inventory

Friend Sub setInventario(ByRef inventario() As Inventory)
    NPCInventory = inventario
End Sub

Friend Sub setNpcSlot(Slot As Byte, Inventory As Inventory)
    NPCInventory(Slot) = Inventory
    
    Call actualizarSiEsElSeleccionadoCompra(Slot)
End Sub
Public Sub setPrecio(Slot As Byte, precio As Long)

    NPCInventory(Slot).valor = precio
    
    Call actualizarSiEsElSeleccionadoCompra(Slot)
End Sub

Public Sub refrescar()
    Call DrawInvSimple(Me.NpcInventarioComercio, NPCInventory, ItemElegidoC)
    Call DrawInvSimple(Me.ComercioInventario, UserInventory, ItemElegidoV)
End Sub

Private Sub actualizarSiEsElSeleccionadoCompra(Slot As Byte)
    If ItemElegidoC = Slot Then
        Call mostrarInfoObjeto(NPCInventory(Slot))
    End If
End Sub

Private Sub mostrarInfoObjeto(objeto As Inventory)
    frmComerciar.Label3.Caption = objeto.Name & " " & "Def: " & objeto.MinDef & "/" & objeto.MaxDef & " Hit: " & objeto.MinHit & "/" & objeto.MaxHit & " Valor: " & objeto.valor
End Sub
Private Sub Boton_Click(Index As Integer)

    Call Sonido_Play(SND_CLICK)
    
    Select Case Index
      Case 0
        If ItemElegidoC = 0 Or val(cantidad.text) = 0 Then
            Exit Sub '---> Bottom
        End If
        If NPCInventory(ItemElegidoC).GrhIndex <= 0 Then
            Exit Sub '---> Bottom
        End If
        If UserGLD >= NPCInventory(ItemElegidoC).valor * val(cantidad) Then
            EnviarPaquete Paquetes.comprar, Chr$(ItemElegidoC) & Codify(val(cantidad.text))
          Else
            AddtoRichTextBox frmConsola.ConsolaFlotante, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
      Case 1
        If ItemElegidoV > MAX_INVENTORY_SLOTS Then
            Exit Sub
        End If
        If ItemElegidoV = 0 Or val(cantidad.text) = 0 Then
            Exit Sub '---> Bottom
        End If
        If UserInventory(ItemElegidoV).GrhIndex <= 0 Then
            Exit Sub '---> Bottom
        End If
        If UserInventory(ItemElegidoV).Equipped = 0 Then
            EnviarPaquete Paquetes.Vender, Chr$(ItemElegidoV) & Codify(val(cantidad.text))
          Else
            AddtoRichTextBox frmConsola.ConsolaFlotante, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub '---> Bottom
        End If
    End Select

End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Boton(Index).tag <> "1" Then
        Boton(Index).tag = 1
        Call DameImagen(Boton(Index), Index + 5)
    End If

End Sub

Private Sub cantidad_Change()

    If val(cantidad.text) < 0 Then
        cantidad.text = 1
    End If
    
    If val(cantidad.text) > MAX_INVENTORY_OBJS Then
        cantidad.text = 1
    End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub ComercioInventario_Click()

    If UserInventory(ItemElegidoV).GrhIndex > 0 Then
        Call mostrarInfoObjeto(UserInventory(ItemElegidoV))
    End If
      
End Sub

Private Sub ComercioInventario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ItemElegidoV > 0 And ItemElegidoV <= MAX_INVENTORY_SLOTS Then
        If Button = 2 And UserStats(SlotStats).UserEstado = 0 And UserInventory(ItemElegidoV).GrhIndex > 0 Then
            Call CambiarCursor(Me, 1) ':( Remove "Call" verb and brackets
            ItemDragued = 100 + ItemElegidoV

            Set itemimg.Picture = clsEnpaquetado_LeerIPicture(pakGraficos, GrhData(UserInventory(ItemElegidoV).GrhIndex).filenum)

            DragMouse = 1
            itemimg.Visible = True
        End If
    End If

End Sub

Private Sub ComercioInventario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragMouse = 1 Then
        Call itemimg.Move(X + 210, Y) ':( Remove "Call" verb and brackets
    End If
End Sub

Private Sub ComercioInventario_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X < 0 Or Y < 0 Or X >= Me.ComercioInventario.ScaleWidth Or Y >= Me.ComercioInventario.ScaleHeight Then
        Call NpcInventarioComercio_MouseUp(Button, 100, 0, 0) ':( Remove "Call" verb and brackets
        Exit Sub
    End If
    
    Dim itemClickeado As Integer
    itemClickeado = DameItemClickeado(X, Y)
    
    If itemClickeado > MAX_INVENTORY_SLOTS Then
        Exit Sub
    End If
    
    ItemElegidoV = itemClickeado
    
    If Button = 2 Then
        If DragMouse = 1 Then
            If ItemDragued > 100 Then
                ItemDragued = ItemDragued - 100
                If Not ItemDragued = ItemElegidoV Then
                    EnviarPaquete ChangeItemsSlot, Chr$(ItemElegidoV) & Chr$(ItemDragued)
                End If
              Else 'NOT ITEMDRAGUED...
                EnviarPaquete Paquetes.comprar, Chr$(ItemElegidoC) & Codify(cantidad.text)
            End If
            DragMouse = 0
            Call CambiarCursor(frmComerciar) ':( Remove "Call" verb and brackets
            Me.itemimg.Visible = False
        End If
    End If

End Sub

Private Sub Command2_Click()
    EnviarPaquete Paquetes.ComOk
    frmMain.Enabled = True
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Command2.tag <> "1" Then
        Call DameImagen(Command2, 7) ':( Remove "Call" verb and brackets
        Command2.tag = "1"
    End If

End Sub

Private Sub Form_Load()

  'Cargamos la interfase

    DameImagenForm Me, 98
    'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.jpg")
    'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.jpg")
    frmMain.Enabled = False
    Call CambiarCursor(frmComerciar) ':( Remove "Call" verb and brackets

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Boton(0).tag = "1" Then
        Boton(0).Picture = Nothing
        Boton(0).tag = "0"
      ElseIf Boton(1).tag = "1" Then 'NOT BOTON(0).TAG...
        Boton(1).Picture = Nothing
        Boton(1).tag = "0"
      ElseIf Command2.tag = "1" Then 'NOT BOTON(1).TAG...
        Command2.Picture = Nothing
        Command2.tag = "0"
    End If

End Sub

Private Function DameItemClickeado(X As Single, Y As Single) As Integer

    X = X \ 32 + 1
    Y = Y \ 32 + 1
    DameItemClickeado = (X + (Y - 1) * 6)

End Function

Private Sub NpcInventarioComercio_Click()
    If ItemElegidoC = 0 Then Exit Sub
    If NPCInventory(ItemElegidoC).GrhIndex > 0 Then
        Call mostrarInfoObjeto(NPCInventory(ItemElegidoC))
    End If
End Sub

Private Sub NpcInventarioComercio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ItemElegidoC > 0 And ItemElegidoC <= MAX_NPC_INVENTORY_SLOTS Then
        If Button = 2 And UserStats(SlotStats).UserEstado = 0 And NPCInventory(ItemElegidoC).GrhIndex > 0 Then
            Call CambiarCursor(Me, 1) ':( Remove "Call" verb and brackets
            ItemDragued = ItemElegidoC

            Set itemimg.Picture = clsEnpaquetado_LeerIPicture(pakGraficos, GrhData(NPCInventory(ItemElegidoC).GrhIndex).filenum)

            DragMouse = 1
            itemimg.Visible = True
        End If
    End If

End Sub

Private Sub NpcInventarioComercio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If DragMouse = 1 Then
        Call itemimg.Move(X, Y) ':( Remove "Call" verb and brackets
    End If
End Sub

Private Sub NpcInventarioComercio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X < 0 Or Y < 0 Or X >= Me.NpcInventarioComercio.ScaleWidth Or Y >= Me.NpcInventarioComercio.ScaleHeight Then
        Call ComercioInventario_MouseUp(Button, Shift, 1, 1): Exit Sub ':( Remove "Call" verb
        Exit Sub
    End If
    
    Dim itemClickeado As Integer
    itemClickeado = DameItemClickeado(X, Y)
    
    If itemClickeado > MAX_INVENTORY_SLOTS_NPC Then
        Exit Sub
    End If
    
    ItemElegidoC = itemClickeado
    
    If Button = 2 Then
        If DragMouse = 1 Then
            If ItemDragued < 100 Then
                If Not ItemDragued = ItemElegidoC Then
                    'EnviarPaquete ChangeItemsSlotboveda, Chr$(ItemElegidoC) & Chr$(ItemDragued)
                End If
              Else 'NOT ITEMDRAGUED...
                ItemDragued = ItemDragued - 100
                EnviarPaquete Paquetes.Vender, Chr$(ItemElegidoV) & Codify(cantidad.text)
            End If
            DragMouse = 0
            Call CambiarCursor(frmComerciar) ':( Remove "Call" verb and brackets
            Me.itemimg.Visible = False
        End If
    End If

End Sub
