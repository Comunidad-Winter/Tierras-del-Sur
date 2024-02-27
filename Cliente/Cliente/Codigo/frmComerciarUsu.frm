VERSION 5.00
Begin VB.Form frmComerciarUsu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   120
      Width           =   480
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8EAEC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   6000
      TabIndex        =   10
      Top             =   1200
      Width           =   2610
   End
   Begin VB.TextBox txtCant 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7380
      TabIndex        =   9
      Text            =   "1"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton optQue 
      BackColor       =   &H00000000&
      Caption         =   "Objeto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   6000
      TabIndex        =   8
      Top             =   960
      Value           =   -1  'True
      Width           =   795
   End
   Begin VB.OptionButton optQue 
      BackColor       =   &H00000000&
      Caption         =   "Oro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   7380
      TabIndex        =   7
      Top             =   960
      Width           =   570
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8EAEC&
      Height          =   3930
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   2610
   End
   Begin VB.ListBox Bolsa 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8EAEC&
      Height          =   3540
      Left            =   3360
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   5
      Left            =   435
      Top             =   5625
      Width           =   1215
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   4
      Left            =   1695
      Top             =   5625
      Width           =   1215
   End
   Begin VB.Image Boton 
      Height          =   495
      Index           =   3
      Left            =   3270
      Top             =   5295
      Width           =   1215
   End
   Begin VB.Image Boton 
      Height          =   495
      Index           =   2
      Left            =   4500
      Top             =   5310
      Width           =   1215
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   4800
      Top             =   105
      Width           =   1095
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   0
      Left            =   6705
      Top             =   5610
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   5325
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   5220
      Width           =   2535
   End
   Begin VB.Label OroLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   5760
      Width           =   2610
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando respuesta..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1755
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Selecionado As Byte
Private oroOfrecido As Long

Private Sub Boton_Click(Index As Integer)
Dim Hola As Variant
Dim i As Byte
Dim ItemsAcomerciar As String
Dim cadena As Variant
Dim ii As Integer

Select Case Index

Case 0
    oroOfrecido = 0
    
    txtCant.text = Int(val(txtCant.text))
    
    If optQue(0).value = True Then
        If List1.ListIndex < 0 Then Exit Sub
            If List1.itemData(List1.ListIndex) <= 0 Then Exit Sub
                If txtCant.text > List1.itemData(List1.ListIndex) Or _
                    txtCant.text <= 0 Then Exit Sub
            ElseIf optQue(1).value = True Then
                If val(txtCant.text) > UserGLD Then
                Exit Sub
                End If
    End If

    If optQue(0).value = True Then
        For i = 0 To Bolsa.ListCount
            If InStr(1, Bolsa.list(i), List1 & " - ") > 0 Then
            Hola = Split(Bolsa.list(i), " - ")
            If val(Hola(1)) + txtCant.text > 10000 Then
            Else
            Me.Bolsa.list(i) = Trim(Hola(0)) & " - " & txtCant.text
            Exit Sub
            End If
            End If
        Next i
            Me.Bolsa.AddItem (List1 & " - " & txtCant.text)
    ElseIf optQue(1).value = True Then
        oroOfrecido = val(txtCant.text)
        Me.OroLabel.Caption = FormatNumber$(oroOfrecido, 0, vbTrue, vbFalse, vbTrue)
    End If
Case 3
        Dim aux(1 To MAX_INVENTORY_SLOTS) As Boolean
        Dim slotElegido As Byte
        Dim minDif As Integer
        
        '¿Hay algo para ofrecer?
        If Me.Bolsa.ListCount < 1 And oroOfrecido <= 0 Then Exit Sub

        'Items para ofrecer?
        If Me.Bolsa.ListCount > 0 Then
            'Recorro la lista
            For i = 0 To Me.Bolsa.ListCount - 1
            'Obtengo el nombre del item
            cadena = Split(Me.Bolsa.list(i), " - ")
                'Recorro los items buscando el nombre del items en el inventario
                minDif = 10000
                For ii = 1 To MAX_INVENTORY_SLOTS
                    If UserInventory(ii).Name = Trim(cadena(0)) And aux(ii) = False Then
                        If val(cadena(1)) <= UserInventory(ii).Amount Then
                            'Obtengo la diferencia
                            If UserInventory(ii).Amount - val(cadena(1)) < minDif Then
                                minDif = UserInventory(ii).Amount - val(cadena(1))
                                slotElegido = ii
                                If minDif = 0 Then Exit For
                            End If
                        End If
                    End If
                Next ii
                
                If minDif = 10000 Then
                Me.Label4 = "No tienes los items que estas ofreciendo (" & cadena(0) & ")."
                Exit Sub
                Else
                    ItemsAcomerciar = ItemsAcomerciar & Chr$(slotElegido) & LongToString(val(cadena(1)))
                    aux(slotElegido) = True
                End If
            Next
        End If

        If oroOfrecido > 0 Then
            If oroOfrecido <= UserGLD Then
            ItemsAcomerciar = ItemsAcomerciar & Chr$(255) & LongToString(oroOfrecido)
            Else
            Me.Label4 = "No tienes esa cantidad de oro."
            Exit Sub
            End If
        End If

        EnviarPaquete Paquetes.OfrecerComUsu, ItemsAcomerciar
        Boton(2).Enabled = False
        Boton(3).Enabled = False
        Boton(0).Enabled = False
        lblEstadoResp.Visible = True
Case 2
        If Me.Bolsa.ListIndex > -1 Then
        Me.Bolsa.RemoveItem Me.Bolsa.ListIndex
        End If
Case 1
        EnviarPaquete Paquetes.FinComUsu
Case 4
        If List2.ListCount <= 0 Then Exit Sub
        EnviarPaquete Paquetes.RechazarComUsu
Case 5
        If List2.ListCount <= 0 Then Exit Sub
        EnviarPaquete Paquetes.ComUsuOk
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Selecionado <> Index Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
    
    If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Call DameImagen(Boton(Index), Index + 40)
    End If

End Sub

Private Sub Form_Load()
'Carga las imagenes...?
lblEstadoResp.Visible = False
DameImagenForm Me, 97
Call CambiarCursor(frmComerciarUsu)
End Sub

Private Sub Form_LostFocus()
Me.SetFocus
Picture1.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End Sub


Private Sub List2_Click()
If List2.ListIndex >= 0 Then
    Label3.Caption = "Cantidad: " & FormatNumber$(List2.itemData(List2.ListIndex), 0, vbTrue, vbFalse, vbTrue)
End If
End Sub

Private Sub optQue_Click(Index As Integer)
Select Case Index
Case 0
    List1.Enabled = True
Case 1
    List1.Enabled = False
End Select
End Sub

Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or _
        KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
    KeyCode = 0
End If
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
    KeyAscii = 0
End If
End Sub
