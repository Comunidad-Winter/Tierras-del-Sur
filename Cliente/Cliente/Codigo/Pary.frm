VERSION 5.00
Begin VB.Form Partym 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   327
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   2175
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Pary.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   2505
      Left            =   840
      TabIndex        =   25
      Top             =   840
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   8
      X2              =   304
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   8
      X2              =   304
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   8
      X2              =   304
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   8
      X2              =   304
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00E0E0E0&
      X1              =   8
      X2              =   304
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   3720
      TabIndex        =   23
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   1920
      TabIndex        =   22
      Top             =   2400
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje1"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   21
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   20
      Top             =   2040
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   1920
      TabIndex        =   19
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje1"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   18
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   17
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   16
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje1"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   3720
      TabIndex        =   14
      Top             =   1320
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje1"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   3720
      TabIndex        =   11
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje1"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porcentaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   0
      Left            =   3720
      Top             =   7500
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image boton 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   3705
      Top             =   3210
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image boton 
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   2535
      Top             =   3210
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image boton 
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   120
      Top             =   3180
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image boton 
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   2520
      Top             =   3645
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image boton 
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   120
      Top             =   3645
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Image boton 
      Height          =   360
      Index           =   1
      Left            =   1440
      Top             =   3645
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Integrantes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "Partym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Marche 7/12/06
'Esto esta bastante desprolijo. Pero funciona :)
Option Explicit
Dim i As Integer
Public Selecionado As Byte

Private Const BTN_CREARPARTY = 1
Private Const BTN_SALIRPARTY = 2
Private Const BTN_RECHZAR = 3
Private Const BTN_CAMBIARPORCENTAJE = 5
Private Const BTN_EXPULSAR = 6
Private Const BTN_APROBARINGRESO = 7

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case BTN_CREARPARTY
        If UserStats(SlotStats).UserEstado = 0 Then
            EnviarPaquete Paquetes.Crearparty
        Else
            AddtoRichTextBox frmConsola.ConsolaFlotante, "Estas Muerto!!.", 255, 0, 0, True, False, False
        End If
Case BTN_SALIRPARTY
        EnviarPaquete Paquetes.Salirparty
        Me.Boton(BTN_SALIRPARTY).Enabled = False
        Me.Boton(BTN_SALIRPARTY).Visible = False
        Me.Boton(BTN_CREARPARTY).Visible = True
        Me.Boton(BTN_EXPULSAR).Enabled = False
        Me.Boton(BTN_APROBARINGRESO).Enabled = False
        Me.List1.Enabled = False
        Me.List2.Enabled = False
        Me.List1.Visible = False
        Me.Label4.Visible = True
        Me.Label3.Visible = False
        Me.Label2.Visible = True
        Me.Boton(BTN_APROBARINGRESO).Enabled = False
        Me.Boton(BTN_EXPULSAR).Enabled = False
        gh = False
        Liderparty = False
        Unload Me
Case BTN_RECHZAR

        If Me.List1.ListIndex = -1 Then Exit Sub
        For i = 0 To 20
            If Listasolicitudes(i) = Me.List1 Then
            Listasolicitudes(i) = ""
            Exit For
            End If
        Next i

        Me.List1.RemoveItem Me.List1.ListIndex
        
        Me.List1.ListIndex = -1
        
        Me.Boton(BTN_RECHZAR).Enabled = False
        Me.Boton(BTN_APROBARINGRESO).Enabled = False
        
        Call DameImagen(Boton(BTN_RECHZAR), 28)
        Call DameImagen(Boton(BTN_APROBARINGRESO), 25)
Case BTN_CAMBIARPORCENTAJE
        Call sSendData(1, Simple.PartyPorcentaje)
Case BTN_EXPULSAR
        If Me.List2.ListIndex = -1 Then Exit Sub
        sSendData Paquetes.Expulsarparty, 0, Me.List2
        
        
        For i = 0 To 20
        If Listaintegrantes(i) <> "" Then
        Listaintegrantes(i) = ""
        End If
        Next i

        Me.List2.RemoveItem Me.List2.ListIndex
        Me.Boton(BTN_EXPULSAR).Enabled = False
        Call DameImagen(Boton(BTN_EXPULSAR), 27)
Case BTN_APROBARINGRESO
        sSendData Paquetes.Aprobaringresoparty, 0, Me.List1
        If Me.List1.ListIndex = -1 Then Exit Sub
        
        For i = 0 To 20
            If Listasolicitudes(i) = Me.List1 Then
            Listasolicitudes(i) = ""
            Exit For
            End If
        Next i

        Me.List1.RemoveItem Me.List1.ListIndex
        
        Me.List1.ListIndex = -1
        
        Me.Boton(BTN_RECHZAR).Enabled = False
        Me.Boton(BTN_APROBARINGRESO).Enabled = False
        
        Call DameImagen(Boton(BTN_RECHZAR), 28)
        Call DameImagen(Boton(BTN_APROBARINGRESO), 25)
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Selecionado <> Index And Selecionado > 0 Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Boton(0).Picture
End If

If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Boton(0).Picture = Boton(Index).Picture
    Call DameImagen(Boton(Index), Index + 32)
End If
End Sub

Public Sub setVisibilidadCrearParty(Visible As Boolean)
    Me.lblInfo.Visible = Visible
    Me.Boton(BTN_CREARPARTY).Visible = Visible
    Me.Boton(BTN_CREARPARTY).Enabled = Visible
    
    If Visible Then Call DameImagen(Boton(BTN_CREARPARTY), 23)
End Sub

Public Sub refrescarPantalla()
        'lo q va a ver la persona
        If Liderparty Then ' si es lider del party
              
               Call setVisibilidadCrearParty(False)
               Call setVisibilidadExperiencia(False)
               Call setVisibleSolicitudes(True)
               
               'Pide lso integranes y las solicitudes.
                For i = 0 To 20
                    If Listaintegrantes(i) = "" Then
                    Exit For
                    Else
                    Me.List2.AddItem Listaintegrantes(i)
                    End If
                Next i
                
                For i = 0 To 20
                    If Listasolicitudes(i) = "" Then
                    Exit For
                    Else
                    Me.List1.AddItem Listasolicitudes(i)
                    End If
                Next i
                
                Me.Label6.Visible = True
                Me.Label6.Enabled = True
                
                Me.Boton(BTN_SALIRPARTY).Visible = True
                Me.Boton(BTN_SALIRPARTY).Enabled = True
                
                Me.Boton(BTN_CAMBIARPORCENTAJE).Visible = True
                Me.Boton(BTN_CAMBIARPORCENTAJE).Enabled = True
                
                Call DameImagen(Boton(BTN_CAMBIARPORCENTAJE), 29)
                Call DameImagen(Boton(BTN_SALIRPARTY), 24)
                
        Else
              If gh Then
              
                    Call setVisibilidadCrearParty(False)
                    Call setVisibilidadExperiencia(True)
                    Call setVisibleSolicitudes(False)
                
                    ' si participa en una party
                    Me.Boton(BTN_SALIRPARTY).Enabled = True
                    Me.Boton(BTN_SALIRPARTY).Visible = True
                                        
                    Me.Boton(BTN_SALIRPARTY).left = 105
                  
                    Call DameImagen(Boton(BTN_SALIRPARTY), 24)
              Else  ' si abre la venta para crear una party
              
                  Call setVisibilidadExperiencia(False)
                  Call setVisibleSolicitudes(False)
                  Call setVisibilidadCrearParty(True)

             End If
        End If
End Sub
Private Sub Form_Load()

Call CambiarCursor(Partym)
DameImagenForm Me, 110

Call refrescarPantalla
End Sub


Private Sub setVisibilidadExperiencia(Visible As Boolean)
    Label4.Visible = Visible
    Label9.Visible = Visible
    Label10.Visible = Visible
    
    Line1.Visible = Visible
    Line2.Visible = Visible
    Line3.Visible = Visible
    Line4.Visible = Visible
    Line5.Visible = Visible
    Label11.Visible = Visible

    Dim loopLInea As Byte
    
    For loopLInea = 0 To Label5.count - 1
        Label5(loopLInea).Visible = Visible
        Label7(loopLInea).Visible = Visible
        Label8(loopLInea).Visible = Visible
    Next loopLInea
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Selecionado > 0 Then
If Boton(Selecionado).tag = "1" Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Boton(0).Picture
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.SetFocus
Unload Me
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()
Unload Me
End Sub
Private Sub setVisibleSolicitudes(Visible As Boolean)
    Me.List1.Visible = Visible
    Me.List2.Visible = Visible
    Me.List1.Enabled = Visible
    Me.List2.Enabled = Visible
    
    Me.Label3.Visible = Visible
    Me.Label2.Visible = Visible
    
    If Visible Then
        Me.Boton(BTN_EXPULSAR).Enabled = False
        Me.Boton(BTN_APROBARINGRESO).Enabled = False
        Me.Boton(BTN_RECHZAR).Enabled = False
    End If
    
    Me.Boton(BTN_EXPULSAR).Visible = Visible
    Me.Boton(BTN_APROBARINGRESO).Visible = Visible
    Me.Boton(BTN_RECHZAR).Visible = Visible

    Call DameImagen(Boton(BTN_APROBARINGRESO), 25)
    Call DameImagen(Boton(BTN_RECHZAR), 28)
    Call DameImagen(Boton(BTN_EXPULSAR), 27)
End Sub

Private Sub Label6_Click()
If Me.List1.Visible = False Then
    Call setVisibilidadExperiencia(False)
    Call setVisibleSolicitudes(True)
    
    Me.Label6.Caption = ">>"
Else
    Me.Label6.Caption = "<<"
    Call setVisibilidadExperiencia(True)
    Call setVisibleSolicitudes(False)
End If
End Sub

Private Sub List1_Click()
Me.Boton(BTN_APROBARINGRESO).Enabled = True
Me.Boton(BTN_RECHZAR).Enabled = True

Call DameImagen(Boton(BTN_APROBARINGRESO), 31)
Call DameImagen(Boton(BTN_RECHZAR), 32)
End Sub

Private Sub List2_Click()
    Me.Boton(BTN_EXPULSAR).Enabled = True
    Call DameImagen(Boton(BTN_EXPULSAR), 30)
End Sub
