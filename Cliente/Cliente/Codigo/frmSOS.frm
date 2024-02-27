VERSION 5.00
Begin VB.Form frmSOS 
   BorderStyle     =   0  'None
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRespuesta 
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   290
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtConsulta 
      Height          =   2295
      Left            =   290
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   290
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Image Boton 
      Height          =   480
      Index           =   1
      Left            =   232
      Top             =   4790
      Width           =   1230
   End
   Begin VB.Image Boton 
      Height          =   480
      Index           =   0
      Left            =   2360
      Top             =   4790
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSOS.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   290
      TabIndex        =   6
      Top             =   360
      Width           =   3285
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label labelRespuesta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   290
      TabIndex        =   5
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label labelNueva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   290
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pregunta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   290
      TabIndex        =   3
      Top             =   1560
      Width           =   780
   End
End
Attribute VB_Name = "frmSOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Selecionado As Byte

Private Sub Boton_Click(Index As Integer)

    Select Case Index
        Case 0 ' Cancelar
            Unload Me
        Case 1
            botonEnviar
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
        Call DameImagen(Boton(Index), 161 + Index)
    End If
End Sub
     
Private Sub Combo1_Click()
If Combo1.text = "Nueva Pregunta" Then
   txtConsulta.Visible = True
   txtRespuesta.Visible = False
   
   txtConsulta.text = "Ingrese su consulta. Sea lo más claro posible. A la brevedad recibirá la respuesta en su correo electronico..."
   txtConsulta.ForeColor = &H808080
   
   labelRespuesta.Visible = False
   labelNueva.Visible = True
Else
    txtRespuesta.text = obtenerRespuestaSOS(Combo1.ListIndex)
    
    labelNueva.Visible = False
    labelRespuesta.Visible = True
    
    txtConsulta.Visible = False
    txtRespuesta.Visible = True
End If
End Sub

Private Sub botonEnviar()
    If Combo1.text = "Nueva Pregunta" Then
        If txtConsulta.ForeColor = &H808080 Or txtConsulta.text = "" Then
            MsgBox "Por favor ingrese su consulta."
        ElseIf Len(txtConsulta.text) < 10 Then
            MsgBox "Debe ingresar al menos 10 caracteres en su consulta."
        ElseIf Len(txtConsulta.text) > 250 Then
            MsgBox "El máximo de caracteres que puede tener tu consulta es de 250."
        Else
            'Esta todo ok, mando la consulta
            Call sSendData(Paquetes.Comandos, Simple.GM, txtConsulta.text)
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Un GM ha recibido tu consulta. Recibiras en tu mail la respuesta en breve.", 65, 190, 156, 0)
        
            Unload Me
        End If
    Else
        Combo1.ListIndex = Combo1.ListCount - 1
    End If
End Sub
Private Sub Form_Load()
    Dim k As Integer
     
    Call CambiarCursor(frmSOS)
    DameImagenForm Me, 160
    
    If cargarRespuestasSOS Then
        For k = 0 To UBound(CLI_RespuestasSOS.respuestasSOS)
            Combo1.AddItem CLI_RespuestasSOS.respuestasSOS(k).Titulo
        Next
    End If
    Combo1.AddItem "Nueva Pregunta"
   
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End Sub

Private Sub txtConsulta_Click()
   txtConsulta.text = ""
   txtConsulta.ForeColor = &H0&
End Sub
