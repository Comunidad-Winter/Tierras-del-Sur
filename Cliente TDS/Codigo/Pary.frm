VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Partym 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Rechazar"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      _Version        =   393217
      BackColor       =   12632256
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Pary.frx":0000
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   1980
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   1980
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CrearParty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
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
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Expulsar"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir Party"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Integrante:    Experiencia"
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   3735
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes:"
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
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Integrantes:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Partym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Dim I As Integer

'echo por marche o pom :)
Private Sub Command1_Click()
If UserEstado = 0 Then
Call SendData("/CREARPARTY")
Me.Command2.enabled = True
Me.Command2.Visible = True
Me.Command1.Visible = False
Me.Command6.enabled = False
Me.Command7.enabled = False
Me.List1.enabled = True
Me.List2.enabled = True
Me.Label6.Visible = True
Me.List1.Visible = True
Me.RecTxt.Visible = False
Me.Label4.Visible = False
Me.Label3.Visible = True
Me.Label2.Visible = True
Call SendData("/CPARTY")
Else
AddtoRichTextBox frmMain.RecTxt, "Estas Muerto!!.", 255, 0, 0, True, False, False
End If
End Sub

Private Sub Command2_Click()
Call SendData("/SALIRPARTY")
Me.Command2.enabled = False
Me.Command2.Visible = False
Me.Command1.Visible = True
Me.Command6.enabled = False
Me.Command7.enabled = False
Me.List1.enabled = False
Me.List2.enabled = False
Me.List1.Visible = False
'Me.RichTextBox1.Visible = True
Me.Label4.Visible = True
Me.Label3.Visible = False
Me.Label2.Visible = True
Me.Command7.enabled = False
Me.Command6.enabled = False
gh = False
ss = False
Unload Me
End Sub

Private Sub Command3_Click()

For I = 0 To 20
If Listasolicitudes(I) = Me.List1 Then
Listasolicitudes(I) = ""
Exit For
End If
Next I
Me.List1.RemoveItem Me.List1.ListIndex

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Call SendData("/EP " & Me.List2)
For I = 0 To 20
If Listaintegrantes(I) <> "" Then
Listaintegrantes(I) = ""
End If
Next I

Me.List2.RemoveItem Me.List2.ListIndex
Me.Command6.enabled = True
End Sub

Private Sub Command7_Click()
Call SendData("/AP " & Me.List1)
For I = 0 To 20
If Listasolicitudes(I) = Me.List1 Then
Listasolicitudes(I) = ""
Exit For
End If
Next I
        
For I = 0 To 20
If Listaintegrantes(I) = "" Then
Listaintegrantes(I) = Me.List1
Exit For
End If
Next
Me.List2.AddItem Me.List1

Me.List1.RemoveItem Me.List1.ListIndex
Me.Command7.enabled = False
Me.Command3.enabled = False
End Sub

Private Sub Form_Load()
Dim aa As Integer
'lo q va a ver la persona
        If ss Then ' si es lider del party
               
               'Pide lso integranes y las solicitudes.
                For I = 0 To 20
                If Listaintegrantes(I) = "" Then
                Exit For
                Else
                Me.List2.AddItem Listaintegrantes(I)
                End If
                Next I
                
                For I = 0 To 20
                If Listasolicitudes(I) = "" Then
                Exit For
                Else
                Me.List1.AddItem Listasolicitudes(I)
                End If
                Next I
              
                Me.Command2.enabled = True
                Me.Command2.Visible = True
                Me.Command1.Visible = False
                Me.Command6.enabled = False
                Me.Label6.Visible = True
                Me.Command7.enabled = False
                Me.Command3.enabled = False
                Me.List1.enabled = True
                Me.List2.enabled = True
                Me.List1.Visible = True
                Me.RecTxt.Visible = False
                Me.Label4.Visible = False
                Me.Label3.Visible = True
                Me.Label2.Visible = True
                Else
                    If gh Then ' si participa en una party
                    Me.Command2.enabled = True
                    Me.Command2.Visible = True
                    Me.Command1.Visible = False
                    Me.Command6.enabled = False
                    Me.Command7.enabled = False
                    Me.Label6.Visible = False
                    Me.Command3.enabled = False
                    Me.List1.enabled = False
                    Me.List2.enabled = False
                    Me.List1.Visible = False
                    Me.RecTxt.Visible = True
                    Me.Label4.Visible = True
                    Me.Label3.Visible = False
                    Me.Label2.Visible = False
                    Else  ' si abre la venta para crear una party
                        Me.Label6.Visible = False
                        Me.Command2.enabled = False
                        Me.Command2.Visible = False
                        Me.Command1.Visible = True
                        Me.Command6.enabled = False
                        Me.Command7.enabled = False
                        Me.Command3.enabled = False
                        Me.List1.enabled = False
                        Me.List2.enabled = False
                        Me.List1.Visible = False
                        Me.RecTxt.Visible = True
                        Me.Label4.Visible = True
                        Me.Label3.Visible = False
                        Me.Label2.Visible = False
                   End If
              End If
End Sub

Private Sub Label5_Click()
Call SendData("/SALIRPARTY")
Me.Command2.enabled = False
Me.Command2.Visible = False
Me.Command1.Visible = True
Me.Command6.enabled = False
Me.Command7.enabled = False
Me.List1.enabled = False
Me.List2.enabled = False
Me.List1.Visible = False
'Me.RichTextBox1.Visible = True
Me.Label4.Visible = True
Me.Label3.Visible = False
Me.Label2.Visible = True
Me.Command7.enabled = False
Me.Command3.enabled = False
Me.Command6.enabled = False
gh = False
ss = False
Unload Me
End Sub

Private Sub Label6_Click()
If Me.RecTxt.Visible Then
Me.RecTxt.Visible = False
Me.Label4.Visible = False
Me.Label6.Caption = ">>"
Me.Label3.Visible = True
Else
Me.Label6.Caption = "<<"
Me.Label4.Visible = True
Me.RecTxt.Visible = True
Me.Label3.Visible = False
End If
End Sub

Private Sub List1_Click()
Me.Command7.enabled = True
Me.Command3.enabled = True
End Sub

Private Sub List2_Click()
Me.Command6.enabled = True
End Sub
'********************Misery_Ezequiel 28/05/05********************'
Private Sub RecTxt_Change()

End Sub
