VERSION 5.00
Begin VB.Form Retos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retos"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2805
   Icon            =   "Retos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reglas"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox Check3 
         Caption         =   "No vale usar elementales."
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "No vale usar estupidez."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No vale usar invisibilidad."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad de oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
   End
   Begin VB.CommandButton Retar 
      Caption         =   "Retar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Retos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub Retar_Click()
If Me.Text1 = "" Then
MsgBox "Por favor inserte la cantidad de oro"
Exit Sub
End If
Dim opcion1 As Integer
Dim opcion2 As Integer
Dim opcion3 As Integer
opcion1 = 1
opcion2 = 1
opcion3 = 1
If Me.Check1 = False Then
opcion3 = 0
End If
If Me.Check2 = False Then
opcion2 = 0
End If
If Me.Check3 = False Then
opcion1 = 0
End If
Call SendData("/r0tar " & opcion1 & opcion2 & opcion3 & Me.Text1)
Unload Me
End Sub
