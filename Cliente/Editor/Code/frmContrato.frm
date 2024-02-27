VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tierras del Sur - Contrato de Confidencialidad"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContrato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Acepto"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4680
      TabIndex        =   3
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   6960
      TabIndex        =   2
      Top             =   7560
      Width           =   2070
   End
   Begin RichTextLib.RichTextBox rtbContrato 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12091
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmContrato.frx":1CCA
   End
   Begin VB.CheckBox chkAcepto 
      Appearance      =   0  'Flat
      Caption         =   "Entiendo el contenido y los alcances del presente contrato y lo acepto en su totalidad."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   6615
   End
End
Attribute VB_Name = "frmContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_aceptoContrato As Boolean

Public Function aceptoContrato() As Boolean
    aceptoContrato = m_aceptoContrato
End Function

Private Sub chkAcepto_Click()

    If Me.chkAcepto.value = vbChecked Then
        Me.cmdAceptar.Enabled = True
    Else
        Me.cmdAceptar.Enabled = False
    End If
    
End Sub

Private Sub cmdAceptar_Click()

    m_aceptoContrato = True
    Hide
End Sub

Private Sub cmdSalir_Click()
    m_aceptoContrato = False
    Hide
End Sub

Private Sub Form_Load()
    m_aceptoContrato = False
End Sub
