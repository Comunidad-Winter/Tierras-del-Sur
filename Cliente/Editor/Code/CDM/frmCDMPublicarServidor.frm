VERSION 5.00
Begin VB.Form frmCDMPublicarServidor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tierras del Sur - Publcar en el servidor"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCDMPublicarServidor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstCambios 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1800
      Width           =   5175
   End
   Begin VB.CommandButton cmdCompartir 
      Caption         =   "Compartir"
      Height          =   360
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      Top             =   3840
      Width           =   1230
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label lblSeleccione 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccioná los elementos modificados que queres compartir"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4230
   End
   Begin VB.Label lblIngresaComentario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingresá un comentario acerca de las novedades que vas a compartir."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   4980
   End
End
Attribute VB_Name = "frmCDMPublicarServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
