VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creaci�n de un Clan"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Alineaci�n"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   3855
      Begin VB.CheckBox ChkCaos 
         Caption         =   "Legi�n Oscura"
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox ChkNeutro 
         Caption         =   "Neutro"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox ChkReal 
         Caption         =   "Armada Real"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      MouseIcon       =   "frmGuildFoundation.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmGuildFoundation.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Web site del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informaci�n b�sica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtClanName 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   $"frmGuildFoundation.frx":02A4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del clan:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Private Sub ChkCaos_Click()
If ChkNeutro.value = 1 Then ChkNeutro.value = 0
If ChkReal.value = 1 Then ChkReal.value = 0
CAlineacion = 3
End Sub

Private Sub ChkNeutro_Click()
If ChkCaos.value = 1 Then ChkCaos.value = 0
If ChkReal.value = 1 Then ChkReal.value = 0
CAlineacion = 1
End Sub

Private Sub ChkReal_Click()
If ChkCaos.value = 1 Then ChkCaos.value = 0
If ChkNeutro.value = 1 Then ChkNeutro.value = 0
CAlineacion = 2
End Sub

Private Sub Command1_Click()
ClanName = txtClanName
Site = Text2
If CAlineacion = 0 Then MsgBox "Debes seleccionar una alineacion para el clan": Exit Sub
Unload Me
frmGuildDetails.Show , Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
If Len(txtClanName.Text) <= 30 Then
    If Not AsciiValidos(txtClanName) Then
        MsgBox "Nombre invalido."
        Exit Sub
    End If
Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub
End If
End Sub
'********************Misery_Ezequiel 28/05/05********************'