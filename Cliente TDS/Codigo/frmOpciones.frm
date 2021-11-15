VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4935
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Informacion"
      Height          =   3495
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   2535
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Rankings (http://rankings.aotds.com.ar"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Estadisticas (http://est.aotds.com.ar"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Soporte (http://soporte.aotds.com.ar"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Manual (http://manual.aotds.com.ar"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Foro (http://foro.aotds.com.ar"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Web (http://www.aotds.com.ar"
         Height          =   375
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Sonido"
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Invertir3D"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Musica desactivada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "FX desactivados"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   1695
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   2990
         _Version        =   393216
         Orientation     =   1
         Max             =   4000
         TickStyle       =   2
         TickFrequency   =   500
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Silencio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   1695
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   2990
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   500
         SmallChange     =   500
         Max             =   4000
         TickStyle       =   2
         TickFrequency   =   500
         TextPosition    =   1
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Musica"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectos"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Private Sub Check1_Click()
If Me.Check1.value = 1 Then
Me.Slider1.value = 4000
Me.Slider2.value = 4000
Me.Check3.value = 1
Me.Check2.value = 1
IMC.Stop
Else
 If bLluvia(UserMap) = 0 Then
    If bRain Then
        IMC.Stop
        End If
    End If
Me.Check2.value = 0
Me.Check3.value = 0
End If
End Sub

Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub Check2_Click()
If Me.Check2.value = 1 Then
Fx = 1
Else
Fx = 0
End If

End Sub

Private Sub Check3_Click()
If Me.Check3.value = 1 Then
            Musica = 1
            Stop_Midi
  Else
        If Musica = 0 Then Exit Sub
            Musica = 0
            'frmMain.Winsock1.SendData "A" & "2.mid"
            Play_Midi
End If
End Sub

Private Sub Command2_Click()
VolumeN = frmOpciones.Slider1.value
If VolumeN = 1 Then VolumeN = 0
If frmOpciones.Check1 Then VolumeN = 5000
Perf.SetMasterVolume -Me.Slider2.value
Me.Visible = False
Call PlayWaveDS(SND_CLICK)
End Sub

Private Sub Command3_Click()
#If ConAlfaB = 1 Then
bNoche = Not bNoche
SurfaceDB.EfectoPred = IIf(bNoche, 1, 0)
SurfaceDB.BorrarTodo
#Else
MsgBox "Que hacés ?"
#End If
End Sub

Private Sub Form_Load()
If Musica = 1 Then
    frmOpciones.Check3 = 1
Else
    frmOpciones.Check3 = 0
End If
If Fx = 0 Then
      frmOpciones.Check2 = 0
Else
      frmOpciones.Check2 = 1
End If
End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
Randomize Timer
RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function

Private Sub Slider1_Change()
If Me.Slider1.value <= 0 Then
Me.Slider1.Text = "Normal"
Me.Check1.value = 0
ElseIf Me.Slider1.value > 0 And Me.Slider1.value < 2000 Then
Me.Slider1.Text = "Bajo"
Me.Check1.value = 0
ElseIf Me.Slider1.value > 2000 And Me.Slider1.value < 4000 Then
Me.Slider1.Text = "Muy Bajo"
Me.Check1.value = 0
Else
Me.Slider1.Text = "Silencio"
End If

End Sub

Private Sub Slider2_Change()
If Me.Slider1.value >= 0 Then
Me.Slider2.Text = "Normal"
Me.Check1.value = 0
ElseIf Me.Slider1.value < 0 And Me.Slider1.value < 2000 Then
Me.Slider2.Text = "Bajo"
Me.Check1.value = 0
ElseIf Me.Slider1.value > 2000 And Me.Slider1.value < 4000 Then
Me.Slider2.Text = "Muy Bajo"
Me.Check1.value = 0
Else
Me.Slider2.Text = "Silencio"
End If

Perf.SetMasterVolume -Me.Slider2.value
End Sub

