VERSION 5.00
Begin VB.Form frmComplementarios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compelemtarios"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Propiedades de la textura"
      Height          =   4575
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox Check1 
         Caption         =   "ColorADD"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox bAdd 
         Caption         =   "BlendONE"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox comp4 
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
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Text            =   "comp4"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox comp3 
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
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Text            =   "comp3"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox comp2 
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
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   "comp2"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox comp1 
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
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Text            =   "comp1"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox comp_ver_solo_comp 
         Caption         =   "Vista previa"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Restaurar valores previos"
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
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label 
         Caption         =   "Bump-Map:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1725
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Gloss:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   2085
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Specular:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   1005
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Non-Specular:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1365
         Width           =   1455
      End
      Begin VB.Label lblfilename 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label 
         Caption         =   "Nombre del archivo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.ListBox ListGraficos 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmComplementarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp_comp1 As Integer
Dim tmp_comp2 As Integer
Dim tmp_comp3 As Integer
Dim tmp_comp4 As Integer

Public tmp_grh_selected As Integer

Dim TIH As INFOHEADER


Private Sub cmdCancel_Click()
cmdSave.Enabled = False
cmdCancel.Enabled = False
TIH.complemento_1 = tmp_comp1
TIH.complemento_2 = tmp_comp2
TIH.complemento_3 = tmp_comp3
TIH.complemento_4 = tmp_comp4
comp1.text = TIH.complemento_1
comp2.text = TIH.complemento_2
comp3.text = TIH.complemento_3
comp4.text = TIH.complemento_4
definir_complementarios tmp_grh_selected, TIH.complemento_1, TIH.complemento_2, TIH.complemento_3, TIH.complemento_4
End Sub

Private Sub cmdSave_Click()

If (tmp_grh_selected > 0) Then

    Actualizar_Comp
    
    comp1.text = TIH.complemento_1
    comp2.text = TIH.complemento_2
    comp3.text = TIH.complemento_3
    comp4.text = TIH.complemento_4

    If pakGraficos.IH_Mod(tmp_grh_selected, TIH) = True Then
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        MsgBox "Archivo guardado."
        'CDMCerebroDeMono.CDM_Commit pakGraficos.Cabezal_GetFilenameName(tmp_grh_selected), tmp_grh_selected, CDM_Upd_Grafico
    Else
        MsgBox "Error al guardar los complementarios."
    End If

End If


End Sub

Private Sub comp1_Change()
Actualizar_Comp
End Sub

Private Sub comp2_Change()
Actualizar_Comp
End Sub

Sub Actualizar_Comp()
cmdSave.Enabled = True
cmdCancel.Enabled = True
TIH.complemento_1 = val(comp1.text) Mod &H7FFF
TIH.complemento_2 = val(comp2.text) Mod &H7FFF
TIH.complemento_3 = val(comp3.text) Mod &H7FFF
TIH.complemento_4 = val(comp4.text) Mod &H7FFF
definir_complementarios tmp_grh_selected, TIH.complemento_1, TIH.complemento_2, TIH.complemento_3, TIH.complemento_4
End Sub

Private Sub comp1_LostFocus()
    comp1.text = val(comp1.text) Mod &H7FFF
End Sub
Private Sub comp2_LostFocus()
    comp2.text = val(comp2.text) Mod &H7FFF
End Sub

Private Sub comp3_LostFocus()
    comp3.text = val(comp3.text) Mod &H7FFF
End Sub

Private Sub comp4_LostFocus()
    comp4.text = val(comp4.text) Mod &H7FFF
End Sub

Private Sub comp3_Change()
    Actualizar_Comp
End Sub

Private Sub comp4_Change()
    Actualizar_Comp
End Sub

Private Sub Form_Load()
ListGraficos.Clear
pakGraficos.Add_To_Listbox_Permisos ListGraficos, -1, 0
End Sub

Private Sub ListGraficos_Click()

tmp_grh_selected = val(ListGraficos.text)

comp1.Enabled = False
comp2.Enabled = False
comp3.Enabled = False
comp4.Enabled = False
cmdSave.Enabled = False
cmdCancel.Enabled = False
    
If (tmp_grh_selected > 0) Then
    lblfilename.Caption = ListGraficos.text
    
    comp1.Enabled = True
    comp2.Enabled = True
    comp3.Enabled = True
    comp4.Enabled = True

    
    pakGraficos.IH_Get tmp_grh_selected, TIH
    
    comp1.text = TIH.complemento_1
    comp2.text = TIH.complemento_2
    comp3.text = TIH.complemento_3
    comp4.text = TIH.complemento_4
    
    tmp_comp1 = TIH.complemento_1
    tmp_comp2 = TIH.complemento_2
    tmp_comp3 = TIH.complemento_3
    tmp_comp4 = TIH.complemento_4
Else
    cmdCancel_Click
End If

End Sub
