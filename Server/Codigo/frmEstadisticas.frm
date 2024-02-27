VERSION 5.00
Begin VB.Form frmEstadisticas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stats"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Label lblStat 
         Height          =   495
         Index           =   3
         Left            =   2400
         TabIndex        =   8
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblStat 
         Height          =   495
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblStat 
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblStat 
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "MAX Bytes Enviados x Seg:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "MAX Bytes Recibidos x Seg:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bytes Enviados x Seg:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bytes Recibidos x Seg:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
