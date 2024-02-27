VERSION 5.00
Begin VB.Form frmPakRollback 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RollBack de archivos"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frmPakRollback.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "NumeroDeArchivo y PAK"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdExtraerTodos 
         Caption         =   "Extraer todos"
         Height          =   480
         Left            =   4800
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdExtraer 
         Caption         =   "Extraer..."
         Height          =   480
         Left            =   4800
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   4800
         TabIndex        =   3
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmdRollBack 
         Caption         =   "Usar versión seleccionada"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4800
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.ListBox lstVersiones 
         Height          =   4740
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmPakRollback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumeroArchivoSeleccionado As Integer
Public Pak As clsEnpaquetado

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExtraer_Click()
    Dim VersionSeleccionada As Integer
    
    VersionSeleccionada = val(Split(lstVersiones.list(lstVersiones.ListIndex), vbTab)(0))
    
    If Pak.ExtraerVersion(NumeroArchivoSeleccionado, VersionSeleccionada, OPath) Then
        MsgBox "Recurso extraido en '" & OPath & "'.'", vbInformation, Me.caption
    Else
        MsgBox "Ocurrio un error al intentar extraer los archivos", vbExclamation, Me.caption
    End If
End Sub

Private Sub cmdExtraerTodos_Click()
    Dim CarpetaSalida As String
    CarpetaSalida = modFolderBrowse.Seleccionar_Carpeta("Seleccione la carpeta donde se van a extraer todas las versiones", OPath)
    
    If right$(CarpetaSalida, 1) <> "\" Then CarpetaSalida = CarpetaSalida & "\"
    
    If FolderExist(CarpetaSalida) Then
        If Not Pak.ExtraerVersiones(NumeroArchivoSeleccionado, CarpetaSalida) Then
            MsgBox "Ocurrio un error al intentar extraer los archivos"
        End If
    Else
        MsgBox "Carpeta invalida"
    End If
End Sub

