VERSION 5.00
Begin VB.Form frmAcercaDe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de Tierras del Sur Editor del Mundo"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcercaDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   3720
      TabIndex        =   2
      Top             =   5160
      Width           =   1470
   End
   Begin VB.Label lblContenidoGrafico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contenido Gráfico: Juan Castagna, Arkadiusz Zygarlicki."
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   4035
   End
   Begin VB.Label lblVersionEngine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-- version motor grafico --"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   2580
   End
   Begin VB.Label lblProgramación 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programación: Marcelo Marcón, Agustín Mendez. Colaboradores: Leandro Mendoza."
      Height          =   390
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4710
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCreditos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Créditos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label lblVersionContenido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-- version contenido ---"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2145
   End
   Begin VB.Label lblProhibicion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAcercaDe.frx":1CCA
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4920
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAcercaDe.frx":1D87
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4965
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-- info version --"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   5100
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim tipoVersion As String
    Dim tipoDesarrollo As String

    #If Colaborativo = 0 Then
        tipoVersion = "No colaborativa"
    #Else
        tipoVersion = "Colaborativa"
    #End If
    
    #If Produccion = 0 Then
        tipoDesarrollo = "Para Pruebas"
    #ElseIf Produccion = 1 Then
        tipoDesarrollo = "Para Producción"
    #ElseIf Produccion = 2 Then
        tipoDesarrollo = "Para Pre-Produccion"
    #End If

    Me.lblVersion = "Software: Versión " & VERSION_EDITOR & " " & app.Major & "." & app.Minor & " Revisión " & app.Revision & ". " & tipoVersion & ". " & tipoDesarrollo & "."
    Me.lblVersionContenido = "Contenido: Versión " & CDM.cerebro.Version & "."
    
    Me.lblVersionEngine = "Motor Gráfico: Versión " & obtenerVersionEngine & "."
End Sub

Private Function obtenerVersionEngine() As String

Dim Version As String * 30
Dim partes() As String

' Aca si ponemos un error personalizado por si falla
On Error GoTo imposibleObtener

Call mzengine3lib.getCompilationDate(Version) ' MMM DD AAAA HH:mm:ii  -  MMM es formato de tres letras

Version = Replace$(Version, ":", " ")
partes = Split(Version, " ", 6)

obtenerVersionEngine = format$(mid$(partes(2), 3), "@@") & "." & format$(HelperDate.AbreviaturaANumero(partes(0)), "00") & _
                    format$(partes(1), "00") & "." & format$(partes(3), "00") & format$(partes(4), "00")
                  
Exit Function
imposibleObtener:
obtenerVersionEngine = "¡ERROR!"
                    
End Function
