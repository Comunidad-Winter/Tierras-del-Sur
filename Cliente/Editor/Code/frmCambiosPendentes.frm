VERSION 5.00
Begin VB.Form frmCambiosPendentes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambios Pendientes de actualizacion"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambiosPendentes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkElemento 
      Appearance      =   0  'Flat
      Caption         =   "Elemento 0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   480
      Left            =   2760
      TabIndex        =   2
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCambiosPendentes.frx":1CCA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4965
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCambiosPendentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualizar_Click()
    
    Dim loopElemento As Byte
    Dim sElemento As String
    Dim ok As Boolean
    
    
    cmdActualizar.Enabled = False
    cmdActualizar.caption = "Actualizando..."
'-----------------------------------------------------------------------------
    For loopElemento = chkElemento.LBound To chkElemento.UBound
        If chkElemento(loopElemento).value = 1 Then
            sElemento = chkElemento(loopElemento).caption
            
            Select Case sElemento
            
                Case "Graficos"
                    ok = Me_indexar_Graficos.compilar()
                Case "Pisos"
                    ok = Me_indexar_Pisos.compilar()
                Case "Cuerpos"
                    ok = Me_indexar_Cuerpos.compilar
                Case "Armas"
                    ok = Me_indexar_Armas.compilar
                Case "Escudos"
                    ok = Me_indexar_Escudos.compilar
                Case "Cabezas"
                    ok = Me_indexar_Cabezas.compilar
                Case "Cascos"
                    ok = Me_indexar_Cascos.compilar
                Case "Efectos"
                    ok = Me_indexar_Efectos.compilar
            End Select
            
            If ok Then
                chkElemento(loopElemento).ForeColor = vbGreen
                chkElemento(loopElemento).FontBold = True
                chkElemento(loopElemento).Enabled = False
                chkElemento(loopElemento).value = False

                chkElemento(loopElemento).caption = sElemento & "(OK)"
                        
                Call ME_ControlCambios.SetCambioActualizado(sElemento)
            Else
                chkElemento(loopElemento).ForeColor = vbRed
                chkElemento(loopElemento).FontBold = True
                chkElemento(loopElemento).caption = sElemento & "(Error)"
            End If
        End If
    Next
'-----------------------------------------------------------------------------
    If ME_ControlCambios.pendientes.count > 0 Then
        cmdActualizar.Enabled = True
    End If
    cmdActualizar.caption = "Actualizar"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim elemento As Variant
    Dim loopElemento As Byte
    
    loopElemento = 0
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Graficos")
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Pisos")
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Cuerpos")
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Armas")
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Escudos")
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Cabezas")
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Cascos")
                    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Efectos")
                       
    For Each elemento In ME_ControlCambios.pendientes
        
        If loopElemento > 0 Then
            load Me.chkElemento(loopElemento)
            Debug.Print Me.chkElemento(loopElemento).top
            Me.chkElemento(loopElemento).visible = True
            Me.chkElemento(loopElemento).top = Me.chkElemento(loopElemento - 1).top + Me.chkElemento(loopElemento - 1).height + 10
        End If
        
        Me.chkElemento(loopElemento).caption = elemento
        loopElemento = loopElemento + 1
    Next
End Sub
