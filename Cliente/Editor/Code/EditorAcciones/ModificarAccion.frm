VERSION 5.00
Begin VB.Form frmModificarAccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nombre del Tile Editando"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   Icon            =   "ModificarAccion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtParametro 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text"
      Top             =   960
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblParametro 
      Caption         =   "descripcion"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label descripcion 
      Caption         =   "Descripción"
      Enabled         =   0   'False
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmModificarAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private accionTileActual As cAccionTileEditor
Private ctlText() As VB.TextBox
Private ctlLabel() As VB.Label

Private Const POSICION_Y = 20
Private Const alto = 20
Private finalizadaEdicion As Boolean

Public Function edicion(padreFormulario As Form) As Boolean
    Me.Show vbModal, padreFormulario
    edicion = finalizadaEdicion
End Function
Public Sub Cargar(accionTile As cAccionTileEditor)
    Dim i As Integer
    Dim valor As Integer
    Dim parametro As cParamAccionTileEditor
    
    
    finalizadaEdicion = False
    
    Set accionTileActual = accionTile
    
    Dim parametros As Collection
    
    Me.caption = "Nuevo " & accionTile.iAccionEditor_getNombre
    Me.descripcion = accionTile.iAccionEditor_getDescripcion
    
    
    Set parametros = accionTileActual.obtenerParametros()
       
    

    
    For i = 1 To parametros.Count
       
            valor = i
            
            Set parametro = parametros.Item(i)

            Call load(Me.txtParametro(valor))
            Call load(Me.lblParametro(valor))
            
            With txtParametro(valor)
                .Height = alto - 5
                .top = POSICION_Y + (alto + 5) * i
                
                .Appearance = 0
                .BorderStyle = 1
                .visible = True
                .Width = 140
                
                .Text = parametro.getValor
                .ToolTipText = parametro.getAyuda
            End With
            
            With lblParametro(valor)
                .Height = alto
                .top = POSICION_Y + (alto + 5) * i
                .Width = 140
                
                .visible = True
                .caption = parametro.GetNombre
                .ToolTipText = parametro.getAyuda
                .MousePointer = 14
            End With

    Next i
    
    Me.cmdAceptar.top = POSICION_Y + (alto + 5) * i + 10
    Me.cmdCancelar.top = POSICION_Y + (alto + 5) * i + 10
    
End Sub

Private Sub cmdAceptar_Click()

    Dim parametros As Collection
    Dim parametro As cParamAccionTileEditor
    Dim modificacionCorrecta As Boolean
    Dim todoCorrecto As Boolean
    Dim contador As Byte
    Dim error As String
    
    error = ""
    todoCorrecto = True
    
    Set parametros = accionTileActual.obtenerParametros()

    contador = 1
    For Each parametro In parametros
                  
            modificacionCorrecta = parametro.setValor(Me.txtParametro(contador))
            
            todoCorrecto = todoCorrecto And modificacionCorrecta
            
            If modificacionCorrecta = False Then
                error = error & parametro.GetNombre() & ": " & parametro.getAyuda & vbCrLf
            End If
            
            contador = contador + 1

    Next

    If Not todoCorrecto Then
        MsgBox "Los siguientes parametros no fueron completados correctamente:" & vbCrLf & error, vbCritical, "¡Error al completar el formulario!"
    Else
        finalizadaEdicion = True
        Unload Me
    End If
    
    
End Sub

Private Sub cmdCancelar_Click()
    finalizadaEdicion = False
    Unload Me
End Sub

