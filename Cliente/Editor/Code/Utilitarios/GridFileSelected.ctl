VERSION 5.00
Begin VB.UserControl GridFileSelected 
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2580
   ScaleWidth      =   4920
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      Begin EditorTDS.FileSelector fileSelector 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      LargeChange     =   50
      Left            =   4650
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "GridFileSelected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Eliminamos todos los elementos
Public Sub limpiar()
    Call borrar(fileSelector.count - 1)
End Sub

'Retorna la cantidad de campos completos que hay
Public Property Get cantidad() As Integer
    cantidad = fileSelector.count - 1
End Property
'Propiedades tipicas de un control
Public Property Get Enabled() As Boolean
   Enabled = fileSelector(0).Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Dim Control As Control
    
    For Each Control In Controls
        Control.Enabled = vNewValue
    Next
End Property

'Borramos CANTIDAD de elementos, empezando desde el ultimo
Private Sub borrar(cantidad As Byte)
    Dim i As Byte
    Dim loopDinamico As Byte
    
    For i = 1 To cantidad
        Unload fileSelector(fileSelector.UBound)
    Next

End Sub

Property Get text(Index As Integer) As String
   text = fileSelector(Index).text
End Property
Private Sub redimensionarControles()
Dim i As Integer
Dim cantidad As Byte
Dim totalCampos As Byte

totalCampos = fileSelector.UBound - fileSelector.LBound + 1

'Cuanto la cantidad con ID  = 0
For i = fileSelector.UBound To fileSelector.LBound Step -1
    If Len(fileSelector(i).text) = 0 Then
        cantidad = cantidad + 1
    Else
        Exit For
    End If
Next i

'Si estan todos ocupados, creo uno nuevo.
If cantidad = 0 Then
       
       'Cargo el textbox y el label
        load fileSelector(totalCampos)
       
        With fileSelector(totalCampos)
            .visible = True
            .left = fileSelector(totalCampos - 1).left
            .top = fileSelector(totalCampos - 1).top + fileSelector(0).Height + 50
            .tag = -1
            .text = ""
        End With
    
        
       'Aumento la cantidad de campos que estoy visualizando
       totalCampos = totalCampos + 1
ElseIf cantidad > 1 Then

    'Tengo más de uno libre en el final, tengo que eliminar los sombrantes
    Call borrar(cantidad - 1)
    
    'Actulizo la cantidad de campos que tenia
    totalCampos = totalCampos - (cantidad - 1)
Else
    Exit Sub
End If


' Actualizo el tamaño del frame
'El nuevo largo sera la cantidad de campos que estoy visualizando y para que me entre toda la lista
Frame1.Height = (totalCampos * (fileSelector(0).Height + 50)) + 70

' Actualizo el valor de la barra
Frame1.top = (Frame1.Height - UserControl.Height) * (VScroll1.value / 100) * -1

If Frame1.Height > UserControl.Height Then
    VScroll1.value = -100 * (Frame1.top / (Frame1.Height - UserControl.Height))
Else
    VScroll1.value = 0
End If

End Sub

Private Sub fileSelector_change(Index As Integer, valor As String)
    redimensionarControles
End Sub

Private Sub UserControl_Resize()
    Dim i As Byte

    VScroll1.left = UserControl.Width - VScroll1.Width
    VScroll1.Height = UserControl.Height
    
    Frame1.Width = VScroll1.left - Frame1.left - 1
    
    For i = fileSelector.LBound To fileSelector.UBound
        fileSelector(i).Width = Frame1.Width - fileSelector(i).left - 50
    Next i
    
End Sub

Private Sub actualizarVisibilidad()
    '¿Es mas grande de lo que puedo ver?
    If (Frame1.Height - UserControl.Height) > 0 Then
        'El top va a estar entre 0 y (la altura del frame - lo altura del control, que es la parte que ve el usuario)
        'Depende el porcentaje es donde se ubica en ese intervalo
        Frame1.top = (Frame1.Height - UserControl.Height) * (VScroll1.value / 100) * -1
    Else
        Frame1.top = 0
    End If
End Sub

Private Sub VScroll1_Change()
    Call actualizarVisibilidad
End Sub

Private Sub VScroll1_VScroll1()
    Call actualizarVisibilidad
End Sub

