VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The private collection object
Private m_ColvControles As New collection

Private ControlSeleccionado As IControl

Event Click(ByVal x As Integer, ByVal y As Integer, ByRef Cancel As Boolean)

Private pX As Integer
Private pY As Integer

Public Property Get ObjetoFocus() As IControl
    Set ObjetoFocus = ControlSeleccionado
End Property

Public Function RemoveControl(controlBuscado As IControl) As Boolean

Dim Index As Integer

Index = FindIndex(controlBuscado)

If Index = 0 Then
    RemoveControl = False
    Exit Function
End If

If TieneFoco(controlBuscado) Then
    ObjetoFocus = Nothing
End If

Remove Index
RemoveControl = True

End Function


Private Function FindIndex(controlBuscado As IControl)
    Dim Control As IControl
    Dim Pos As Byte
    Pos = 1
    For Each Control In m_ColvControles
        If Control Is controlBuscado Then
            FindIndex = Pos
            Exit Function
        End If
        Pos = Pos + 1
    Next
End Function

Public Property Let ObjetoFocus(objeto As IControl)
    Set ControlSeleccionado = objeto
End Property

Public Function TieneFoco(objeto As IControl) As Boolean
    TieneFoco = objeto Is ControlSeleccionado
End Function

Public Sub GetPos(ByRef x As Integer, ByRef y As Integer)
    x = pX
    y = pY
End Sub

' Agrega al principio de la coleccion
Public Function AddFirst(ICoItem As IControl, Optional VarKey As Variant) As IControl
    m_ColvControles.Add ICoItem, VarKey, 1
    Set AddFirst = ICoItem
End Function

' This sub adds a new IControl item to the collection.
Public Function Add(ICoItem As IControl, Optional VarKey As Variant) As IControl
    m_ColvControles.Add ICoItem, VarKey
    Set Add = ICoItem
End Function

' This sub removes an item from the collection.
Sub Remove(VarIndex As Variant)
    m_ColvControles.Remove VarIndex
End Sub

Function item(VarIndex As Variant) As IControl
Attribute item.VB_UserMemId = 0
    
    Set item = Nothing

    'Marce On error resume next
    Set item = m_ColvControles.item(VarIndex)
    'Marce 'Marce 'Marce On error goto 0

End Function

Function Exists(VarIndex As Variant) As Boolean

    Dim objTest As IControl
    Set objTest = Nothing

    If count > 0 Then

        'Marce On error resume next
        Set objTest = item(VarIndex)
        'Marce 'Marce 'Marce On error goto 0

    End If

    Exists = Not (objTest Is Nothing)

End Function


' This property returns the number of items in the collection.
Property Get count() As Long

    count = m_ColvControles.count

End Property


' This sub removes all items from the collection.
Sub Clear()

    Set m_ColvControles = New collection

End Sub

' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_ColvControles.[_NewEnum]
End Function

Public Function Click(ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim o As IControl
    
    Set ControlSeleccionado = Nothing
    
    For Each o In m_ColvControles
        If o.Click(x, y) Then
            Set ControlSeleccionado = o
            Click = True
            Exit For
        End If
    Next
    
    If Click = False Then
        RaiseEvent Click(x, y, Click)
    End If
End Function

Public Function KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
    If Not (ControlSeleccionado Is Nothing) Then
        KeyDown = ControlSeleccionado.KeyDown(KeyCode, Shift)
    Else
        Dim Control As IControl
        
        For Each Control In m_ColvControles
            KeyDown = Control.KeyDown(KeyCode, Shift)
            If KeyDown Then Exit Function
        Next
    End If
End Function

Public Function KeyPress(ByVal KeyAscii As Integer) As Boolean
    If Not (ControlSeleccionado Is Nothing) Then
        KeyPress = ControlSeleccionado.KeyPress(KeyAscii)
    Else
        Dim Control As IControl
        
        For Each Control In m_ColvControles
            KeyPress = Control.KeyPress(KeyAscii)
            If KeyPress Then Exit Function
        Next
    End If
End Function

Public Function AdvanceFocus() As Boolean
    Dim startPos As Byte
    
    If ControlSeleccionado Is Nothing Then
        startPos = 1
    Else
        startPos = FindIndex(ControlSeleccionado) + 1
    End If
    
    If startPos > m_ColvControles.count Then
        Set ControlSeleccionado = Nothing
        AdvanceFocus = False
        Exit Function
    End If
    
    Set ControlSeleccionado = m_ColvControles(startPos)

    If ControlSeleccionado.IsFocuseable = False Then
         AdvanceFocus = AdvanceFocus()
        Exit Function
    End If
    
    AdvanceFocus = True
End Function

Public Function KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
    If Not (ControlSeleccionado Is Nothing) Then
        KeyUp = ControlSeleccionado.KeyUp(KeyCode, Shift)
    Else
        Dim Control As IControl
        
        For Each Control In m_ColvControles
            KeyUp = Control.KeyUp(KeyCode, Shift)
            If KeyUp Then Exit Function
        Next
    End If
    
End Function

Public Function MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim o As IControl
    
    For Each o In m_ColvControles
        If o.MouseDown(Button, Shift, x, y) Then
            If o.IsMouseFocus Then
                MouseDown = True
                Set ControlSeleccionado = o
            End If
            Exit For
        End If
    Next
End Function

Public Function MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Boolean
    Dim o As IControl
    
    For Each o In m_ColvControles
        If o.MouseMove(Button, Shift, x, y, z) Then
            MouseMove = True
            Exit For
        End If
    Next
End Function

Public Function MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim o As IControl

    For Each o In m_ColvControles
        If o.MouseUp(Button, Shift, x, y) Then
'            If Not (ControlSeleccionado Is o) Then
'                Set ControlSeleccionado = Nothing
'            End If
            MouseUp = True
            Exit For
        End If
    Next
End Function

Public Sub Render()
    Dim Control As IControl
    
    Dim i As Integer
    
    For i = m_ColvControles.count() To 1 Step -1
        Set Control = m_ColvControles(i)
        If Not Control Is Nothing Then Control.Render Control Is ControlSeleccionado
    Next
    
'Dim Control As IControl
    
'    For Each Control In m_ColvControles
'            Control.Render Control Is ControlSeleccionado
'    Next
End Sub
