VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Click(ByVal x As Integer, ByVal y As Integer) As Boolean

End Function

Public Function KeyDown(KeyCode As Integer, Shift As Integer) As Boolean

End Function

Public Function KeyPress(KeyAscii As Integer) As Boolean

End Function

Public Function KeyUp(KeyCode As Integer, Shift As Integer) As Boolean

End Function

Public Function MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Boolean

End Function

Public Function MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

End Function

Public Function MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

End Function

Public Sub SetPos(ByVal x As Integer, ByVal y As Integer)

End Sub

Public Sub SetSize(ByVal ancho As Integer, ByVal alto As Integer)
End Sub

Public Sub Render(ByVal focus As Boolean)
End Sub

' Si se le hace clic con el mouse, el componente mantiene el foco?
Public Property Get IsMouseFocus() As Boolean
End Property

' Este componente puede mantener el foco?. Por ejemplo, un input si, pero un label no.
Public Property Get IsFocuseable() As Boolean
End Property
