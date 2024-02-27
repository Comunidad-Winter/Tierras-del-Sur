VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Interface de las vWindows



Public Function Render() As Boolean

End Function

Public Sub SetPos(ByVal X As Integer, ByVal Y As Integer)

End Sub

Public Property Get X() As Integer
End Property

Public Property Get Y() As Integer
End Property

Public Property Get width() As Integer
End Property

Public Property Get Height() As Integer
End Property

Public Property Get IsVisible() As Boolean
End Property

Public Sub Hide()

End Sub

Public Sub Show()

End Sub

Public Property Get GetControl() As vControles

End Property

Public Sub SetPrev(vWindow As vWindow)

End Sub

Public Function GetPrev(vWindow As vWindow) As Boolean

End Function

Public Sub SetNext(vWindow As vWindow)

End Sub

Public Function GetNext(vWindow As vWindow) As Boolean

End Function
