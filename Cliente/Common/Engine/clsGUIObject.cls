VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUIObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click(x!, Y!)
Event MouseMove(x!, Y!)

Public Property Get Color() As Long

End Property

Public Property Let Color(ByVal n As Long)

End Property

Public Sub SetPos(ByVal x%, ByVal Y%)

End Sub

Public Sub GetPos(ByRef x%, ByRef Y%)

End Sub

Public Sub Render()

End Sub

Public Property Let Text(ByVal txt As String)

End Property

Public Property Get Text() As String

End Property

Public Property Get Color2() As Long

End Property

Public Property Let Color2(ByVal nColor2 As Long)

End Property

Public Property Get GradientMode() As eGuiColorFill

End Property

Public Property Let GradientMode(ByVal vNewValue As eGuiColorFill)

End Property
