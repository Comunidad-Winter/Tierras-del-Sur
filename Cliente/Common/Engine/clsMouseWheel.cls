VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long


Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetParent Lib "user32" ( _
                ByVal hWnd As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157

Private phwnd As Long

Public Sub WheelHook(ByVal hWnd As Long)
  'Marce On error resume next
  If phwnd Then WheelUnHook phwnd
  SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf mouseWindowProc)
  phwnd = hWnd
End Sub

Public Sub WheelUnHook(ByVal hWnd As Long)
  'Marce On error resume next
  SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")
  RemoveProp hWnd, "PrevWndProc"
  phwnd = 0
End Sub

Private Sub Class_Terminate()
    If phwnd Then WheelUnHook phwnd
End Sub
