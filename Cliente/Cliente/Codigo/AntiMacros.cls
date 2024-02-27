VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AntiMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PerformanceValue
    pvSecond = 1                's
    pvDeciSecond = 10           'ds
    pvCentiSecond = 100         'cs
    pvMilliSecond = 1000        'ms
    pvMicroSecond = 1000000     'µs
    pvNanoSecond = 1000000000   'ns
End Enum

Private Const MINIMOCLICK = 10 'en ms
Private Const MINIMOKEY = 15 'en ms
Private Const MAXKEYCODES = 1023

Private m_CountsPerSecond As Currency
Private m_Start As Currency
Private m_Stop As Currency
Private k_Start() As Currency
Private k_Stop() As Currency
Private m_ApiOverhead As Currency

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Sub Class_Initialize()
'on error Resume Next
    'Does the system support a performance counter
    If QueryPerformanceFrequency(m_CountsPerSecond) Then
        Dim I As Long, TotalOverhead As Currency
        
        'Find out how long it takes the system to call the API function
        For I = 1 To 1000
            QueryPerformanceCounter m_Start
            QueryPerformanceCounter m_Stop
            TotalOverhead = TotalOverhead + m_Stop - m_Start
        Next I
        m_ApiOverhead = TotalOverhead / 1000
        Debug.Print m_ApiOverhead
    Else
        m_CountsPerSecond = 1
    End If
    m_Start = 0
    m_Stop = 0
    ReDim k_Start(1 To MAXKEYCODES)
    ReDim k_Stop(1 To MAXKEYCODES)
End Sub

Public Property Get Supported() As Boolean
'on error Resume Next
    'Does the system support a performance counter
    Supported = QueryPerformanceCounter(0)
End Property

'Public Sub StartTimer()
Public Sub ClickRatonDown()
'on error Resume Next
    'Get the start time
    QueryPerformanceCounter m_Start
    m_Stop = 0
End Sub

'Public Sub StopTimer()
Public Function ClickRatonUP() As Boolean
'on error GoTo errh:
    'Get the end time
    QueryPerformanceCounter m_Stop
    If m_Start And m_Stop Then
        ClickRatonUP = (m_Stop - m_Start - m_ApiOverhead) / m_CountsPerSecond * 1000 > MINIMOCLICK
    Else
        ClickRatonUP = False
    End If
Exit Function
errh:
ClickRatonUP = True
End Function
Public Sub ClickKeyDown(Key As Integer)
'on error Resume Next
    'Get the start time
    QueryPerformanceCounter k_Start(Key)
    m_Stop = 0
End Sub

'Public Sub StopTimer()
Public Function ClickKeyUP(Key As Integer) As Boolean
'on error GoTo errh:
    'Get the end time
        QueryPerformanceCounter k_Stop(Key)
    If k_Start(Key) And k_Stop(Key) Then
        ClickKeyUP = (k_Stop(Key) - k_Start(Key) - m_ApiOverhead) / m_CountsPerSecond * 1000 > MINIMOKEY
        k_Start(Key) = 0
        k_Stop(Key) = 0
    Else
        ClickKeyUP = False
    End If
Exit Function
errh:
ClickKeyUP = True
End Function

'Public Property Get TimeElapsed(ByVal ReturnAccuracy As PerformanceValue) As Double
    'Return the time taken
'    If m_Start And m_Stop Then TimeElapsed = (m_Stop - m_Start - m_ApiOverhead) / m_CountsPerSecond * ReturnAccuracy
'End Property
