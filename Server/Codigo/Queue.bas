Attribute VB_Name = "Queue"
Option Explicit


Public Type tVertice
    x As Integer
    y As Integer
End Type

Private Const MAXELEM = 1000
Private m_array() As tVertice
Private m_lastelem As Integer
Private m_firstelem As Integer
Private m_size As Integer

'---------------------------------------------------------------------------------------
' Procedure : IsEmpty
' DateTime  : 18/02/2007 20:19
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IsEmpty() As Boolean
IsEmpty = m_size = 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsFull
' DateTime  : 18/02/2007 20:20
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IsFull() As Boolean
IsFull = m_lastelem = MAXELEM
End Function

'---------------------------------------------------------------------------------------
' Procedure : Push
' DateTime  : 18/02/2007 20:20
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Push(ByRef Vertice As tVertice) As Boolean

If Not IsFull Then
    If IsEmpty Then m_firstelem = 1
    m_lastelem = m_lastelem + 1
    m_size = m_size + 1
    m_array(m_lastelem) = Vertice
    Push = True
Else
    Push = False
End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : Pop
' DateTime  : 18/02/2007 20:20
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Pop() As tVertice

If Not IsEmpty Then
    Pop = m_array(m_firstelem)
    m_firstelem = m_firstelem + 1
    m_size = m_size - 1
    If m_firstelem > m_lastelem And m_size = 0 Then
            m_lastelem = 0
            m_firstelem = 0
            m_size = 0
    End If
End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : InitQueue
' DateTime  : 18/02/2007 20:19
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub InitQueue()

ReDim m_array(MAXELEM) As tVertice
m_lastelem = 0
m_firstelem = 0
m_size = 0

End Sub
