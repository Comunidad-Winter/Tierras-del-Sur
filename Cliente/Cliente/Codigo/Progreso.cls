VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FinTick As Long
Private StartTick As Long

Private inicio_ As Integer
Private incremento_ As Integer
Private isCompleto_ As Boolean

Public Sub SetRango(inicio As Integer, fin As Integer)
    inicio_ = inicio
    incremento_ = fin - inicio
End Sub

Public Sub SetTicks(endTick As Long)
    FinTick = endTick
    StartTick = GetTimer
End Sub

Public Function Calcular() As Single
    Dim TmpTick As Long
    Dim porcentaje As Single
    
    TmpTick = GetTimer
    
    If TmpTick < FinTick Then
        porcentaje = CosInterp(0, 1, (TmpTick - StartTick) / (FinTick - StartTick))
    Else
        porcentaje = 1
        isCompleto_ = True
    End If
    
    Calcular = inicio_ + incremento_ * porcentaje
End Function

Public Property Get isCompleto() As Boolean
    isCompleto = isCompleto_
End Property


