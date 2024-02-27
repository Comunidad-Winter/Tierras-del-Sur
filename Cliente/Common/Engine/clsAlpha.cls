VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Velocidad As Long

Private FinalAnimacion As Long
Private InicioAnimacion As Long

Private InicioValor As Single
Private FinalValor As Single

Private ValorActual As Single

Private Animando As Byte

Public Property Get value() As Single
Attribute value.VB_UserMemId = 0
Attribute value.VB_MemberFlags = "200"
    If Animando Then
        Dim Delta As Single
        Delta = (GetTimer - InicioAnimacion) / (FinalAnimacion + 1)
        
        If Delta < 1 Then
            ValorActual = InicioValor + Delta * (FinalValor - InicioValor)
        Else
            ValorActual = FinalValor
            Animando = False
        End If
    End If
    
    value = ValorActual
End Property

Public Property Let value(ByVal fvalue As Single)
    'If ValorActual <> fvalue And (FinalValor <> fvalue Or Animando = False) Then
    If FinalValor <> fvalue Or (Animando = False And ValorActual <> fvalue) Then
        InicioValor = ValorActual
        FinalValor = fvalue
        InicioAnimacion = GetTimer
        FinalAnimacion = (Velocidad * Abs(FinalValor - InicioValor)) / 256
        Animando = True
    End If
End Property

Public Property Let InitialValue(ByVal fvalue As Single)
    Animando = False
    InicioValor = fvalue
    ValorActual = fvalue
    FinalValor = fvalue
End Property

Public Property Get Speed() As Long
    Speed = Velocidad
End Property

Public Property Let Speed(ByVal fspeed As Long) 'Esta velocidad es los milisegundos que tarda en hacer 256 posiciones el valor.
    Velocidad = fspeed
    Animando = False
    value = FinalValor
End Property


Private Sub Class_Initialize()
    Velocidad = 1000
    InitialValue = 0
End Sub
