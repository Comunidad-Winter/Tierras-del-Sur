VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "desencadenadorCrearParticula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Engine_IDesencadenador

Private NumeroEmisor As Integer
Private x As Byte
Private y As Byte

Private Anterior    As Engine_IDesencadenador
Private Siguiente   As Engine_IDesencadenador
Private Tick        As Long

Private DesencadenadorMuerte            As String

Private Property Get Engine_IDesencadenador_Anterior() As Engine_IDesencadenador
    Set Engine_IDesencadenador_Anterior = Anterior
End Property

Private Property Set Engine_IDesencadenador_Anterior(ByRef RHS As Engine_IDesencadenador)
    Set Engine_IDesencadenador_Anterior = RHS
End Property

Private Property Set Engine_IDesencadenador_Siguiente(ByRef RHS As Engine_IDesencadenador)
    Set Siguiente = RHS
End Property

Private Property Get Engine_IDesencadenador_Siguiente() As Engine_IDesencadenador
    Set Engine_IDesencadenador_Siguiente = Siguiente
End Property

Private Property Get Engine_IDesencadenador_Tick() As Long
    Engine_IDesencadenador_Tick = Tick
End Property

Private Property Let Engine_IDesencadenador_Tick(ByVal RHS As Long)
    Tick = RHS
End Property

Private Sub Engine_IDesencadenador_Ejecutar()
    Engine_Particles.Particle_Group_Make 0, x, y, NumeroEmisor
End Sub

Private Function Engine_IDesencadenador_Serialize() As String
    Dim buffer As String * 5

        msSetPuntero StrPtr(buffer)
        msWriteByte eDesencadenadores.CrearParticula
        msWriteByte x
        msWriteByte y
        msWriteInteger NumeroEmisor
    
    Engine_IDesencadenador_Serialize = buffer
End Function

Private Function Engine_IDesencadenador_Unserialize(SerializedString As String) As Boolean
        msSetPuntero StrPtr(SerializedString)
        msReadByte
        x = msReadByte
        y = msReadByte
        NumeroEmisor = msReadInteger
    Engine_IDesencadenador_Unserialize = True
End Function
