VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine_IDesencadenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Property Get Anterior() As Engine_IDesencadenador

End Property

Private Property Set Anterior(ByRef RHS As Engine_IDesencadenador)

End Property

Private Property Set Siguiente(ByRef RHS As Engine_IDesencadenador)

End Property

Private Property Get Siguiente() As Engine_IDesencadenador

End Property

Private Property Get Tick() As Long

End Property

Private Property Let Tick(ByVal RHS As Long)

End Property

Public Function Serialize() As String
' Serialize: Devuelve una cadena comprimida que pueda ser enviada y/o guardada
End Function

Public Function Unserialize(SerializedString As String) As Boolean
' Unserializa: Es una especie de constructor, asigna los valores de la clase en base a una cadena
End Function

Public Sub Ejecutar()
' Ejecuta la accion almacenada en los campos.
End Sub

