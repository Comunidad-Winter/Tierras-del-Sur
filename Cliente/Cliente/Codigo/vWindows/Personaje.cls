VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Personaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Nick As String

Public id As Long

Public Password As String

Public Clase As String

Public cuerpo As Integer

Public cuerpoDesnudo As Integer

Public cabeza As Integer

Public mapa As Integer

Public modoCandado As Boolean

Public online As Boolean

Public bloqueado As Boolean

Public pelo As Integer

Public barba As Integer

Public ropaInterior As Integer
