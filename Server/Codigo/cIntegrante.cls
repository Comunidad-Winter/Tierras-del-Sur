VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIntegrante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private nombre As String
Private IDPJ As Long


Public Sub iniciar(NombreUser As String, IDUser As Long)
    nombre = UCase$(NombreUser)
    IDPJ = IDUser
End Sub

Public Function getNombre() As String
    getNombre = nombre
End Function

Public Function getIDPJ() As Long
    getIDPJ = IDPJ
End Function
