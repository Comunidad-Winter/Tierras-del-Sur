VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemMapaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Información de un item en un mapa
Option Explicit

Public index As Integer         ' Index del Objeto
Public x As Integer             ' Posicion en la cual esta en el mapa
Public y As Integer
Public fecha As Long            ' Fecha en el cual fue puesto en ese mapa

