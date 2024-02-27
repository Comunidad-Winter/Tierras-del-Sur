VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChequeoMemoriaPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public personajeID As Long              ' Personaje sobre el cual se hace el chequeo
Public rtaEsperada As String            ' Numero del chequeo que se esta haciendo sobre el personaje
Public fecha As Long                    ' Fecha y hora en la cual comenzo el chequeo
Public nombreChequeo As String          ' Nombre del chequeo
