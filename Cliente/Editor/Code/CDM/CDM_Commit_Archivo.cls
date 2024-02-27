VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDM_Archivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private archivo_ As String ' Ruta absoluta al archivo
Private destino_ As String ' A que corresponde este archivo GRAFICOS, RECURSO_IMAGEN, o directorio ./ ../
Private cambios_ As Object ' Modificaciones que contiene el archivo
Private version_ As Long 'Si corresponde, version del archivo

Public Sub iniciar(archivo As String, destino As String, cambios As Object, Optional ByVal Version As Long = 0)
    archivo_ = archivo
    destino_ = destino
    version_ = Version
    Set cambios_ = cambios
End Sub

Public Property Get archivo() As String
    archivo = archivo_
End Property

Public Property Let archivo(ruta As String)
    archivo_ = ruta
End Property

Public Property Get destino() As String
    destino = destino_
End Property

Public Property Get cambios() As Object
    Set cambios = cambios_
End Property

Public Property Get Version() As Integer
    Version = version_
End Property
