VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDM_Commit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private comentario_ As String
Private archivos As Collection

Public Sub agregarArchivo(Archivo As CDM_Archivo)
    Call archivos.Add(Archivo)
End Sub

Public Function obtenerArchivos() As Collection
    Set obtenerArchivos = archivos
End Function
Private Sub Class_Initialize()
    Set archivos = New Collection
End Sub

Private Sub Class_Terminate()
    Set archivos = Nothing
End Sub

Public Property Let comentario(texto As String)
    comentario_ = texto
End Property

Public Property Get comentario() As String
    comentario = comentario_
End Property

