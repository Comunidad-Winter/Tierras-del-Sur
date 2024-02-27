VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccionTileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iAccionEditor

Private nombre As String
Private descripcion As String
Private listaParametros As New Collection
Private hijo As cAccionTileEditor
Private accion As iAccion

Public Sub crear(nombre_ As String, descripcion_ As String, accion_ As iAccion)
    nombre = nombre_
    descripcion = descripcion_
    Set accion = accion_
End Sub

Public Sub agregarParametro(parametro As cParamAccionTileEditor)
    Call listaParametros.Add(parametro)
End Sub

'/***********************************************************/
'/**************** SETERS Y GETER ***************************/
Public Sub setID(id As Integer)
End Sub

Public Function iAccionEditor_generarAccionReal() As iAccion
    Call accion.cargarParametros(listaParametros)
    Set iAccionEditor_generarAccionReal = accion
End Function

Public Function iAccionEditor_getID() As Integer
    iAccionEditor_getID = 0
End Function

Public Function iAccionEditor_getDescripcion() As String
    iAccionEditor_getDescripcion = descripcion
End Function

Public Function iAccionEditor_getNombre() As String
    iAccionEditor_getNombre = nombre
End Function
Public Function iAccionEditor_getNombreExtendido() As String
'Aparte del nuevo devuelve la información de sus parametros
    iAccionEditor_getNombreExtendido = nombre & " ("
    
    Dim parametro As cParamAccionTileEditor
    
    For Each parametro In listaParametros
        iAccionEditor_getNombreExtendido = iAccionEditor_getNombreExtendido & " " & mid$(parametro.GetNombre, 1, 3) & ":" & parametro.getValor
    Next
    
    iAccionEditor_getNombreExtendido = iAccionEditor_getNombreExtendido & ")"
End Function

Public Function obtenerParametros() As Collection
    Set obtenerParametros = listaParametros
End Function
'**********************************************************/
Public Function Clonar() As cAccionTileEditor
    Dim parametro As cParamAccionTileEditor
    
    Set Clonar = New cAccionTileEditor
    
    Dim listaParametrosClon As New Collection
    
    Call Clonar.crear(nombre, descripcion, getAccion.Clonar)
    
    For Each parametro In listaParametros
        Call Clonar.agregarParametro(parametro.Clonar())
    Next

End Function

Public Function getAccion() As iAccion
     Set getAccion = accion
End Function
'**********************************************************/
'********** FUNCIONES DE PERSISTENCIA *********************/
Public Function iAccionEditor_persistir(archivoSalida As Integer)
        
        '0) tipo
        Put archivoSalida, , CByte(0)
        
        '1) Nombre
        Put archivoSalida, , CByte(Len(nombre))
        Put archivoSalida, , nombre
        
        '2) Descripción
        Put archivoSalida, , CInt(Len(descripcion))
        Put archivoSalida, , descripcion
        
        '3) Cantidad de parametros
        Put archivoSalida, , CByte(listaParametros.Count)
        
        '4) Persistencia de parametros
        Dim parametro As cParamAccionTileEditor
        
        For Each parametro In listaParametros
                Call parametro.persistir(archivoSalida)
        Next parametro
        
        '5) Tipo de la accion que representa
        Put archivoSalida, , CByte(accion.getID)
End Function


Public Function iAccionEditor_cargar(archivoFuente As Integer)

Dim tempbyte As Byte
Dim tempInt As Integer
Dim contadorParametro As Byte
        
        '1) Nombre
        Get archivoFuente, , tempbyte
        nombre = Space$(tempbyte)
        Get archivoFuente, , nombre
        
        '2) Descripcion
        Get archivoFuente, , tempInt
        descripcion = Space$(tempInt)
        Get archivoFuente, , descripcion
        
        '3) Cantidad de parametros
        Get archivoFuente, , tempbyte
        
        '4) Carga de los parametros
        Dim parametro As cParamAccionTileEditor
        
        For contadorParametro = 1 To tempbyte
                Set parametro = New cParamAccionTileEditor
                Call parametro.Cargar(archivoFuente)
                Call listaParametros.Add(parametro)
        Next
        
        '5) ID de la acción que representa
        Get archivoFuente, , tempbyte
        
        Set accion = Sv_Acciones.obtenerAccion(tempbyte)
End Function

Private Function iAccionEditor_getTIPO() As Long
    iAccionEditor_getTIPO = 0
End Function

Private Sub iAccionEditor_setNombre(nombre As String)

End Sub
