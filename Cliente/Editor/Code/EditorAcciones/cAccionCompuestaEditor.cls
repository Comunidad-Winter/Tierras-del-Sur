VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccionCompuestaEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iAccionEditor

Option Explicit

Private contenido As Collection
Private nombre As String
Private ejecutarSiempre() As Boolean
Private indentificador As Integer
Private visible As Boolean


Private Sub Class_Initialize()
    Set contenido = New Collection
    visible = True
End Sub

'TODO ¿Descripcion esta muerta?
Public Sub iAccionEditor_crear(nombre_ As String, descripcion_ As String)
    nombre = nombre_
End Sub

Public Sub establecerAccionar(numeroHijo As Byte, accionar As Boolean)
    ejecutarSiempre(numeroHijo - 1) = accionar
End Sub

Public Function seEjecutaSiempre(numeroHijo As Byte) As Boolean
    If numeroHijo = 1 Then
        seEjecutaSiempre = True
    Else
        seEjecutaSiempre = ejecutarSiempre(numeroHijo - 1)
    End If
End Function

Public Sub cambiar(hijo As iAccionEditor, remplazo As iAccionEditor)

    Dim accion As iAccionEditor
    Dim posicion As Integer
    
    posicion = 1
    For Each accion In contenido
    
        If accion Is hijo Then
            contenido.Remove posicion
            
            If contenido.Count > 0 Then
                Call contenido.Add(remplazo, , , posicion - 1)
            Else
                Call contenido.Add(remplazo)
            End If
            
            MsgBox remplazo.GetNombre
            Exit For
        End If
        
        posicion = posicion + 1
        
    Next

    For Each accion In contenido
        MsgBox accion.GetNombre
    Next
End Sub

Public Sub agregarHijo(hijo As iAccionEditor)
    Call contenido.Add(hijo)
    ReDim Preserve ejecutarSiempre(1 To contenido.Count)
End Sub

Public Sub eliminarHijo(hijo As iAccionEditor)

Dim accion As iAccionEditor
Dim posicion As Integer
Dim aux As Integer
posicion = 1

For Each accion In contenido
    If accion Is hijo Then
        contenido.Remove posicion
        Exit For
    End If
    posicion = posicion + 1
Next

For aux = posicion To UBound(ejecutarSiempre) - 1
    ejecutarSiempre(aux) = ejecutarSiempre(aux + 1)
Next

If contenido.Count > 0 Then
    ReDim Preserve ejecutarSiempre(1 To contenido.Count)
End If

End Sub

Public Function obtenerHijos() As Collection
    Set obtenerHijos = contenido
End Function

'/************ SETERS Y GETERS ***********************'
Public Function iAccionEditor_getDescripcion() As String
End Function

Private Function iAccionEditor_getID() As Integer
    iAccionEditor_getID = indentificador
End Function

Public Function iAccionEditor_getNombre() As String
    iAccionEditor_getNombre = nombre
End Function

Private Sub iAccionEditor_setNombre(nombre_ As String)
    nombre = nombre_
End Sub

Public Function iAccionEditor_getNombreExtendido() As String
    iAccionEditor_getNombreExtendido = nombre
End Function

Public Function iAccionEditor_getTIPO() As Long
    iAccionEditor_getTIPO = 1
End Function

Public Function esVisible() As Boolean
    esVisible = visible
End Function
Public Function getID() As Long
    getID = indentificador
End Function

Public Sub setID(id_ As Long)
    indentificador = id_
End Sub

Public Sub setVisible(esVisible As Boolean)
    visible = esVisible
End Sub

'/**************** PERSISTENCIA **********************'
Public Function iAccionEditor_cargar(archivoFuente As Integer) As Variant
    Dim contador As Integer
    Dim tempbyte As Byte
    '1) Tipo de accion
    'Ya viene pre cargada
    
    '2) ID
    Get archivoFuente, , indentificador
    
    '3) Nombre
    Get archivoFuente, , tempbyte
    nombre = Space$(tempbyte)
    Get archivoFuente, , nombre
    
    '3b) Visible
    Get archivoFuente, , visible
    
    '4) Cantidad de hijos
    Dim cantidadHijos As Byte
    Get archivoFuente, , cantidadHijos
    
    '5) Vector de Booleanos
    ReDim ejecutarSiempre(1 To cantidadHijos)
    Get archivoFuente, , ejecutarSiempre
    
    '6) Cargo a los hijos
    Dim tipoHijo As Byte
    
    Dim hijo As iAccionEditor
    For contador = 1 To cantidadHijos
        Get archivoFuente, , tipoHijo
    
        If tipoHijo = 1 Then
            Set hijo = New cAccionCompuestaEditor
            Call hijo.Cargar(archivoFuente)
        Else
            Set hijo = New cAccionTileEditor
            Call hijo.Cargar(archivoFuente)
        End If
        
        Call contenido.Add(hijo)
    Next
    
End Function

Public Function iAccionEditor_persistir(archivoSalida As Integer) As Variant

    '1) Tipo de Accion
    Put archivoSalida, , CByte(1)
    
    '2) id
    Put archivoSalida, , indentificador
    
    '3) Nombre
    Put archivoSalida, , CByte(Len(nombre))
    Put archivoSalida, , nombre
    
    '3b)
    Put archivoSalida, , visible
    
    '4) Cantidad de accion que contiene
    Put archivoSalida, , CByte(UBound(ejecutarSiempre))
    
    '5) Vector de booleanos de patron de forma de ejución
    Put archivoSalida, , ejecutarSiempre
    
    '6) Persistencia de los hijos
    Dim accion As iAccionEditor
    
    For Each accion In contenido
        Call accion.persistir(archivoSalida)
    Next
    
End Function

Private Function iAccionEditor_generarAccionReal() As iAccion
    Dim accion As iAccionEditor
    Dim numeroHijo As Byte
    
    numeroHijo = 1
    
    'Si hay un sólo elemento es al pedo que se persista la accion compuesta
    If contenido.Count = 1 Then
       Set accion = contenido.Item(1)
       Set iAccionEditor_generarAccionReal = accion.generarAccionReal

    Else
        
        Dim accionCompuesta As cAccionCompuesta
        
        Set accionCompuesta = New cAccionCompuesta
        
        For Each accion In contenido
           Call accionCompuesta.agregarHijo(accion.generarAccionReal, ejecutarSiempre(numeroHijo))
            numeroHijo = numeroHijo + 1
        Next
    
        Set iAccionEditor_generarAccionReal = accionCompuesta
    End If
    
End Function
