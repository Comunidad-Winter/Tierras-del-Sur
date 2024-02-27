VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarEntidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Me puede venir un solo elemento en el array (cuando inserto, borro o corto/copio desde un lugar donde hay solo uno)
'Pueden venir muchos cuando estoy cortando o pegando.

    'Si ID entidad es > 0 y posicion es 0, agrego al final. (Insertar simple)
    'Si ID entidad es > 0 y posicion es > 0, agrego antes de la posicion establecida 1 en la cola (Insertar desde lista)
    'Si ID entidad es 0 y posicion es 0, borro todas las entidades. (Borrar simple)
    'Si ID entidad es 0 y posicion es mayor a 0, borro la entidad en dicha posicion (Borrar desde lista)
    
Option Explicit

Implements iComando

Private nombre As String
Private posx As Integer
Private posy As Integer
Private estoyInsertando As Boolean

Private infoEntidades() As tEntidadSeleccionada
Private infoEntidadesViejas() As tEntidadSeleccionada


Friend Sub crear(x As Integer, y As Integer, infoEntidades_() As Me_Tools_Entidades.tEntidadSeleccionada)

    posx = x
    posy = y
    
    infoEntidades = infoEntidades_

End Sub

'Guarda todas las entidades de un tile
Private Sub backupTile()
    
    
    Dim idEntidad As Integer
    Dim cantidad As Byte
    
    idEntidad = EntidadesMap(posx, posy)
    
    'Cuantos cuantas son
    cantidad = 0
    Do While idEntidad > 0
        cantidad = cantidad + 1
        idEntidad = Entidades(idEntidad).Next
    Loop
    
    'Redimensiono mi backup
    If cantidad = 0 Then
        ReDim infoEntidadesViejas(0)
        Exit Sub
    End If
    
    'Vamos de 0 a N
    ReDim infoEntidadesViejas(0 To cantidad - 1)
    
    'Empiezo por el root
    idEntidad = EntidadesMap(posx, posy)
    
    cantidad = 0
    
    Do While idEntidad > 0

        'Copio los datos de la entidad que hay
        infoEntidadesViejas(cantidad).posicion = cantidad + 1
        infoEntidadesViejas(cantidad).IndexEntidad = Entidades(idEntidad).numeroIndexadoEntidad
        Set infoEntidadesViejas(cantidad).accion = Entidades(idEntidad).accion
        
        'Obtenemos la siguiente entidad
        idEntidad = Entidades(idEntidad).Next
        
        'Avanzamos
        cantidad = cantidad + 1
    Loop


End Sub

'Pone las entidades del backup en el tile
Private Sub RestaurarBackup()

Dim loopEntidad As Integer
Dim tmpId As Integer

'Tengo que restaurar algo? o estaba vacio?
If UBound(infoEntidadesViejas) = 0 And infoEntidadesViejas(0).IndexEntidad = 0 Then
    Call borrarTile
Else
    'Proceso cada entidad
    For loopEntidad = UBound(infoEntidadesViejas) To 0 Step -1
        'Creo cada una de las entidades
        tmpId = Engine_Entidades.Entidades_Crear_Indexada(posx, posy, SV_Simulador.ObtenerIDEntidad, EntidadesIndexadas(infoEntidadesViejas(loopEntidad).IndexEntidad))
        
        'Seteo
        If Not tmpId = -1 Then
            Call Engine_Entidades.Entidades_SetIDIndexada(tmpId, infoEntidadesViejas(loopEntidad).IndexEntidad)
            Call Engine_Entidades.Entidades_SetAccion(tmpId, infoEntidadesViejas(loopEntidad).accion)
        End If
    Next
End If
End Sub

'Elimina todas las entidades de un tile
Private Sub borrarTile()

Dim idEntidad As Integer
Dim siguiente As Integer

'Comienzo por el root
idEntidad = EntidadesMap(posx, posy)

'Hasta que haya una entidad
Do While idEntidad > 0
    'Backupeo la siguiente
    siguiente = Entidades(idEntidad).Next
    'La eliminimo
    Call eliminarEntidad(idEntidad)
    'Avanzo al siguiente
    idEntidad = EntidadesMap(posx, posy)
Loop

End Sub
Private Sub iComando_desHacer()
    borrarTile
    RestaurarBackup
End Sub

Private Sub iComando_hacer()

    Dim simple As Boolean
    Dim posicionDefinida As Boolean
    Dim idnuevaentidad As Integer
    Dim loopEntidad As Byte
    Dim tmpId As Integer
    Dim aux As Integer
    
    estoyInsertando = infoEntidades(1).IndexEntidad > 0 'Si hay
    posicionDefinida = infoEntidades(1).posicion > 0 'Tengo que poner en una posicion o adelante de todo?
    simple = UBound(infoEntidades) = 1
    
    If estoyInsertando Then
        'Guardo el nombre para el caption
        nombre = ""
        For loopEntidad = 1 To UBound(infoEntidades)
            nombre = nombre & EntidadesIndexadas(infoEntidades(loopEntidad).IndexEntidad).nombre & ", "
        Next
        nombre = mid$(nombre, 1, Len(nombre) - 2)
    End If
        
    ' Lo hago
    If estoyInsertando Then
    
        'Si no hay una posicion definida, lo agrego al principio de la lista
        If Not posicionDefinida Then
            'Guardo una copia de lo que habia
            Call backupTile
        
            tmpId = Engine_Entidades.Entidades_Crear_Indexada(posx, posy, SV_Simulador.ObtenerIDEntidad, EntidadesIndexadas(infoEntidades(1).IndexEntidad))
            
            If Not tmpId = -1 Then
                Call Engine_Entidades.Entidades_SetIDIndexada(tmpId, infoEntidades(1).IndexEntidad)
                Call Engine_Entidades.Entidades_SetAccion(tmpId, infoEntidades(1).accion)
            End If
            'En el deshacer, quito la primer entidad de la lista
        Else 'Quiero insertar en una posicion determinada
            
            'Es simple, significa que es solo una entidad que deseo poner en algun lugar
            If simple Then
                'Me fijo en que posicion hay que ponerlo.
                loopEntidad = infoEntidades(1).posicion - 1
            
                tmpId = EntidadesMap(posx, posy)
            
                Do While loopEntidad > 0 And tmpId > 0
                    loopEntidad = loopEntidad - 1
                    tmpId = Entidades(tmpId).Next
                Loop
                
                'Agrego la entidad
                idnuevaentidad = SV_Simulador.ObtenerIDEntidad
                aux = Engine_Entidades.Entidades_Crear_Indexada(0, 0, idnuevaentidad, EntidadesIndexadas(infoEntidades(1).IndexEntidad))
                
                If Not tmpId = -1 Then
                    'Establezco la accion y la referencia a la entidad original
                    Call Engine_Entidades.Entidades_SetIDIndexada(aux, infoEntidades(1).IndexEntidad)
                    Call Engine_Entidades.Entidades_SetAccion(aux, infoEntidades(1).accion)
                
                    'Lo agrego al final
                    If Entidades(tmpId).Next = 0 Then
                        
                        'AA <---> BB
                        Entidades(tmpId).Next = aux 'A con B
                        Entidades(aux).prev = tmpId 'B con A
                    Else 'Lo agrego en el medio
                    
                        ' AA <---> BBB (nuevo) <---> CCC
                        
                        'C tiene como anterior a B
                        Entidades(Entidades(tmpId).Next).prev = aux
                        'B tiene como siguiente a C
                        Entidades(aux).Next = Entidades(tmpId).Next
                        'B tiene como anterior a A
                        Entidades(aux).prev = tmpId
                        'A tiene como siguiente a B
                        Entidades(tmpId).Next = aux
                    End If
                    
                    'Le establezcio la posicion a la entidad
                    Entidades(aux).map_x = posx
                    Entidades(aux).map_y = posy
                End If
            Else
            
                'Tengo varios graficos que insertar
                
                'Backup los actuales.
                'Pongo los nuevos
            End If
        End If
    Else 'Quiero borrar
    
        'Primero hago un backup del tile que voy a borrar
        backupTile
        
        'Tengo que borrar todas o solo una?
        If Not posicionDefinida Then
            'Borro todo
            tmpId = EntidadesMap(posx, posy)
        
            Do While tmpId > 0
                Call eliminarEntidad(tmpId)
                'Obtenemos la siguiente en el tile
                tmpId = EntidadesMap(posx, posy)
            Loop
        Else
            loopEntidad = infoEntidades(1).posicion - 1
            
            tmpId = EntidadesMap(posx, posy)
            
            Do While loopEntidad > 0 And tmpId > 0
                loopEntidad = loopEntidad - 1
                tmpId = Entidades(tmpId).Next
            Loop
            
            'Si esta, la elimino
            If tmpId > 0 Then Call eliminarEntidad(tmpId)
        End If
    End If
End Sub

Private Sub eliminarEntidad(entidad As Integer)
    Call SV_Simulador.EliminarIDEntidad(CInt(Entidades(entidad).id))
    Call Engine_Entidades.eliminar(entidad)
End Sub
Private Function iComando_obtenerNombre() As String
    If estoyInsertando Then
        iComando_obtenerNombre = "Agregar " & nombre & " en (" & posx & "," & posy & ")"
    Else
        iComando_obtenerNombre = "Borrar entidad en (" & posx & "," & posy & ")"
    End If
End Function
