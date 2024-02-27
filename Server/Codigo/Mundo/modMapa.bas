Attribute VB_Name = "modMapa"
Option Explicit


Type ObjInfoPoseedor
    UserIndex As Integer                ' Identificador dueño del usuario
    fecha As Long                       ' Momento en el cual se hizo dueño de este objeto
End Type

' Información de cada tile
Type MapBlock
    UserIndex As Integer                ' Usuario que esta parado en ese momento
    npcIndex As Integer                 ' Criatura que esta parada ene se momento
    ObjInfoPoseedor As ObjInfoPoseedor  ' Quien es el dueño del ObjInfo que esta en este tile
    OBJInfo As obj                      ' Objeto que esta en este tile
  

    Trigger As Long                   ' Propiedades del tile
    accion As iAccion
End Type


'Info del mapa
Type MapInfo
    Existe As Boolean '¿Esta cargado el mapa?

    usuarios As EstructurasLib.ColaConBloques       ' Index de Usuarios en el mapa
    NPCs As EstructurasLib.ColaConBloques           ' Index de Criaturas en el mapa
    fogatas As Collection                           ' Fogatas que hay en el mapa. Para evitar que se llene
    
    Music As String                                 ' Musica del mapa
    Name As String                                  ' Nombre del mapa
    MapVersion As Integer                           ' Version del mapa.
    Pk As Boolean                                   ' PK=0 Zona segura, Pk=1 Insegura
    Terreno As String                               ' Tipo de Terreno
    zona As String                                  ' Tipo de Zona
    Frio As Byte                                    ' ¿Hace frio en este mapa?
    Calor As Byte                                   ' ¿Hace extremo calor?
    restringir As Byte                              '  1 Es un mapa para newbies
    clima As Integer                                ' Climas admitidos
    climaActual As Integer                          ' Clima Actual
    
    BackUp As Byte                                  ' Se guarda el estado del mapa
    
    ' Restricciones
    Nivel As Integer                                ' Nivel minimo para ingresar al mapa
    MaxLevel As Integer                             ' Nivel maximo para ingresar al mapa
    
    SeCaeiItems As Byte                             ' Los items se cae o no en esta mapa?
    Aotromapa As WorldPos                           ' Si muere el personaje es enviado a otro mapa?
    AntiHechizosPts As Byte                         ' No se permiten algunos hechizos como elementales
    PermiteRoboNPC As Byte                          ' Se pueden robar npcs
    
    UsuariosMaximo As Integer                       ' Cantidad maxima de usuarios que puede estar en el mapa
    
    SoloCiudas As Byte                              ' Solo pueden entrar ciudadanos
    SoloCrimis As Byte                              ' Solo pueden entrar criminales
    SoloCaos As Byte                                ' Solo pueden entrar Caos
    SoloArmada As Byte                              ' Solo pueden entrar armadas
    
    Continente As Integer                           ' Numero de continente.
End Type

Public MapData() As MapBlock
Public MapInfo() As MapInfo


Public Function EsPosicionParaAtacarSinPenalidad(pos As WorldPos) As Boolean
    EsPosicionParaAtacarSinPenalidad = (MapData(pos.map, pos.x, pos.y).Trigger And eTriggers.CombateSeguro)
End Function

Public Function TirarItemAlPisoConDuenio(pos As WorldPos, ByRef objeto As obj, ByVal duenio As Integer) As WorldPos

    ' Posicion donde lo voy a tirar
    Dim NuevaPos As WorldPos
    Dim objetoACrear As obj
    
    NuevaPos.x = 0
    NuevaPos.y = 0
    NuevaPos.map = 0
    
    ' Obtengo la Posicion
    Call TileLibreParaObjeto(pos, NuevaPos, objeto)
    
    ' Si encontre, lo tiro
    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
    
        ' Objeto Auxiliar
        objetoACrear.Amount = objeto.Amount
        objetoACrear.ObjIndex = objeto.ObjIndex
            
        ' Creamos el objeto
        If MapData(NuevaPos.map, NuevaPos.x, NuevaPos.y).OBJInfo.ObjIndex = objeto.ObjIndex Then
            If MapData(NuevaPos.map, NuevaPos.x, NuevaPos.y).OBJInfo.Amount + objeto.Amount < MAX_OBJETOS_X_SLOT Then
                'Acumulamos
                objetoACrear.Amount = objetoACrear.Amount + MapData(NuevaPos.map, NuevaPos.x, NuevaPos.y).OBJInfo.Amount
            End If
        End If
            
        ' Creamos remplazando el existente
        Call MakeObjDuenio(ToMap, 0, NuevaPos.map, objetoACrear, NuevaPos.map, NuevaPos.x, NuevaPos.y, duenio)
        
        ' Devolvemos la posicion
        TirarItemAlPisoConDuenio = NuevaPos
    End If
    

End Function

Public Function TirarItemAlPiso(pos As WorldPos, ByRef objeto As obj) As WorldPos
   Call TirarItemAlPisoConDuenio(pos, objeto, 0)
End Function

Public Sub TileLibreParaObjeto(ByRef pos As WorldPos, ByRef nPos As WorldPos, ByRef objeto As obj)

Dim encontrado As Boolean ' ¿Encontre?
Dim loopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean

loopC = 0

nPos.map = pos.map
nPos.x = 0
nPos.y = 0

Do
    
    ' Recorro armando un cuadrado
    ' TODO Aca se analizan varias veces los mismos tiles
    For tY = pos.y - loopC To pos.y + loopC
        For tX = pos.x - loopC To pos.x + loopC
        
            ' ¿Posicion legal?
            If esPosicionJugable(tX, tY) = True Then
                
                If isTileBloqueado(MapData(pos.map, tX, tY)) = False Then
                    ' No puede caer sobre portales
                    If MapData(nPos.map, tX, tY).accion Is Nothing Then
                    
                        If MapData(nPos.map, tX, tY).OBJInfo.ObjIndex = 0 Then
                            ' Tile libre
                            nPos.x = tX
                            nPos.y = tY
                            Exit Sub
                        ElseIf MapData(nPos.map, tX, tY).OBJInfo.ObjIndex = objeto.ObjIndex Then
                            ' En el tile hay un objeto similar ¿Hay espacio para acumularlo?
                            If MapData(nPos.map, tX, tY).OBJInfo.Amount + objeto.Amount <= MAX_OBJETOS_X_SLOT Then
                                    nPos.x = tX
                                    nPos.y = tY
                                    Exit Sub
                            End If ' No hay espacio
                        End If
                        
                    End If
                End If
            End If
        Next tX
    Next tY
    
    ' Proximo intento ampliamso el rango
    loopC = loopC + 1
    
' Lo hago hasta que no lo encuentre o haya superado la cantidad de intentos
Loop Until loopC >= 15

End Sub


Private Function transformarObjeto(ObjIndex)

If ObjIndex = 6 Then transformarObjeto = 833: Exit Function
If ObjIndex = 46 Then transformarObjeto = 834: Exit Function
If ObjIndex = 47 Then transformarObjeto = 835: Exit Function
If ObjIndex = 49 Then transformarObjeto = 836: Exit Function
If ObjIndex = 50 Then transformarObjeto = 837: Exit Function
If ObjIndex = 147 Then transformarObjeto = 838: Exit Function
If ObjIndex = 148 Then transformarObjeto = 839: Exit Function
If ObjIndex = 149 Then transformarObjeto = 839: Exit Function
If ObjIndex = 150 Then transformarObjeto = 840: Exit Function
If ObjIndex = 151 Then transformarObjeto = 841: Exit Function
If ObjIndex = 152 Then transformarObjeto = 842: Exit Function
If ObjIndex = 153 Then transformarObjeto = 843: Exit Function
If ObjIndex = 154 Then transformarObjeto = 844: Exit Function
If ObjIndex = 635 Then transformarObjeto = 845: Exit Function
If ObjIndex = 636 Then transformarObjeto = 846: Exit Function
If ObjIndex = 637 Then transformarObjeto = 847: Exit Function
If ObjIndex = 713 Then transformarObjeto = 848: Exit Function
If ObjIndex = 714 Then transformarObjeto = 849: Exit Function
If ObjIndex = 715 Then transformarObjeto = 850: Exit Function
If ObjIndex = 716 Then transformarObjeto = 851: Exit Function
If ObjIndex = 717 Then transformarObjeto = 852: Exit Function
If ObjIndex = 718 Then transformarObjeto = 853: Exit Function
If ObjIndex = 719 Then transformarObjeto = 854: Exit Function
If ObjIndex = 720 Then transformarObjeto = 855: Exit Function
If ObjIndex = 721 Then transformarObjeto = 856: Exit Function
If ObjIndex = 722 Then transformarObjeto = 857: Exit Function

transformarObjeto = ObjIndex
End Function


Public Function listarPersonajesOnline(mapaInfo As MapInfo) As String
    
    If mapaInfo.usuarios.getCantidadElementos > 0 Then

        With mapaInfo.usuarios
        
            .itIniciar
            
            Do While .ithasNext
                listarPersonajesOnline = listarPersonajesOnline & UserList(.itnext).Name & ", "
            Loop
            
        End With
        
        listarPersonajesOnline = Left$(listarPersonajesOnline, Len(listarPersonajesOnline) - 2)
    Else
        listarPersonajesOnline = "Mapa vacio."
    End If
    

End Function

Public Function HayAgua(map As Integer, x As Integer, y As Integer) As Boolean
    HayAgua = (MapData(map, x, y).Trigger And eTriggers.Navegable)
End Function

