VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private npcAgregar As tNPCSeleccionado
Private npcAnterior As tNPCSeleccionado

Private posx As Byte
Private posy As Byte

Public Sub crear(NpcIndex As Integer, NpcZona As Byte, x As Byte, y As Byte)
    npcAgregar.Index = NpcIndex
    npcAgregar.zona = NpcZona
    
    posx = x
    posy = y
End Sub
Private Sub iComando_desHacer()
    'Si habia puesto algo, lo saco
    If npcAgregar.Index > 0 Then Call eliminarNpc
    
    'Si habia borrado uno... lo pongo de nuevo
    If npcAnterior.Index > 0 Then Call agregarNpc(npcAnterior)
End Sub

Private Sub iComando_hacer()
    npcAnterior.Index = MapData(posx, posy).NpcIndex
    npcAnterior.zona = MapData(posx, posy).NpcZona
    
    'Hay uno viejo?. Ya lo tengo guardado en npcAnterior. Lo elimino
    If npcAnterior.Index > 0 Then Call eliminarNpc
    
    '¿Tengo que agregar algo? (Sino solo es borrar)
    If npcAgregar.Index > 0 Then Call agregarNpc(npcAgregar)
End Sub

Private Sub agregarNpc(Npc As tNPCSeleccionado)
    Dim Charindex As Integer
    
    'Le solicito al simulador del server el charindex
    Charindex = SV_Simulador.NextOpenChar(True)
        
    'Creo el NPC
    Call MakeChar(Charindex, NpcData(Npc.Index).body, _
                                    NpcData(Npc.Index).Head, _
                                    NpcData(Npc.Index).heading, _
                                    posx, posy, 0, 0, 0)

    'Si utilizo ActivateChar y no estoy en Modo Caminata, no se van a visualizar
    'los npcs, ya que activate char solo actuva con los npcs que estoy viendo y para esto
    'utiliza la posicion del Userli(CharIndex)
    
    CharList(Charindex).active = 1
    'Actualiamos la neuva posicion ya sea mia o del otro charindex
    CharMap(posx, posy) = Charindex
    'Call ActivateChar(CharList(Char))
    
    'Lo agrego al mapa
    MapData(posx, posy).NpcIndex = Npc.Index
    MapData(posx, posy).NpcZona = Npc.zona
End Sub
Private Sub eliminarNpc()
    Dim charIndexAnterior As Integer
    
    'Obtengo el CharIndex y el NpcIndex (el tipo de NPC que es)
    charIndexAnterior = CharMap(posx, posy)
    
    'Lo elimino del mapa
    MapData(posx, posy).NpcIndex = 0
    MapData(posx, posy).NpcZona = 0
    
    'Borro el caracter
    Call EraseChar(charIndexAnterior)
        
    'Retorno el CharIndex al simulador de servidor
    Call SV_Simulador.EraseIndexChar(charIndexAnterior)
End Sub
Private Function iComando_obtenerNombre() As String
    If npcAgregar.Index > 0 Then
        iComando_obtenerNombre = "Insertar " & NpcData(npcAgregar.Index).Name & " en (" & posx & "," & posy & ")"
    Else
        If npcAnterior.Index > 0 Then
            iComando_obtenerNombre = "Borrar " & NpcData(npcAnterior.Index).Name & " en (" & posx & "," & posy & ")"
        End If
    End If
End Function
