Attribute VB_Name = "CLI_TileEngine"
Option Explicit


'Info de un objeto
Public Type obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public VerMapa As Boolean

Public pbTecho       As Boolean 'hay techo?

Function EstaPCarea(ByVal Index2 As Integer) As Boolean
  Dim uX As Integer, uY As Integer, CX As Integer, cY As Integer

    uX = UserPos.x
    uY = UserPos.y
    CX = CharList(Index2).Pos.x
    cY = CharList(Index2).Pos.y
    EstaPCarea = (Abs(CX - uX) <= (WindowTileWidth \ 2)) And _
                 (Abs(cY - uY) <= (WindowTileHeight \ 2))

End Function

Function Distancia(wp1 As Position, wp2 As Position) ':( As Variant ?':( Missing Scope

    'Encuentra la distancia entre dos WorldPos
    Distancia = Abs(wp1.x - wp2.x) + Abs(wp1.y - wp2.y)

End Function

