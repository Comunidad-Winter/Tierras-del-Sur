VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarTileSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private texturaAplicada As Integer
Private texturaNumero As Integer
'Posiciones donde lo aplico
Private posx As Byte
Private posy As Byte

Private viejaTextura As Integer
Private viejoNumero As Integer


Public Sub crear(X As Byte, Y As Byte, textura As Integer, numero As Integer)
    posx = X
    posy = Y
    texturaAplicada = textura
    texturaNumero = numero
End Sub

Private Sub iComando_desHacer()

    'Pongo lo viejo
    mapdata(posx, posy).tile_texture = viejaTextura
    mapdata(posx, posy).tile_number = viejoNumero

   ' Call actualizarEfectoPisada
    
End Sub

Private Sub iComando_hacer()

    'Guardo lo viejo
    viejaTextura = mapdata(posx, posy).tile_texture
    viejoNumero = mapdata(posx, posy).tile_number
            
    'Pongo lo nuevo
    mapdata(posx, posy).tile_texture = texturaAplicada
    mapdata(posx, posy).tile_number = texturaNumero
    
   ' Call actualizarEfectoPisada
End Sub

Private Function iComando_obtenerNombre() As String
    iComando_obtenerNombre = "Insertar " & Tilesets(texturaAplicada).nombre & " en (" & posx & "," & posy & ")"
End Function
