VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILucesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'INTERFACE
Public Sub Iterador_Iniciar(): End Sub
Public Function Iterar() As Integer: End Function
Public Sub Actualizar(): End Sub
Public Function count() As Integer: End Function
Public Function Find(ByVal ID As Long) As Long: End Function
Public Function Remove_All() As Boolean: End Function
Public Function Quitar(ByRef light_index As Integer) As Boolean: End Function
Public Function Get_Color(ByVal light_index As Long, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte) As Boolean: End Function
Public Function Get_Light(ByVal light_index As Long, map_x As Byte, map_y As Byte, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte, radio As Byte, brillo As Byte, ID As Long, tipo As Integer) As Boolean: End Function
Public Function Crear(ByVal map_x As Integer, ByVal map_y As Integer, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, Optional ByVal range As Byte = 1, Optional ByVal brillo As Byte = 255, Optional ByVal ID As Long, Optional ByVal tipo As TipoLuces = Luz_Normal) As Long: End Function
Public Function Move(ByVal light_index As Long, Optional ByVal map_x As Integer, Optional ByVal map_y As Integer, Optional ByVal PixelOffsetX As Integer, Optional ByVal PixelOffsetY As Integer) As Boolean: End Function
Public Function MovePixel(ByVal light_index As Long, ByVal PixelPosX As Integer, ByVal PixelPosY As Integer) As Boolean: End Function
Public Property Get NeedUpdateLights() As Boolean: End Property
Public Property Get NeedUpdateMap() As Boolean: End Property
Public Sub MapUpdated(): End Sub
Public Sub LightsUpdated(): End Sub
Public Sub Iniciar(Cantidad As Long, DestArray As Long): End Sub
