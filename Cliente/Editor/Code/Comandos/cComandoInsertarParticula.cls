VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarParticula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim posX As Byte
Dim posY As Byte
Dim Capa As Byte

Private BackUpGroup As Engine_Particle_Group

Private PGID_Nuevo As Integer

Implements iComando

Public Sub crear(ByVal pCapa As Byte, ByVal X As Byte, ByVal Y As Byte, ByVal PGID As Integer)
    posX = X
    posY = Y
    
    Capa = pCapa Mod 3
    
    Set BackUpGroup = MapData(posX, posY).Particles_groups(Capa)
    
    PGID_Nuevo = PGID
    
End Sub

Private Sub iComando_desHacer()
    Set MapData(posX, posY).Particles_groups(Capa) = BackUpGroup
End Sub

Private Sub iComando_hacer()
    If PGID_Nuevo <> -1 Then
        Dim t As Engine_Particle_Group
        
        Set t = New Engine_Particle_Group
        t = PGID_Nuevo
        t.SetPos posX, posY
        
        Set MapData(posX, posY).Particles_groups(Capa) = t
    Else
        Set MapData(posX, posY).Particles_groups(Capa) = Nothing
    End If
End Sub

Private Function iComando_obtenerNombre() As String
If PGID_Nuevo Then
    iComando_obtenerNombre = "Insertar particula en la capa " & Capa & " en (" & posX & "," & posY & ")"
Else
    iComando_obtenerNombre = "Borrar particula en la capa " & Capa & " en (" & posX & "," & posY & ")"
End If
End Function

