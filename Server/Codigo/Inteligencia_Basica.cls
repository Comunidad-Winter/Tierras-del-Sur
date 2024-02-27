VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inteligencia_Basica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Inteligencia

Private Sub Inteligencia_determinarAtaque(npcIndex As Integer, indexAtaque As Integer, ByRef tipoobjetivo As eTipoObjetivo, ByRef tipoAtaque As eTipoAtaque, ByRef subTipo As Byte)
    Call determinarAtaque(NpcList(npcIndex), indexAtaque, tipoobjetivo, tipoAtaque, subTipo)
End Sub

Private Sub Inteligencia_determinarObjetivo(npcIndex As Integer, ByRef index As Integer, ByRef tipoobjetivo As eTipoObjetivo)
    Call determinarObjetivo(NpcList(npcIndex), index, tipoobjetivo)
End Sub

Private Function Inteligencia_determinarMovimiento(npcIndex As Integer, tipoobjetivo As eTipoObjetivo, IndexObjetivo As Integer) As eHeading
    Inteligencia_determinarMovimiento = determinarMovimiento(NpcList(npcIndex), tipoobjetivo, IndexObjetivo)
End Function


'Funciones privadas
Private Sub determinarObjetivo(criatura As npc, ByRef index As Integer, ByRef tipoobjetivo As eTipoObjetivo)
    Call IA_BASICA.determinarObjetivo_A1(criatura, index, tipoobjetivo)
End Sub

Private Function determinarMovimiento(criatura As npc, tipoobjetivo As eTipoObjetivo, IndexObjetivo As Integer) As eHeading
    determinarMovimiento = IA_BASICA.determinarMovimiento_A1(criatura, tipoobjetivo, IndexObjetivo)
End Function


Private Sub determinarAtaque(criatura As npc, indexAtaque As Integer, ByRef tipoobjetivo As eTipoObjetivo, ByRef tipoAtaque As eTipoAtaque, ByRef subTipo As Byte)
   Call IA_BASICA.determinarAtaque_A1(criatura, indexAtaque, tipoobjetivo, tipoAtaque, subTipo)
End Sub

