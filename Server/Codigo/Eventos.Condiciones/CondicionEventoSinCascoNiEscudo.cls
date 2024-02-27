VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoSinCascoEscudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento
               
Friend Sub setParametros(objetosLimites_() As modEvento.tEventoObjetoRestringido, NoPermitirOtrosItems_ As Boolean, NoPermitirOroEnBille_ As Boolean)
End Sub

Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    iCondicionEvento_obtenerDescripcion = "No se puede utilizar ningún tipo de casco ni de escudo."
End Function

Public Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean
    
Dim ObjetoIndex As Integer
Dim loopIntegrante As Integer
Dim loopItem As Integer

iCondicionEvento_puedeIngresarEquipo = False

For loopIntegrante = 1 To UBound(integrantesIndex)
        
    'O(i x l)
    With UserList(integrantesIndex(loopIntegrante)).Invent
    
        For loopItem = 1 To UserList(integrantesIndex(loopIntegrante)).Stats.MaxItems
            
            ObjetoIndex = .Object(loopItem).ObjIndex
            
            If ObjetoIndex > 0 Then
                
                ' ¿Es una armadura?
                If ObjData(ObjetoIndex).ObjType = OBJTYPE_ARMOUR Then
                    
                    ' Es un Casco o un Escudo y además no es sombrero de aprendiz o de mago
                    If (ObjData(ObjetoIndex).subTipo = OBJTYPE_CASCO _
                            Or ObjData(ObjetoIndex).subTipo = OBJTYPE_ESCUDO) _
                        And Not (ObjetoIndex = SOMBRERO_DE_APRENDIZ Or ObjetoIndex = SOMBRERO_DE_MAGO) Then
                    
                        If ObjData(ObjetoIndex).MaxDef > 0 Then
                            Exit Function
                        End If
                        
                    End If

                End If

            End If
       Next loopItem
    End With
    
Next loopIntegrante

iCondicionEvento_puedeIngresarEquipo = True

End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantesid() As Long)
End Sub

Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
End Sub

Private Function contarCantidadItems(inventario As inventario, IndexItem As Integer) As Long
    
    contarCantidadItems = 0
    
    Dim loopItemInventario As Byte
    
    For loopItemInventario = 1 To UBound(inventario.Object)
            If inventario.Object(loopItemInventario).ObjIndex = IndexItem Then
                contarCantidadItems = contarCantidadItems + inventario.Object(loopItemInventario).Amount
            End If
    Next loopItemInventario
    
End Function


