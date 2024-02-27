Attribute VB_Name = "modObjeto"
Option Explicit

Public Function EsMineral(ObjIndex As Integer) As Boolean
    If ObjIndex = OBJTYPE_MINERALES Then
        EsMineral = True
        Exit Function
    End If
    EsMineral = False
End Function

Public Function isFaccionario(ByRef objeto As ObjData) As Boolean

    isFaccionario = False
    
    If objeto.alineacion = eAlineaciones.Neutro Or objeto.alineacion = eAlineaciones.indefinido Then
        Exit Function
    End If
    
    isFaccionario = True
    
End Function

Public Function ObjEsRobable(ByRef objeto As ObjData) As Boolean
    ' Agregué los barcos
    ' Esta funcion determina qué objetos son robables.
    
    If objeto.ObjType = OBJTYPE_LLAVES Then
        ObjEsRobable = False
        Exit Function
    End If
    
    If objeto.ObjType = objeto.ObjType <> OBJTYPE_BARCOS Then
        ObjEsRobable = False
        Exit Function
    End If
    
    If isFaccionario(objeto) Then
        ObjEsRobable = False
        Exit Function
    End If

    ObjEsRobable = True

End Function

Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean
    ItemNoEsDeMapa = ObjData(index).ObjType <> OBJTYPE_PUERTAS And _
                ObjData(index).ObjType <> OBJTYPE_CARTELES And _
                ObjData(index).ObjType <> OBJTYPE_ARBOLES And _
                ObjData(index).ObjType <> OBJTYPE_YACIMIENTO And _
                ObjData(index).ObjType <> OBJTYPE_TELEPORT
End Function

