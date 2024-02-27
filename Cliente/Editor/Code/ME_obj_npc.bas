Attribute VB_Name = "ME_obj_npc"
'*****************************************************************************
'        MODULO AUXILIAR QUE CAGAR ALGUNOS DATOS DE LAS CRIATURAS
'               y de los objetos.
'*****************************************************************************
Option Explicit

Public Type ObjData
    Name As String 'Nombre del obj
    OBJType As Integer 'Tipo enum que determina cuales son las caract del obj
    GrhIndex As Integer ' Indice del grafico que representa el obj
End Type

Public ObjData() As ObjData

'******************************************************************************
Public Type NpcData
    Name As String
    body As Integer
    Head As Integer
    heading As Byte
End Type

Public NpcData() As NpcData
       
Public Sub cargarInformacionObjetos()
    
    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopObjeto As Integer
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DBPath & "\objetos.dat"
    
    ultimo = CInt(val(m_iniFile.getNameLastSection))
    
    
    ReDim ObjData(1 To ultimo)
    
    For loopObjeto = 1 To ultimo
        With ObjData(loopObjeto)
            .GrhIndex = CInt(val(m_iniFile.getValue(loopObjeto, "GRHINDEX")))
            .Name = m_iniFile.getValue(loopObjeto, "NAME")
            .OBJType = CInt(val(m_iniFile.getValue(loopObjeto, "OBJTYPE")))
        End With
    Next loopObjeto
    
    Set m_iniFile = Nothing
End Sub

Public Sub cargarListaObjetos()

    Dim loopObjeto As Integer
    Dim nombre As String
    Dim id As Long
    
    frmMain.ListaConBuscadorObjetos.vaciar

    For loopObjeto = 1 To UBound(ObjData)
        nombre = ObjData(loopObjeto).Name
        id = loopObjeto
        If Len(nombre) > 0 Then Call frmMain.ListaConBuscadorObjetos.addString(id, id & " - " & nombre)
    Next

End Sub

Public Sub cargarInformacionNPCs()

    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopElemento As Integer
    Dim descripcionInterna As String
    Dim domable As Boolean
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DBPath & "\npcs.dat"
    
    ultimo = CInt(val(m_iniFile.getNameLastSection))
    
    
    ReDim NpcData(1 To ultimo)
    
    For loopElemento = 1 To ultimo
        With NpcData(loopElemento)
          
            descripcionInterna = m_iniFile.getValue(loopElemento, "DescInterna")
            
            domable = (CInt(val(m_iniFile.getValue(loopElemento, "Domable"))) > 0)
            .Name = m_iniFile.getValue(loopElemento, "NAME")
            
            If Len(descripcionInterna) > 0 Or domable Then
                .Name = .Name & " (" & IIf(domable, "Domable" & IIf(Len(descripcionInterna) > 0, ". " & descripcionInterna, ""), descripcionInterna) & ")"
            End If
            
            .Head = CInt(val(m_iniFile.getValue(loopElemento, "HEAD")))
            .body = CInt(val(m_iniFile.getValue(loopElemento, "BODY")))
            .heading = CInt(val(m_iniFile.getValue(loopElemento, "HEADING")))
        End With
    Next loopElemento
    
    Set m_iniFile = Nothing
    
End Sub

Public Sub cargarListaNPC()

    Dim loopElemento As Integer
    Dim nombre As String
    Dim id As Long
    
    frmMain.ListaConBuscadorNpcs.vaciar

    For loopElemento = 1 To UBound(NpcData)
        nombre = NpcData(loopElemento).Name
        id = loopElemento
        If Len(nombre) > 0 Then Call frmMain.ListaConBuscadorNpcs.addString(id, id & " - " & nombre)
    Next

End Sub
