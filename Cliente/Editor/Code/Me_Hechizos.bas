Attribute VB_Name = "Me_Hechizos"
Option Explicit

Type tHechizo
    nombre As String
End Type

Public HechizosData() As tHechizo

Public Sub cargarInformacionHechizos()
    
    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopElemento As Integer
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DBPath & "\hechizos.dat"
    
    ultimo = CInt(val(m_iniFile.getNameLastSection))
    
    ReDim HechizosData(1 To ultimo)
    
    For loopElemento = 1 To ultimo
        With HechizosData(loopElemento)
            .nombre = m_iniFile.getValue(loopElemento, "NOMBRE")
        End With
    Next loopElemento
    
    Set m_iniFile = Nothing
End Sub

