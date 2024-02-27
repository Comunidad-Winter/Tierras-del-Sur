Attribute VB_Name = "Me_indexar_Pixels"
Option Explicit

Public Type tPixelShader
    codigo As String
    #If esME = 1 Then
        nombre As String
    #End If
End Type

Public PixelShaders() As tPixelShader

Public Function compilar() As Boolean
    
        
End Function

Public Sub cargarDesdeIni()

    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopElemento As Integer
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DBPath & "\pixels.dat"
    
    ultimo = CInt(val(m_iniFile.getNameLastSection))
        
    ReDim PixelShaders(1 To ultimo) As tPixelShader
    
    For loopElemento = 1 To ultimo
        With PixelShaders(loopElemento)
            .codigo = m_iniFile.getValue(loopElemento, "COD")
            .nombre = m_iniFile.getValue(loopElemento, "NOMBRE")
        End With
    Next loopElemento
    
    Set m_iniFile = Nothing
    
End Sub
