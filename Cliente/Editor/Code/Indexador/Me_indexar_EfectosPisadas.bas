Attribute VB_Name = "Me_indexar_EfectosPisadas"
Option Explicit

Public Type tEfectoPisada
    sonido_derecha As Integer
    sonido_izquierda As Integer
    
    #If esME = 1 Then
        nombre As String
    #End If
End Type

Public EfectosPisadas() As tEfectoPisada

Public Sub cargarInformacionEfectosPisadas()
    
    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopElemento As Integer
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DBPath & "\pisadas.ini"
    
    ultimo = CInt(val(m_iniFile.getNameLastSection))
    
    ReDim EfectosPisadas(1 To ultimo)
    
    For loopElemento = 1 To ultimo
        With EfectosPisadas(loopElemento)
            .nombre = m_iniFile.getValue(loopElemento, "NOMBRE")
            .sonido_derecha = val(m_iniFile.getValue(loopElemento, "DER"))
            .sonido_izquierda = val(m_iniFile.getValue(loopElemento, "IZQ"))
        End With
    Next loopElemento
    
    Set m_iniFile = Nothing
End Sub


Public Function existe(ByVal id As Integer) As Boolean
    Dim direccion As Byte
    
    existe = True
        
    If id > UBound(EfectosPisadas) Then
        existe = False
        Exit Function
    End If
    
    If EfectosPisadas(id).nombre = "" Then
        existe = False
    End If

End Function

Public Function compilar() As Boolean

End Function


