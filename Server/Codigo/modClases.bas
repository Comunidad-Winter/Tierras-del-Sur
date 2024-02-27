Attribute VB_Name = "modClases"
Option Explicit

Public Type claseConfig
    nombre As String
    evasion As Single
    ataqueArmas As Single
    ataqueProyectiles As Single
    danoArmas As Single
    danoProyectiles As Single
    evasionEscudo As Single
End Type

Public Const NUMCLASES = 19

Public Const CANTIDAD_CLASES = 15 'Quitando las de game masters

Public Enum eClases
    indefinido = 0 'Lo que seria el valor nulo
    Mago = 1
    Clerigo = 2
    Guerrero = 4
    asesino = 8
    Ladron = 16
    Bardo = 32
    Druida = 64
    Paladin = 128
    Cazador = 256
    Pescador = 512
    Herrero = 1024
    Leñador = 2048
    Minero = 4096
    Carpintero = 8192
    Pirata = 16384
    Consejero = 32768
    SemiDios = 65536
    Dios = 131072
    Administrador = 262144
End Enum

Public clasesConfig(1 To NUMCLASES) As claseConfig

Public Sub inicializarClases()
    Call inicializarInformacionClases
End Sub

Private Sub inicializarInformacionClases()
    
    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopElemento As Integer
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DatPath & "\clases.dat"
        
    For loopElemento = 1 To NUMCLASES
        With clasesConfig(loopElemento)
            .nombre = eliminarTildesMayus(UCase$(m_iniFile.getValue(loopElemento, "NOMBRE")))
            .evasion = val(m_iniFile.getValue(loopElemento, "EVASION")) / 100
            .ataqueArmas = val(m_iniFile.getValue(loopElemento, "ATAQUE_ARMAS")) / 100
            .ataqueProyectiles = val(m_iniFile.getValue(loopElemento, "ATAQUE_PROYECTILES")) / 100
            .danoArmas = val(m_iniFile.getValue(loopElemento, "DANO_ARMAS")) / 100
            .danoProyectiles = val(m_iniFile.getValue(loopElemento, "DANO_PROYECTILES")) / 100
            .evasionEscudo = val(m_iniFile.getValue(loopElemento, "EVASION_ESCUDO")) / 100
        End With
    Next loopElemento
    
    Set m_iniFile = Nothing
End Sub

Public Function claseToByte(clase As String) As Long
    Dim loopClase As Integer
    Dim claseNormalizada  As String
    
    claseNormalizada = UCase$(clase)
    
    For loopClase = 1 To NUMCLASES
        If clasesConfig(loopClase).nombre = claseNormalizada Then
            claseToByte = 2 ^ (loopClase - 1)
            Exit Function
        End If
    Next
End Function

Public Function claseConfigToEnum(configId As Byte) As eClases
    claseConfigToEnum = 2 ^ (configId - 1)
End Function


Public Function clasesToString(clases As Long) As String

    Dim loopClase As Byte
    Dim listaClases As String
    Dim clase As eClases

    For loopClase = 1 To NUMCLASES
        If ((2 ^ (loopClase - 1)) And clases) Then
            clasesToString = clasesToString & " " & clasesConfig(loopClase).nombre
        End If
    Next
    
End Function
Public Function byteToClase(clase As eClases) As String
    Select Case clase
        Case eClases.Mago
            byteToClase = "MAGO"
        Case eClases.Clerigo
            byteToClase = "CLERIGO"
        Case eClases.Guerrero
            byteToClase = "GUERRERO"
        Case eClases.asesino
            byteToClase = "ASESINO"
        Case eClases.Ladron
            byteToClase = "LADRON"
        Case eClases.Bardo
            byteToClase = "BARDO"
        Case eClases.Druida
            byteToClase = "DRUIDA"
        Case eClases.Paladin
            byteToClase = "PALADIN"
        Case eClases.Cazador
            byteToClase = "CAZADOR"
        Case eClases.Pescador
            byteToClase = "PESCADOR"
        Case eClases.Herrero
            byteToClase = "HERRERO"
        Case eClases.Leñador
            byteToClase = "LEÑADOR"
        Case eClases.Minero
            byteToClase = "MINERO"
        Case eClases.Carpintero
            byteToClase = "CARPINTERO"
        Case eClases.Pirata
            byteToClase = "PIRATA"
        Case eClases.Consejero
            byteToClase = "CONSEJERO"
        Case eClases.SemiDios
            byteToClase = "SEMIDIOS"
        Case eClases.Dios
            byteToClase = "DIOS"
        Case eClases.Administrador
            byteToClase = "ADMINISTRADOR"
        End Select
End Function

Public Function claseToConfigID(clase As String) As Long
    Dim loopClase As Integer
    Dim claseNormalizada  As String
    
    claseNormalizada = UCase$(clase)
    
    For loopClase = 1 To NUMCLASES
        If clasesConfig(loopClase).nombre = claseNormalizada Then
            claseToConfigID = loopClase
            Exit Function
        End If
    Next
End Function

Function ModificadorEvasion(ByVal claseConfigId As Byte) As Single
    ModificadorEvasion = clasesConfig(claseConfigId).evasion
End Function

Function ModificadorPoderAtaqueArmas(ByVal claseConfigId As Byte) As Single
    ModificadorPoderAtaqueArmas = clasesConfig(claseConfigId).ataqueArmas
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal claseConfigId As Byte) As Single
    ModificadorPoderAtaqueProyectiles = clasesConfig(claseConfigId).ataqueProyectiles
End Function

Function ModicadorDañoClaseArmas(ByVal claseConfigId As Byte) As Single
    ModicadorDañoClaseArmas = clasesConfig(claseConfigId).danoArmas
End Function

Function ModicadorDañoClaseProyectiles(ByVal claseConfigId As Byte) As Single
    ModicadorDañoClaseProyectiles = clasesConfig(claseConfigId).danoProyectiles
End Function

Function ModEvasionDeEscudoClase(ByVal claseConfigId As Byte) As Single
    ModEvasionDeEscudoClase = clasesConfig(claseConfigId).evasionEscudo
End Function
