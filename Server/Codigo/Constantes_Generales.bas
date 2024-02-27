Attribute VB_Name = "Constantes_Generales"
Option Explicit

Public Const PRIV_USUARIO = 0
Public Const PRIV_CONSEJERO = 1
Public Const PRIV_GAMEMASTER = 2
Public Const PRIV_DIOS = 3
Public Const PRIV_ADMINISTRADOR = 4

'Este privilegio no se asigna a los personajes, sino que es el enviado al cliente
'para indicar de que color debe pintar a estos personajes
Public Const PRIV_USUARIOS_CONSEJO = 5

Public Const COMANDOS_USUARIOS = 1
Public Const COMANDOS_CONSEJEROS = 67
Public Const COMANDOS_GAMEMASTERS = 68
Public Const COMANDOS_DIOSES = 69
Public Const COMANDOS_ADMINISTRADORES = 70

Public STAT_MAXELV As Byte                   ' Máximo Nivel

Public Const STAT_MAXHP = 999
Public Const STAT_MAXSTA = 999
Public Const STAT_MAXMAN = 5000
Public Const STAT_MAXHIT = 300
Public Const STAT_MAXDEF = 99

Private EXPERIENCIA_NIVEL() As Double       ' Experiencia para pasar de nivel
Public LevelSkill() As Integer              ' Skills Naturales que puede tener cada nivel

Public Enum eTrabajos
    Ninguno = 0
    Pesca = 1
    Tala = 2
    Mineria = 3
    Fundicion = 4
    Herreria = 5
    Carpinteria = 6
End Enum

Public Type RazaConfig
    nombre As String
    atributos(1 To NUMATRIBUTOS) As Byte
End Type

Public razasConfig(1 To NUMRAZAS) As RazaConfig

Public Const PENALIZACION_CRIATURA_MENOR_NIVEL_USUARIO As Double = 0.25

Public Sub inicializarConstantes()
    Call inicializarNiveles
    Call inicializarRazas
    Call inicializarInformacionRazas
    Call inicializarSkills
End Sub

' Devuelve la cantidad de puntos de experiencia necesario para pasar el nivel
' indicado por parametro.
Public Function obtenerExperienciaNecesaria(Nivel As Integer) As Double

    If Nivel >= STAT_MAXELV Then
        obtenerExperienciaNecesaria = 0
        Exit Function
    End If

    obtenerExperienciaNecesaria = EXPERIENCIA_NIVEL(Nivel)

End Function

Private Sub inicializarAtributos()

    ReDim AtributosNames(1 To NUMATRIBUTOS) As String
    
    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"

End Sub

Private Sub inicializarSkills()
    ReDim SkillsNames(1 To NUMSKILLS) As String
    
    SkillsNames(1) = "Resistencia Mágica"
    SkillsNames(2) = "Magia"
    SkillsNames(3) = "Robar"
    SkillsNames(4) = "Tacticas de combate"
    SkillsNames(5) = "Combate con armas"
    SkillsNames(6) = "Meditar"
    SkillsNames(7) = "Apuñalar"
    SkillsNames(8) = "Ocultarse"
    SkillsNames(9) = "Supervivencia"
    SkillsNames(10) = "Talar arboles"
    SkillsNames(11) = "Comercio"
    SkillsNames(12) = "Defensa con escudos"
    SkillsNames(13) = "Pesca"
    SkillsNames(14) = "Mineria"
    SkillsNames(15) = "Carpinteria"
    SkillsNames(16) = "Herreria"
    SkillsNames(17) = "Liderazgo"
    SkillsNames(18) = "Domar animales"
    SkillsNames(19) = "Armas de proyectiles"
    SkillsNames(20) = "Wresterling"
    SkillsNames(21) = "Navegacion"
End Sub

Private Sub inicializarRazas()

    ReDim listaRazas(1 To NUMRAZAS) As String
    
    listaRazas(1) = "Humano"
    listaRazas(2) = "Elfo"
    listaRazas(3) = "Elfo Oscuro"
    listaRazas(4) = "Gnomo"
    listaRazas(5) = "Enano"
End Sub

Private Sub inicializarNiveles()
    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopElemento As Integer
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DatPath & "\niveles.dat"
    
    ultimo = CInt(val(m_iniFile.getNameLastSection))
    
    ReDim EXPERIENCIA_NIVEL(1 To ultimo)
    ReDim LevelSkill(1 To ultimo)
        
    For loopElemento = 1 To ultimo
        EXPERIENCIA_NIVEL(loopElemento) = CDbl(m_iniFile.getValue(loopElemento, "EXP"))
        LevelSkill(loopElemento) = CInt(m_iniFile.getValue(loopElemento, "SKILLS"))
    Next loopElemento
    
    STAT_MAXELV = ultimo
    
    Set m_iniFile = Nothing
End Sub

Private Sub inicializarInformacionRazas()
    
    Dim m_iniFile As cIniManager
    Dim ultimo As Integer
    Dim loopElemento As Integer
    
    Set m_iniFile = New cIniManager
    
    m_iniFile.Initialize DatPath & "\razas.dat"
        
    For loopElemento = 1 To NUMRAZAS
        With razasConfig(loopElemento)
            .nombre = m_iniFile.getValue(loopElemento, "NOMBRE")
            
            Dim loopAtributo As Byte
            
            For loopAtributo = 1 To NUMATRIBUTOS
                .atributos(loopAtributo) = m_iniFile.getValue(loopElemento, "ATRIBUTO_" & loopAtributo)
            Next
    
        End With
    Next loopElemento
    
    Set m_iniFile = Nothing
End Sub
