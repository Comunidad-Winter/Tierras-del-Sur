VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_iniFilePath As String
Private m_iniFile As cIniManager

Private m_sections() As cSection
Private m_sectionCount As Long

Private m_fileJson As cFileJSON

Public Function getSections() As cSection()
    getSections = m_sections
End Function

Public Function getSectionCount() As Long
    getSectionCount = m_sectionCount
End Function

Public Function getSectionByListIndex(ByVal li As Long) As cSection
    If li < 0 Or li > m_sectionCount Then
        Set getSectionByListIndex = Nothing
        Exit Function
    End If
    
    Set getSectionByListIndex = m_sections(li)
End Function

Public Function getSectionByID(ByVal id As Long) As cSection
    Dim i As Long
    
    For i = 1 To m_sectionCount
        If m_sections(i).getID = id Then
            Set getSectionByID = m_sections(i)
            Exit Function
        End If
    Next i
    
    ':(
    Set getSectionByID = Nothing
End Function

Public Function getSectionByName(ByVal Name As String) As cSection
    Dim i As Long
    
    For i = 1 To m_sectionCount
        If UCase$(m_sections(i).getName) = UCase$(Name) Then
            Set getSectionByName = m_sections(i)
            Exit Function
        End If
    Next i
    
    ':(
    Set getSectionByName = Nothing
End Function

Public Function newSection(ByRef Name As String) As cSection
    m_sectionCount = m_sectionCount + 1
    ReDim Preserve m_sections(1 To m_sectionCount)
    
    Set m_sections(m_sectionCount) = New cSection
    Call m_sections(m_sectionCount).init(m_fileJson, m_iniFile, m_sectionCount, Name)
    
    Set newSection = m_sections(m_sectionCount)
End Function

Public Sub delSection(ByRef Name As String)
    Dim i As Long
    
    For i = 1 To m_sectionCount
        If UCase$(m_sections(i).getName) = UCase$(Name) Then
            Exit For
        End If
    Next i
    
    If i > m_sectionCount Then Exit Sub
    m_iniFile.EraseMain m_sections(i).getName
    
    If i <> m_sectionCount Then
        Set m_sections(i) = m_sections(m_sectionCount)
    End If
    
    Set m_sections(m_sectionCount) = Nothing
    m_sectionCount = m_sectionCount - 1
            
End Sub


Public Function load(strFile As String, ByRef fileJson As cFileJSON) As Boolean
On Error GoTo errHandler:
    Dim mainKeys() As String
    Dim i As Long
    
    m_iniFilePath = strFile
    'Iniciamos el archivo
    Set m_fileJson = fileJson
    
    'Creamos nuestra ayuda de persistencia
    Set m_iniFile = New cIniManager
    
    'Inicializamos con el archivo
    m_iniFile.Initialize strFile
    
    'Obtenemos todas las secciones qeu tiene
    mainKeys = m_iniFile.GetMainKeys()
    
    'Guardamos la cantidad
    m_sectionCount = UBound(mainKeys) + 1
    
    'Redimensionamos
    ReDim m_sections(1 To m_sectionCount)
    
    For i = 1 To m_sectionCount
        Set m_sections(i) = New cSection
        Call m_sections(i).init(fileJson, m_iniFile, m_sectionCount, mainKeys(i - 1))
    Next i
    
    load = True
    Exit Function
    
errHandler:
    Debug.Print "Error al cargar INI" & strFile
    load = False
    
End Function

Public Sub save(Optional ByVal strFile As String = vbNullString)
    If strFile = vbNullString Then
        strFile = m_iniFilePath
    End If
    
    Dim i As Long
    For i = 1 To m_sectionCount
        Call m_sections(i).saveInFile
    Next i
    
    m_iniFile.DumpFile strFile
End Sub

Private Sub Class_Terminate()
    Erase m_sections
End Sub
