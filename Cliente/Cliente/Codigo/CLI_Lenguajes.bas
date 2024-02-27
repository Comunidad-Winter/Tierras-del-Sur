Attribute VB_Name = "CLI_Lenguajes"
Option Explicit


Public Function CargarLenguaje(lenguaje As String) As Boolean

 Dim mensajes As String
Dim cantidades As Variant
Dim i As Integer
Dim archivoLenguaje As Integer

On Error GoTo hayerror:

If FileExist(app.Path & "\Lenguajes\" & lenguaje & ".leg", vbArchive) Then

archivoLenguaje = FreeFile()

Open (app.Path & "\Lenguajes\" & lenguaje & ".leg") For Input As #archivoLenguaje
'Tomamos la cantidad de mensajes
Line Input #archivoLenguaje, mensajes
cantidades = Split(mensajes, ";")
'Cargamos los mensajes

ReDim mensaje(1 To cantidades(0)) As String
    For i = 1 To cantidades(0)
    Line Input #archivoLenguaje, mensajes
    mensaje(i) = mensajes
    Next
    
ReDim mapa(1 To cantidades(1)) As String
    For i = 1 To cantidades(1)
    Line Input #archivoLenguaje, mensajes
    mapa(i) = mensajes
    Next
    

ReDim ListaRazas(1 To cantidades(2)) As String
    For i = 1 To cantidades(2)
    Line Input #archivoLenguaje, mensajes
    ListaRazas(i) = mensajes
    Next
    

ReDim RangoArmada(0 To cantidades(3)) As String
    For i = 1 To cantidades(3)
    Line Input #archivoLenguaje, mensajes
    RangoArmada(i) = mensajes
    Next
    
ReDim RangoCaos(0 To cantidades(4)) As String
    For i = 1 To cantidades(4)
    Line Input #archivoLenguaje, mensajes
    RangoCaos(i) = mensajes
    Next
    
ReDim ListaClases(1 To cantidades(5)) As String
    For i = 1 To cantidades(5)
    Line Input #archivoLenguaje, mensajes
    ListaClases(i) = mensajes
    Next
    
ReDim ListaGeneros(1 To 2) As String

ListaGeneros(1) = "Hombre"
ListaGeneros(2) = "Mujer"

ReDim UserSkills(1 To NUMSKILLS) As Integer

ReDim SkillsNames(1 To cantidades(6)) As String
    For i = 1 To cantidades(6)
    Line Input #archivoLenguaje, mensajes
    SkillsNames(i) = mensajes
    Next
    
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To cantidades(7)) As String
    For i = 1 To cantidades(7)
    Line Input #archivoLenguaje, mensajes
    AtributosNames(i) = mensajes
    Next
    

ReDim Ciudades(1 To cantidades(8)) As String
    For i = 1 To cantidades(8)
    Line Input #archivoLenguaje, mensajes
    Ciudades(i) = mensajes
    Next


ReDim CityDesc(1 To cantidades(9)) As String
    For i = 1 To cantidades(9)
    Line Input #archivoLenguaje, mensajes
    CityDesc(i) = mensajes
    Next
    

ReDim objeto(1 To cantidades(10)) As String
    For i = 1 To cantidades(10)
    Line Input #archivoLenguaje, mensajes
    objeto(i) = mensajes
    Next

    For i = 1 To cantidades(11)
    Line Input #archivoLenguaje, mensajes
    tips(i) = mensajes
    Next



Close #archivoLenguaje
CargarLenguaje = True
Else
CargarLenguaje = False
End If
Exit Function
hayerror:
MsgBox Err.Description
CargarLenguaje = False

End Function
