Attribute VB_Name = "ME_Mapas"
Option Explicit

Public mapDataConfig As cFileJSON
Public mapDataFile As cFileINI
    
    
Public Sub cargarInformacionMapas()
    
    'Cargamos el archivo de configuracion
    Set mapDataConfig = New cFileJSON
    mapDataConfig.init DBPath & "\JSON\mapa.json"
    
    'Cargamos el archivo donde esta la info de los objetos
    Set mapDataFile = New cFileINI
    mapDataFile.load DBPath & "\mapas.dat", mapDataConfig

End Sub

Public Sub cargarInformacionDeMapa(ByVal numero As Integer, Mapa As MapInfo)
    Dim tempLong As Long
    Dim tempStr As String
    Dim tempStrings() As String
    Dim loopElemento As Long
    
    Dim infoMapa As cSection
    
    Set infoMapa = mapDataFile.getSectionByName(numero)
    
    If Not infoMapa Is Nothing Then
        
        
        Mapa.Name = infoMapa.getItemByName("NOMBRE").getValue
        
        Mapa.Music = val(infoMapa.getItemByName("MUSICA").getValue)
        
        'Cargo los climas
        tempStr = infoMapa.getItemByName("CLIMA").getValue
        tempStrings = Split(tempStr, ",")
        
        tempLong = 0
        
        For loopElemento = LBound(tempStrings) To UBound(tempStrings)
            tempLong = tempLong Or CLng(val(tempStrings(loopElemento)))
        Next
    
        Mapa.puede_niebla = (tempLong And Tipos_Clima.ClimaNiebla)
        Mapa.puede_neblina = (tempLong And Tipos_Clima.ClimaNeblina)
        Mapa.puede_nieve = (tempLong And Tipos_Clima.ClimaNieve)
        Mapa.puede_lluvia = (tempLong And Tipos_Clima.ClimaLluvia)
        Mapa.puede_sandstorm = (tempLong And Tipos_Clima.ClimaTormenta_de_arena)
        Mapa.puede_nublado = (tempLong And Tipos_Clima.ClimaNublado)
            
    Else
        Mapa.Music = 0
        Mapa.Name = ""
        Mapa.puede_neblina = False
        Mapa.puede_niebla = False
        Mapa.puede_nieve = False
        Mapa.puede_nublado = False
        Mapa.puede_sandstorm = False
        Mapa.puede_lluvia = False
    End If
End Sub

