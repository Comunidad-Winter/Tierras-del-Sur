Attribute VB_Name = "modClima"
Option Explicit

Public Lloviendo As Boolean
Public Nevando As Boolean

Public Minutoslloviendo As Integer
Public Minutossinlluvia As Integer

Public Enum eClimas
    ClimaNeblina = 1
    climalluvia = 2
    ClimaNiebla = 4
    ClimaTormenta_de_arena = 8
    ClimaNublado = 16
    ClimaNieve = 32
    ClimaRayos_de_luz = 64
End Enum

Public Sub enviarNoche(ByRef personaje As User)
    If esMapaDeEvento(personaje.pos.map) Then
        EnviarPaquete Paquetes.Noche, ByteToString(fraccionDelDia) & ByteToString(0), personaje.UserIndex, ToIndex
    Else
        EnviarPaquete Paquetes.Noche, ByteToString(fraccionDelDia) & ByteToString(ForzarDia), personaje.UserIndex, ToIndex
    End If
End Sub
Public Sub calcularClima()
    
    Dim cambiaClima As Boolean
    
    ' Inicia o para la lluvia
    If Lloviendo = False Then
         If Minutossinlluvia > 30 Then
            If RandomNumber(1, 500) <= 7 Then
                cambiaClima = True
            End If
         Else
            Minutossinlluvia = Minutossinlluvia + 1
         End If
    Else
        If Minutoslloviendo > 3 Then  'Minutos
            cambiaClima = True
        Else
            Minutoslloviendo = Minutoslloviendo + 1
        End If
    End If
    
    If Not cambiaClima Then
        Exit Sub
    End If
       
    Call cambiarClima

End Sub

Public Sub cambiarClima()

    Dim mapa As Integer
    
    If Lloviendo Then
        Minutoslloviendo = 0
        Minutossinlluvia = 0
        Lloviendo = False
        Nevando = False
                
        For mapa = LBound(MapInfo) To UBound(MapInfo)
            If Not MapInfo(mapa).climaActual = 0 Then
                MapInfo(mapa).climaActual = 0
                EnviarPaquete Paquetes.lluvia, ITS(0), 0, ToMap, mapa
            End If
        Next
        
    Else
        Minutossinlluvia = 0
        Minutoslloviendo = 0
        Lloviendo = True
        Nevando = True
        
        For mapa = LBound(MapInfo) To UBound(MapInfo)
            If MapInfo(mapa).clima > 0 Then
                MapInfo(mapa).climaActual = MapInfo(mapa).clima
                EnviarPaquete Paquetes.lluvia, ITS(MapInfo(mapa).clima), 0, ToMap, mapa
            End If
        Next
    End If
    
End Sub
