VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoRemplazarGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private buscar As String
Private remplazo As String
Private Capa As Byte


Public Sub crear(ByVal buscar_ As String, ByVal remplazar_ As String, ByVal capa_ As Byte)
    buscar = buscar_
    remplazo = remplazar_
    Capa = capa_
End Sub

Private Sub iComando_desHacer()
    Call remplazar(remplazo, buscar, Capa)
End Sub

Private Sub iComando_hacer()
    Call remplazar(buscar, remplazo, Capa)
End Sub

Private Function iComando_obtenerNombre() As String
    iComando_obtenerNombre = "Remplazar " & buscar & " por " & remplazo & " en capa " & Capa
End Function

Private Sub remplazar(ByVal buscar As String, ByVal remplazo As String, ByVal Capa As Byte)
    
    ' Pasos
    ' 1. Busco todos los graficos que tengan el texto que voy a remplazo (para optimizar)
    ' 2. Recorro el mapa revisando la capa solicitada.
    ' 3. Si machea con lo buscado.
    '   a) Remplazo el nombre original el nuevo nombre y lo busco en la lista que genere prevaiemtne para obtener el id del grafico
    '   b) Remplazo en el mapa con el id del nuevo grafico
    Dim loopX As Integer
    Dim loopY As Integer
    Dim loopGrh As Integer
    
    Dim nombreNuevo As String
    Dim idNuevo As Integer
    Dim graficoData As cGrafico
    
    Dim posiblesGraficos As Collection
    
    Set posiblesGraficos = New Collection
    
    ' Buscamos todos los graficos que pueda necesitar para optimizar y no tener que recorrer la lista todo el tiempo
    For loopGrh = 1 To UBound(GrhData)
        If Me_indexar_Graficos.existe(loopGrh) Then
            
            If InStr(1, GrhData(loopGrh).nombreGrafico, remplazo, vbTextCompare) Then
            
                ' Agrego
                Set graficoData = New cGrafico
                
                graficoData.nombre = GrhData(loopGrh).nombreGrafico
                graficoData.id = loopGrh
                
                Call posiblesGraficos.Add(graficoData)
            End If ' Omito
        
        End If
    Next
    

    For loopX = SV_Constantes.X_MINIMO_USABLE To SV_Constantes.X_MAXIMO_USABLE
        For loopY = SV_Constantes.Y_MINIMO_USABLE To SV_Constantes.Y_MAXIMO_USABLE
            
            
                '¿Hay grafico?
                If mapdata(loopX, loopY).Graphic(Capa).GrhIndex > 0 Then
                    
                    With mapdata(loopX, loopY).Graphic(Capa)
                        ' ¿Coincide
                        If InStr(1, GrhData(.GrhIndex).nombreGrafico, buscar, vbTextCompare) > 0 Then
                        
                        ' Obtengo el nombre del grafico que aparentemente lo va a remplazo
                        nombreNuevo = Replace$(GrhData(.GrhIndex).nombreGrafico, buscar, remplazo, 1, -1, vbTextCompare)
                        
                        idNuevo = 0
                        ' Obtengo el identificador
                        For Each graficoData In posiblesGraficos
                        
                            If graficoData.nombre = nombreNuevo Then
                                idNuevo = graficoData.id
                                Exit For
                            End If
                        Next
    
                        ' Remplazo
                        If idNuevo > 0 Then
                            Debug.Print " Cambio " & GrhData(.GrhIndex).nombreGrafico & " por " & GrhData(idNuevo).nombreGrafico & "."
                            .GrhIndex = idNuevo
                        Else
                            MsgBox "No se encontró. " & nombreNuevo
                        End If
                        End If
                   End With
                   
                   ' Lo iniciamos en el mapa
                    InitGrh mapdata(loopX, loopY).Graphic(Capa), mapdata(loopX, loopY).Graphic(Capa).GrhIndex
                End If
        Next loopY
    Next loopX
        
End Sub




