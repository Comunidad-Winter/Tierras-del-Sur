Attribute VB_Name = "ME_Tools_Graficos"
Option Explicit

Public Const CANTIDAD_CAPAS = 5

Public Enum eHerramientaGraficos
    ninguna = 0
    insertar = 1
    borrar = 2
End Enum

Public Type tCapasPosicion
    GrhIndex  As Integer
    seleccionado As Boolean
End Type

Public herramientaInternaGraficos As eHerramientaGraficos
Public Const cantidadSubHerramienta As Byte = 2

Public grhInfoSeleccionada() As tGhInfoSeleccionada

Public Type tGhInfoSeleccionada
    grhInfoPosicion(1 To CANTIDAD_CAPAS)     As tCapasPosicion
End Type

Public Sub iniciarToolGraficos()

    Dim aux(1 To CANTIDAD_CAPAS) As tCapasPosicion
    Dim i As Integer
    
    For i = 1 To CANTIDAD_CAPAS
        aux(i).GrhIndex = 0
        aux(i).seleccionado = False
    Next i
    
    Call establecerInfoGrhPosicion(aux)
    
End Sub

Public Sub establecerInfoGrhPosicion(grHInfo() As tCapasPosicion)
    
    Dim loopCapa As Byte
    ReDim grhInfoSeleccionada(1 To 1, 1 To 1) As tGhInfoSeleccionada
    
    For loopCapa = 1 To CANTIDAD_CAPAS
        grhInfoSeleccionada(1, 1).grhInfoPosicion(loopCapa) = grHInfo(loopCapa)
    Next
    
End Sub

Public Sub establecerInfoGrhCapa(graficoIndex As Integer, nCapa As Byte)
    Dim tilesAncho As Integer
    Dim tilesAlto As Integer

    tilesAlto = -Int(GrhData(graficoIndex).pixelHeight / TilePixelHeight * (-1))
    tilesAncho = -Int(GrhData(graficoIndex).pixelWidth / TilePixelWidth * (-1))
    
    Call resetearGrhInfo(CByte(tilesAncho), CByte(tilesAlto))

    grhInfoSeleccionada(1, 1).grhInfoPosicion(nCapa).seleccionado = True
    grhInfoSeleccionada(1, 1).grhInfoPosicion(nCapa).GrhIndex = graficoIndex
End Sub

Private Sub resetearGrhInfo(ancho As Byte, alto As Byte)
    Dim loopCapa As Byte
    Dim loopX As Byte
    Dim loopY As Byte
    If ancho = 0 Then Exit Sub
    ReDim grhInfoSeleccionada(1 To ancho, 1 To alto)
    
    For loopX = 1 To ancho
        For loopY = 1 To alto
            For loopCapa = 1 To CANTIDAD_CAPAS
                grhInfoSeleccionada(loopX, loopY).grhInfoPosicion(loopCapa).seleccionado = False
                grhInfoSeleccionada(loopX, loopY).grhInfoPosicion(loopCapa).GrhIndex = 0
            Next
        Next
    Next
End Sub

Public Sub setGrhInfoBorrado()
    Dim loopCapa As Byte

    ReDim grhInfoSeleccionada(1 To 1, 1 To 1)
    
    For loopCapa = 1 To CANTIDAD_CAPAS
        grhInfoSeleccionada(1, 1).grhInfoPosicion(loopCapa).seleccionado = True
        grhInfoSeleccionada(1, 1).grhInfoPosicion(loopCapa).GrhIndex = 0
    Next
End Sub

Public Sub click_InsertarGrafico()
    Dim algunaSelect As Boolean
    Dim loopCapa As Integer
    
    algunaSelect = False
    
    For loopCapa = 1 To CANTIDAD_CAPAS
        algunaSelect = algunaSelect Or (grhInfoSeleccionada(1, 1).grhInfoPosicion(loopCapa).seleccionado = True)
    Next
    
    If (algunaSelect) Then
        herramientaInternaGraficos = eHerramientaGraficos.insertar
        Call ME_Tools.seleccionarTool(frmMain.cmdInsertarGrafico, tool_grh)
    Else
        MsgBox "Debe seleccionar un grafico y una capa donde insertarlo."
    End If
End Sub

Public Sub click_BorrarGrafico()
    herramientaInternaGraficos = eHerramientaGraficos.borrar
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarGrafico, tool_grh)
End Sub


Public Sub rotarHerramientaInterna(paraArriba As Boolean)

    If paraArriba Then
        herramientaInternaGraficos = herramientaInternaGraficos + 1
        If herramientaInternaGraficos > cantidadSubHerramienta Then herramientaInternaGraficos = 1
    Else
        herramientaInternaGraficos = herramientaInternaGraficos - 1
        If herramientaInternaGraficos < 1 Then herramientaInternaGraficos = cantidadSubHerramienta
    End If
    
    Select Case herramientaInternaGraficos
        
        Case eHerramientasAccion.insertar
            ME_Tools_Graficos.click_InsertarGrafico
        Case eHerramientasAccion.borrar
            ME_Tools_Graficos.click_BorrarGrafico
    End Select


End Sub

Public Sub activarUltimaHerramienta()

    Select Case herramientaInternaGraficos
        
        Case eHerramientasAccion.insertar
            ME_Tools_Graficos.click_InsertarGrafico
        Case eHerramientasAccion.borrar
            ME_Tools_Graficos.click_BorrarGrafico
    End Select
    
End Sub

Public Sub actualizarListaUltimosUsados(grhInfoPosicion() As tCapasPosicion)
    Dim loopElemento As Byte
    Dim loopCapa As Byte
    
    For loopCapa = 1 To CANTIDAD_CAPAS
    
        If grhInfoPosicion(loopCapa).GrhIndex > 0 Then
        
            For loopElemento = 1 To frmMain.lstUltimosGraficosUsados.ListCount
            
                If val(frmMain.lstUltimosGraficosUsados.list(loopElemento - 1)) = grhInfoPosicion(loopCapa).GrhIndex Then
                    If loopElemento <> 1 Then
                        Call frmMain.lstUltimosGraficosUsados.RemoveItem(loopElemento - 1)
                    End If
                    Exit For
                End If
                
            Next

            If Not loopElemento = 1 Or frmMain.lstUltimosGraficosUsados.ListCount = 0 Then
                If frmMain.lstUltimosGraficosUsados.ListCount = 8 Then
                    Call frmMain.lstUltimosGraficosUsados.RemoveItem(frmMain.lstUltimosGraficosUsados.ListCount - 1)
                End If
            
                Call frmMain.lstUltimosGraficosUsados.AddItem(grhInfoPosicion(loopCapa).GrhIndex & " - " & GrhData(grhInfoPosicion(loopCapa).GrhIndex).nombreGrafico, 0)
            End If
        End If
        
    Next
End Sub
