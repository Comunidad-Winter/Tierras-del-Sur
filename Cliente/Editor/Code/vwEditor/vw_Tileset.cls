VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vw_Tileset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vWindow

Private WithEvents ListaTilesets            As ControlList
Attribute ListaTilesets.VB_VarHelpID = -1
Private WithEvents MouseReaderTileset       As ControlMouseReader
Attribute MouseReaderTileset.VB_VarHelpID = -1

Private WithEvents BotonCerrar              As ControlBoton
Attribute BotonCerrar.VB_VarHelpID = -1
Private WithEvents BotonFillMap             As ControlBoton
Attribute BotonFillMap.VB_VarHelpID = -1

Private WithEvents Controles                As vControles
Attribute Controles.VB_VarHelpID = -1

'Implementacion de ventanas. No modificar
' LA PUTA QUE LO PARIOOOOOOO PORQUE NO LO PUEDO EXTENDER Y YA!?!?. LA VIDA SERIA MEJOR SI SE PUDIERA EXTENDER UNA CLASE EN ESTE VB DE MIERDA
    Private pNext As vWindow
    Private pPrev As vWindow


Public TilesetSeleccionado                  As Integer
Public TileSetVirtualSeleccionado           As Byte ' Nuevo formato de pisos comprimid

Private Sub BotonCerrar_Click()
    Call EsconderVentanaTilesets
End Sub

Private Sub BotonFillMap_Click()

    Dim respuesta As VbMsgBoxResult
   
    respuesta = MsgBox("Se aplicará el piso seleccionado a toda el mapa. ¿Estás seguro?. Podés volver para atrás con CONTROL + Z.", vbQuestion + vbYesNo, "Aplicar piso al mapa")
    
    If respuesta = vbNo Then Exit Sub

    Call backupearEstadoHerramientas
           
    ' Generamos
    If Me_Tools_TileSet.aplicarTexturaTodoMapa(TilesetSeleccionado, Area_Tileset) Then
    
        'Ejecutamos la acción sobre esa area
        click_tool vbLeftButton
        
        ' Restablecemos
        Call restablecerBackupHerramientas
    
        ' Suponemos que estaba con los tilesets
       ' Call activarUltimaHerramientaCorrespondienteASolapa(eSolapasEditor.Tilesets)
    Else
        Call GUI_Alert("No se puede poner el piso acá. Revisa de estar haciendo las cosas bien.")
    End If
    
    Call Me_Tools_TileSet.EsconderVentanaTilesets
    
End Sub


    Private Property Get vWindow_GetControl() As vControles
        Set vWindow_GetControl = Controles
    End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function
    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
'/Implementacion

Private Sub Class_Initialize()

'Inicio los controles
    Set Controles = New vControles
    
    Set ListaTilesets = Controles.Add(New ControlList)
    Set MouseReaderTileset = Controles.Add(New ControlMouseReader)
    
    Set BotonCerrar = Controles.Add(New ControlBoton)
    Set BotonFillMap = Controles.Add(New ControlBoton)
    
' Posiciono los controles

    ListaTilesets.SetPos 532, 0
    ListaTilesets.SetSize 250, D3DWindow.BackBufferHeight
    
    MouseReaderTileset.SetColor &HFF000000
    MouseReaderTileset.SetPos 16, 0
    MouseReaderTileset.SetSize 512, 512
    
    BotonCerrar.SetPos 532 + 250 + 16, 16
    BotonFillMap.SetPos 532 + 250 + 16, 64
    
    BotonCerrar.text = "Cerrar"
    BotonFillMap.text = "Llenar mapa"
    
    ActualizarLista

End Sub


Private Sub vWindow_Hide()

End Sub

Private Sub vWindow_SetPos(ByVal x As Integer, ByVal y As Integer)

End Sub

Private Sub vWindow_Show()
    GUI_SetFocus Me
End Sub

Public Sub ActualizarLista()
    Dim i As Integer
    Dim formatoPiso As eFormatoTileSet
    Dim Index As String
 
    ListaTilesets.Clear
    TilesetSeleccionado = 0
    TileSetVirtualSeleccionado = 0
    
    For i = 1 To Tilesets_count

        If Me_indexar_Pisos.existe(i) Then
            
            formatoPiso = Tilesets(i).formato

            If Not (formatoPiso = eFormatoTileSet.camino_grande_parte1 Or formatoPiso = eFormatoTileSet.costa_tipo_1_parte1) Then
                If formatoPiso = eFormatoTileSet.camino_chico Or formatoPiso = eFormatoTileSet.camino_grande_parte2 Or formatoPiso = eFormatoTileSet.costa_tipo_1_parte2 Then
                    Call agregarAlista(i & " + " & Tilesets(i).nombre, CStr(i), True)
                Else
                    Call agregarAlista(i & " - " & Tilesets(i).nombre, CStr(i), True)
                End If
            End If

            If formatoPiso = eFormatoTileSet.camino_chico Then
                Call agregarAlista("       Camino vertical", i & "-" & 1, False)
                Call agregarAlista("       Camino horizontal", i & "-" & 2, False)
                Call agregarAlista("       Curvas", i & "-" & 3, False)
                Call agregarAlista("       Bifurcacion", i & "-" & 4, False)
            ElseIf formatoPiso = camino_grande_parte2 Then
                Call agregarAlista("       Camino vertical", i & "-" & 1, False)
                Call agregarAlista("       Camino horizontal", i & "-" & 2, False)
                Call agregarAlista("       Curvas", i & "-" & 3, False)
                Call agregarAlista("       Bifurcacion ", i & "-" & 4, False)
                Call agregarAlista("       Transiciones ", i & "-" & 5, False)
            ElseIf formatoPiso = costa_tipo_1_parte2 Then
                Call agregarAlista("       Derecha e Izquierda", i & "-" & 1, False)
                Call agregarAlista("       Superior e Inferior ", i & "-" & 2, False)
                Call agregarAlista("       Curvas Externas", i & "-" & 3, False)
                Call agregarAlista("       Curvas Internas", i & "-" & 4, False)
            End If
        End If
    Next i
    
End Sub

Private Sub agregarAlista(texto As String, Data As String, Visible As Boolean)
    Dim Index As Integer
    
    Index = ListaTilesets.AddItem(texto)
    ListaTilesets.itemData(Index) = Data
    
    If Not Visible Then ListaTilesets.ItemVisible(Index) = False
    
End Sub
Private Function vWindow_Render() As Boolean

    Dim c As Integer
    Dim loopX As Byte
    Dim loopY As Byte
    Dim xEnTextura As Integer
    Dim yEnTextura As Integer
    Dim tileEnTextura As Integer
    Dim posicionEnMatriz As Byte
    Dim deboDibujarGrilla256 As Boolean
        
    vWindow_Render = True

    ' Grilla que separa el tileset en 4 cuadrados de 256x256, depende el tileset es si se pone o no
    deboDibujarGrilla256 = False
    
    ' Fondo
    Engine.Draw_FilledBox 0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, &H8F000000, 0

    ' Renderizo los botones y lista
    ListaTilesets.Render
    MouseReaderTileset.Render

    BotonCerrar.Render
    BotonFillMap.Render
    
    ' ¿Tengo seleccionado algun tileset?
    If TilesetSeleccionado > 0 Then
    
        deboDibujarGrilla256 = Not (Tilesets(TilesetSeleccionado).formato = formato_viejo Or Tilesets(TilesetSeleccionado).formato = rocas_acuaticas)
        
        If TileSetVirtualSeleccionado = 0 Then
            c = Tilesets(TilesetSeleccionado).filenum
            Grh_Render_Simple_box c, 16, 0, &HEEFFFFFF, 512
        Else
        
            posicionEnMatriz = TileSetVirtualSeleccionado
            ' Tengo que generar el tileset dinamicamente
            With Tilesets(TilesetSeleccionado)
            
                For loopY = 0 To 15
                
                    For loopX = 0 To 15
                    
                        c = .matriz_transformacion(posicionEnMatriz, loopX, loopY).textura
                        
                        If c > 0 Then
                        
                            tileEnTextura = .matriz_transformacion(posicionEnMatriz, loopX, loopY).numero
                        
                            xEnTextura = tileEnTextura Mod 16
                            yEnTextura = tileEnTextura \ 16
                   
                            'Dibujo la parte de la textura del tileset que me interesa en donde corresponde para que se visualice
                            'el tileset virutal
                            Grh_Render_Simple_box_offset Tilesets(c).filenum, xEnTextura * 32, yEnTextura * 32, 32, 32, 16 + loopX * 32, loopY * 32, &HEEFFFFFF, 512
                        End If
                        
                    Next loopX
                    
                Next loopY
            
            End With
        End If
    
    End If

    ' Grilla
    Grh_Render_Simple_box GRILLA_TEXTURA, 16, 0, &H44FFFFFF, 512

    ' Marco una Grilla con cuadrados de 256 ya que siempre solemos trabajar con partes de 256
    If deboDibujarGrilla256 Then
        Draw_FilledBox 272, 0, 1, 512, mzWhite, mzWhite, 0
        Draw_FilledBox 16, 256, 512, 1, mzWhite, mzWhite, 0
    End If
    
    ' Area seleccionada
   ' Draw_FilledBox Area_Tileset.izquierda * 32 + 16, Area_Tileset.arriba * 32, (Area_Tileset.derecha - Area_Tileset.izquierda + 1) * 32, (Area_Tileset.abajo - Area_Tileset.arriba + 1) * 32, &HFF000000, &HFFFFFFFF, 1
    Draw_FilledBox Area_Tileset.izquierda * 32 + 16, Area_Tileset.arriba * 32, (Area_Tileset.derecha - Area_Tileset.izquierda + 1) * 32, (Area_Tileset.abajo - Area_Tileset.arriba + 1) * 32, &H33FFFFFF, &H0, 0

End Function


Private Sub seleccionarTileSetDesdeLista(listIndex As Integer)
    Dim idtileset As Integer
    Dim formato As eFormatoTileSet
    Dim itemData As String
    
    If listIndex <> -1 Then
    
        itemData = ListaTilesets.itemData(listIndex)
        
        'Obtengo el ID del tileset que representa este item
        TilesetSeleccionado = val(itemData)
        
        ' Obtengo, si tiene, el id del tileset virtual al que corresponde este item
        If InStr(1, itemData, "-", vbBinaryCompare) > 0 Then
            TileSetVirtualSeleccionado = val(mid$(itemData, InStr(1, itemData, "-", vbBinaryCompare) + 1))
        Else
            TileSetVirtualSeleccionado = 0
        End If
        
        ' Si es un elemento que es representado virtualmente, fuerzo al tileser virtual 1. Sino, no. :p
        formato = Tilesets(TilesetSeleccionado).formato
                
        If formato = eFormatoTileSet.camino_chico Or formato = eFormatoTileSet.camino_grande_parte2 Or formato = eFormatoTileSet.costa_tipo_1_parte2 Then
            If TileSetVirtualSeleccionado = 0 Then TileSetVirtualSeleccionado = 1
        End If
        
        Debug.Print "Hover Seleccione " & TilesetSeleccionado & "  Virtual: " & TileSetVirtualSeleccionado
    Else
        TilesetSeleccionado = 0
        TileSetVirtualSeleccionado = 0
    End If
        
End Sub
Private Sub ListaTilesets_HoverChange(HoverIndex As Integer)
    If HoverIndex = -1 Then
        Call seleccionarTileSetDesdeLista(ListaTilesets.SelectedIndex)
    Else
        Call seleccionarTileSetDesdeLista(HoverIndex)
    End If
End Sub

Private Sub clicSobreTileSet(SelectedIndex As Integer)
    Dim formato As eFormatoTileSet
    Dim itemData As String
    Dim loopVirtual As Byte
    Dim Visible As Boolean
    
    ' Obtenemos la informacion de Item
    itemData = ListaTilesets.itemData(SelectedIndex)
        
    'Obtengo el ID del tileset que representa este item
    TilesetSeleccionado = val(itemData)
        
    ' Obtengo, si tiene, el id del tileset virtual al que corresponde este item
    If InStr(1, itemData, "-", vbBinaryCompare) > 0 Then
        TileSetVirtualSeleccionado = val(mid$(itemData, InStr(1, itemData, "-", vbBinaryCompare) + 1))
    Else
        TileSetVirtualSeleccionado = 0
    End If
    
    formato = Tilesets(TilesetSeleccionado).formato
    
    ' ¿Es un tileset que no tiene hijos o estoy seleccionadno un hijo?
    If Not (TileSetVirtualSeleccionado > 0 Or formato = formato_viejo Or formato = textura_agua Or formato = textura_simple Or formato = rocas_acuaticas) Then
    
        ' Estoy seleccionando un padre entonces tengo que abrir o cerrar la lista de virtuales
        If InStr(1, ListaTilesets.ItemTexto(SelectedIndex), "-") > 0 Then
            ' Cerrar
            ListaTilesets.ItemTexto(SelectedIndex) = TilesetSeleccionado & " + " & Tilesets(TilesetSeleccionado).nombre
            Visible = False
        Else ' Abrir
            ListaTilesets.ItemTexto(SelectedIndex) = TilesetSeleccionado & " - " & Tilesets(TilesetSeleccionado).nombre
            Visible = True
        End If
        
        ' Aplico sobr elos elementos del mismo tileset
        For loopVirtual = 1 To obtenerCantidadVirtuales(formato)
            ListaTilesets.ItemVisible(SelectedIndex + loopVirtual) = Visible
        Next
           
        ListaTilesets.SelectedIndex = SelectedIndex + 1
        
        ' Seleccionamos el primero
        tileset_actual = TilesetSeleccionado
        TileSetVirtualSeleccionado = 1
    End If
    
    tileset_actual = TilesetSeleccionado
    tileset_actual_virtual = TileSetVirtualSeleccionado
    
    If Area_Tileset.arriba - Area_Tileset.abajo > 0 And Area_Tileset.derecha - Area_Tileset.izquierda > 0 Then
        Call establecerAreaTileSet(TilesetSeleccionado, TileSetVirtualSeleccionado, Area_Tileset)
    End If
    
    ' Si esta Abriero lo cierro y sino al revez
    Debug.Print "Click Virtual " & tileset_actual & " Virtual " & TileSetVirtualSeleccionado
End Sub
Private Sub ListaTilesets_Change(SelectedIndex As Integer)
   
    If Not SelectedIndex = -1 Then Call clicSobreTileSet(SelectedIndex)

End Sub

Private Sub MouseReaderTileset_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    Call modSeleccionArea.puntoArea(Area_Tileset, minl(x \ 32, 16), minl(y \ 32, 16))
End Sub

Private Sub MouseReaderTileset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    If Button Then
        Call modSeleccionArea.actualizarArea(Area_Tileset, minl(x \ 32, 16), minl(y \ 32, 16))
    End If
End Sub

Private Sub MouseReaderTileset_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    Call modSeleccionArea.actualizarArea(Area_Tileset, CInt(minl(x \ 32, 16)), CInt(minl(y \ 32, 16)))
    Call establecerAreaTileSet(TilesetSeleccionado, TileSetVirtualSeleccionado, Area_Tileset)
End Sub

