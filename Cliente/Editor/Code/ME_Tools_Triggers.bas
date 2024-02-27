Attribute VB_Name = "ME_Tools_Triggers"
Option Explicit

Public Enum herramientasBloqueo
    InsertarSimple = 1
    InsertarDoble = 2
    BorrarSimple = 3
End Enum

Public Enum herramientasTriggers
    ninguna = 0
    insertar = 1
    borrar = 2
End Enum

Private Const cantidadSubHerramientaBloqueo As Byte = 3 ' De 1 a ...
Private Const cantidadSubHerramientaTrigger As Byte = 2 ' De 1 a ...

Public herramientaInternaBloqueo As herramientasBloqueo

Public herramientaInternaTrigger As herramientasTriggers
Public triggerSeleccionado() As Long

Public Type tTriggerData
    nombre As String
    bitAfectado As Byte
    descripcion As String
    abreviatura As String * 2
End Type

Private TriggersDisponibles() As tTriggerData

Public Sub iniciarToolTrigger()
    herramientaInternaTrigger = herramientasTriggers.ninguna
    Call establecerTrigger(0)
End Sub

Private Function isTriggerSeleccionado() As Boolean

Dim x As Integer
Dim y As Integer

For x = LBound(triggerSeleccionado, 1) To UBound(triggerSeleccionado, 1)
    For y = LBound(triggerSeleccionado, 2) To UBound(triggerSeleccionado, 2)
        If triggerSeleccionado(x, y) > 0 Then
            isTriggerSeleccionado = True
            Exit Function
        End If
    Next y
Next x

isTriggerSeleccionado = False

End Function
Public Sub establecerTrigger(Trigger As Long)
    ReDim triggerSeleccionado(1 To 1, 1 To 1) As Long
    triggerSeleccionado(1, 1) = Trigger
End Sub
Public Sub click_InsertarTrigger()
    If isTriggerSeleccionado Then
        herramientaInternaTrigger = herramientasTriggers.insertar
        Call ME_Tools.seleccionarTool(frmMain.cmdInsertarTrigger, Tools.tool_triggers)
    Else
        herramientaInternaTrigger = herramientasTriggers.ninguna
        Call ME_Tools.deseleccionarTool
    End If
End Sub

Public Sub click_BorrarTrigger()
    herramientaInternaTrigger = herramientasTriggers.borrar
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarTrigger, Tools.tool_triggers)
End Sub
Public Sub rotarHerramientaInternaTrigger(paraArriba As Boolean)
    
    If paraArriba Then
        herramientaInternaTrigger = herramientaInternaTrigger + 1
        If herramientaInternaTrigger > cantidadSubHerramientaTrigger Then herramientaInternaTrigger = 1
    Else
        herramientaInternaTrigger = herramientaInternaTrigger - 1
        If herramientaInternaTrigger < 1 Then herramientaInternaTrigger = cantidadSubHerramientaTrigger
    End If

    
    Call ME_Tools_Triggers.activarUltimaHerramientaTriggers

End Sub

'*****************************************************************
'Triggers

Public Sub LoadTriggersRaw(ByRef FileName As String)
    '*****************************************************************
    'Menduz
    '*****************************************************************
    Dim cantidad As Byte
    Dim loopC As Integer
    Dim tempAbreviatura As String
     
    cantidad = val(GetVar(FileName, "Triggers", "Cantidad"))

    ReDim TriggersDisponibles(0 To cantidad - 1)
       
    For loopC = 0 To cantidad - 1
    
       TriggersDisponibles(loopC).nombre = GetVar(FileName, "Trigger" & CStr(loopC + 1), "Nombre")
       TriggersDisponibles(loopC).descripcion = GetVar(FileName, "Trigger" & CStr(loopC + 1), "Descripcion")
       TriggersDisponibles(loopC).bitAfectado = CByte(GetVar(FileName, "Trigger" & CStr(loopC + 1), "Bit"))
       
       tempAbreviatura = GetVar(FileName, "Trigger" & CStr(loopC + 1), "Abreviatura")
       
       If Len(tempAbreviatura) = 0 Then
         TriggersDisponibles(loopC).abreviatura = Space(2)
       Else
        TriggersDisponibles(loopC).abreviatura = tempAbreviatura
       End If
       
    Next loopC
End Sub

Public Sub cargarTriggersALista(lista As ListBox)
    Dim loopC As Integer
     
    lista.Clear

    For loopC = 0 To UBound(TriggersDisponibles)
        If Len(Trim(TriggersDisponibles(loopC).abreviatura)) = 0 Then
            lista.AddItem TriggersDisponibles(loopC).nombre
        Else
            lista.AddItem TriggersDisponibles(loopC).nombre & " (" & TriggersDisponibles(loopC).abreviatura & ")"
        End If
    Next loopC
End Sub

Public Function obtenerDescripcionAbreviatura(Trigger As Long) As String

    Dim loopC As Integer
    obtenerDescripcionAbreviatura = ""

    For loopC = 0 To UBound(TriggersDisponibles)
        If (Trigger And bitwisetable(TriggersDisponibles(loopC).bitAfectado)) Then
            If Len(Trim(TriggersDisponibles(loopC).abreviatura)) > 0 Then
                obtenerDescripcionAbreviatura = obtenerDescripcionAbreviatura & TriggersDisponibles(loopC).abreviatura
            End If
        End If
    Next loopC

End Function

Public Function obtenerDescripcion(numeroTipoTrigger As Byte) As String
    obtenerDescripcion = TriggersDisponibles(numeroTipoTrigger).descripcion
End Function

Public Function calcular_trigger_lista(lstTriggers As ListBox) As Long
    Dim i As Integer
    
    calcular_trigger_lista = 0
    
    For i = 0 To lstTriggers.ListCount - 1
        If lstTriggers.Selected(i) Then
            calcular_trigger_lista = calcular_trigger_lista Or bitwisetable(TriggersDisponibles(i).bitAfectado)
        End If
    Next i
End Function

Public Sub activarUltimaHerramientaTriggers()
    Select Case herramientaInternaTrigger
        Case herramientasTriggers.insertar
            Call ME_Tools_Triggers.click_InsertarTrigger
        Case herramientasTriggers.borrar
            Call ME_Tools_Triggers.click_BorrarTrigger
    End Select
End Sub


'******************************************************************
'Bloqueos
Public Sub click_InsertarBloqueo()
    herramientaInternaBloqueo = herramientasBloqueo.InsertarSimple
    Call ME_Tools.seleccionarTool(frmMain.cmdInsertarBloqueo, tool_bloqueo)
End Sub

Public Sub click_InsertarDobleBloqueo()
    herramientaInternaBloqueo = herramientasBloqueo.InsertarDoble
    Call ME_Tools.seleccionarTool(frmMain.cmdInsertarDobleBloqueo, tool_bloqueo)
End Sub

Public Sub click_BorrarBloqueo()
    herramientaInternaBloqueo = herramientasBloqueo.BorrarSimple
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarBloqueo, tool_bloqueo)
End Sub
    
    
Public Sub rotarHerramientaInternaBloqueo(paraArriba As Boolean)

    
    If paraArriba Then
        herramientaInternaBloqueo = herramientaInternaBloqueo + 1
        If herramientaInternaBloqueo > cantidadSubHerramientaBloqueo Then herramientaInternaBloqueo = 1
    Else
        herramientaInternaBloqueo = herramientaInternaBloqueo - 1
        If herramientaInternaBloqueo < 1 Then herramientaInternaBloqueo = cantidadSubHerramientaBloqueo
    End If

    
    Select Case herramientaInternaBloqueo
        
        Case herramientasBloqueo.InsertarSimple
            ME_Tools_Triggers.click_InsertarBloqueo
        Case herramientasBloqueo.InsertarDoble
            ME_Tools_Triggers.click_InsertarDobleBloqueo
        Case herramientasBloqueo.BorrarSimple
            ME_Tools_Triggers.click_BorrarBloqueo
    
    End Select


End Sub

Public Sub activarUltimaHerramientaBloqueo()

    Select Case herramientaInternaBloqueo
        
        Case herramientasBloqueo.InsertarSimple
            ME_Tools_Triggers.click_InsertarBloqueo
        Case herramientasBloqueo.InsertarDoble
            ME_Tools_Triggers.click_InsertarDobleBloqueo
        Case herramientasBloqueo.BorrarSimple
            ME_Tools_Triggers.click_BorrarBloqueo
    
    End Select
    
End Sub

Public Function BloquearTile(ByVal x As Byte, ByVal y As Byte, Optional ByVal DobleBloqueo As Boolean = True) As iComando
    '*****************************************************************
    'Menduz
    '*****************************************************************
    
    Dim comando As cComandoInsertarTrigger
    Set comando = New cComandoInsertarTrigger
    
    
    Call comando.crear(CInt(x), CInt(y), mapdata(x, y).Trigger Or eTriggers.TodosBordesBloqueados)
        
    If DobleBloqueo Then
        Dim comandoCompuesto As cComandoCompuesto
        Set comandoCompuesto = New cComandoCompuesto
        
        Call comandoCompuesto.SetNombre("Bloqueo tile doble (" & x & "," & y & ")")
        Call comandoCompuesto.agregarComando(comando)
        
        Call comandoCompuesto.agregarComando(ME_Tools_Triggers.BloquearLinea(x - 1, y, EAST, False))
        Call comandoCompuesto.agregarComando(BloquearLinea(x + 1, y, WEST, False))
        Call comandoCompuesto.agregarComando(BloquearLinea(x, y - 1, SOUTH, False))
        Call comandoCompuesto.agregarComando(BloquearLinea(x, y + 1, NORTH, False))
        
        Set BloquearTile = comandoCompuesto
    Else
        Set BloquearTile = comando
    End If
End Function

Public Function DesBloquearTile(ByVal x As Byte, ByVal y As Byte) As iComando
    '*****************************************************************
    'Menduz
    '*****************************************************************
    
    Set DesBloquearTile = Nothing
    
    If InMapBounds(x, y) Then
        Dim comando As cComandoInsertarTrigger
        Set comando = New cComandoInsertarTrigger
        
        Call comando.crear(CInt(x), CInt(y), mapdata(x, y).Trigger And Not eTriggers.TodosBordesBloqueados)
        
        Set DesBloquearTile = comando
    End If
End Function

Public Function bloquearLineaArea(area As tAreaSeleccionada, Optional ByVal heading As E_Heading = E_Heading.NONE, Optional ByVal lineaDoble As Boolean = True) As iComando
    
    Dim x As Integer
    Dim y As Integer
    Dim descripcion As String
    Dim comando As cComandoCompuesto
    
    Set comando = New cComandoCompuesto

    If area.abajo = area.arriba And area.derecha = area.izquierda Then
        descripcion = "Bloquear (" & area.izquierda & "," & area.arriba & ")"
    Else
        descripcion = "Bloquear desde (" & area.izquierda & "," & area.arriba & ") a (" & area.derecha & "," & area.abajo & ")"
    End If
    
    Call comando.crear(Nothing, descripcion)
    
    For x = area.izquierda To area.derecha
        For y = area.arriba To area.abajo
            Call comando.agregarComando(BloquearLinea(x, y, heading, lineaDoble))
        Next y
    Next x

    Set bloquearLineaArea = comando

End Function
Public Function BloquearLinea(ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As E_Heading = E_Heading.NONE, Optional ByVal HagoSegundaLlamada As Boolean = True) As iComando
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long
            
    If ME_Mundo.puedeModificarComporamientoTile(CInt(x), CInt(y)) Then
    
        If heading = E_Heading.NONE Then
           Set BloquearLinea = BloquearTile(x, y)
        Else
            tTrigger = mapdata(x, y).Trigger
            
            Dim comando As cComandoInsertarTrigger
            Dim comandoCompuesto As cComandoCompuesto
            Set comando = New cComandoInsertarTrigger
    
            If HagoSegundaLlamada Then
                Set comandoCompuesto = New cComandoCompuesto
                Call comandoCompuesto.SetNombre("Bloqueo doble (" & x & "," & y & ")")
                
                Set BloquearLinea = comandoCompuesto
            Else
                Set BloquearLinea = comando
            End If
            
            Select Case heading
                Case E_Heading.EAST
                    tTrigger = (tTrigger Or eTriggers.BloqueoEste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call comandoCompuesto.agregarComando(ME_Tools_Triggers.BloquearLinea(x + 1, y, WEST, False))
                        
                Case E_Heading.WEST
                    tTrigger = (tTrigger Or eTriggers.BloqueoOeste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call comandoCompuesto.agregarComando(ME_Tools_Triggers.BloquearLinea(x - 1, y, EAST, False))
                        
                Case E_Heading.NORTH
                    tTrigger = (tTrigger Or eTriggers.BloqueoNorte)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call comandoCompuesto.agregarComando(ME_Tools_Triggers.BloquearLinea(x, y - 1, SOUTH, False))
                        
                Case E_Heading.SOUTH
                    tTrigger = (tTrigger Or eTriggers.BloqueoSur)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call comandoCompuesto.agregarComando(ME_Tools_Triggers.BloquearLinea(x, y + 1, NORTH, False))
                        
            End Select
            
            If HagoSegundaLlamada Then
                Call comandoCompuesto.agregarComando(comando)
            End If
            
            Call comando.crear(CInt(x), CInt(y), tTrigger)
             
        End If
    Else
        Set BloquearLinea = Nothing
    End If
End Function

Private Sub bloquearExpansivoPos(posicion As cPosition, posiciones As Collection)
    Dim posicionNuevaParaAnalizar As cPosition
    Dim loopX As Integer
    Dim loopY As Integer
    
    ' ¿Dentro de los limites?
    If posicion.x < X_MINIMO_USABLE Or posicion.x > X_MAXIMO_USABLE Then Exit Sub
    If posicion.y < Y_MINIMO_USABLE Or posicion.y > Y_MAXIMO_USABLE Then Exit Sub
        
    ' Ya esta bloqueado o lo marque para bloquear?. Si es así, paro, sino bloqueo y reviso alrededor
    If (mapdata(posicion.x, posicion.y).Trigger And eTriggers.TodosBordesBloqueados) = 0 And triggerSeleccionado(posicion.x, posicion.y) = 0 Then
    
        ' Marcamos para bloquear
        triggerSeleccionado(posicion.x, posicion.y) = eTriggers.TodosBordesBloqueados

        For loopX = posicion.x - 1 To posicion.x + 1
        
            For loopY = posicion.y - 1 To posicion.y + 1
                
                ' Evitamos procesar la actual
                If Not (posicion.x = loopX And posicion.y = loopY) Then
                        
                        ' Creamos la nueva posicion a analizar
                        Set posicionNuevaParaAnalizar = New cPosition
                        posicionNuevaParaAnalizar.x = loopX
                        posicionNuevaParaAnalizar.y = loopY
                        Call posiciones.Add(posicionNuevaParaAnalizar)
                        
                End If
            
            Next
        
        Next

    End If

End Sub
' A partir de una posicion hace un bloqueo expansivo hasta ser encerrado por bloqueos
Public Sub generarBloqueoExpansivo(ByVal x As Byte, ByVal y As Byte)
    
    Dim posicionesPendientes As Collection 'Posiciones que tengo que analizar
    Dim posicion As cPosition ' Posicion actual a analizar
    
    ' Generamos el espacio para guardar los bloqueos
    ReDim triggerSeleccionado(X_MINIMO_USABLE To X_MAXIMO_USABLE, Y_MINIMO_USABLE To Y_MAXIMO_USABLE)
    
    Set posicionesPendientes = New Collection
    
    ' Posicion incial
    Set posicion = New cPosition
    posicion.x = x
    posicion.y = y
    
    ' Expando
    Call bloquearExpansivoPos(posicion, posicionesPendientes)
        
    ' Mientras haya posiciones que debo analizar
    Do While posicionesPendientes.count > 0
        
        ' Obtengo la posicion
        Set posicion = posicionesPendientes.item(1)
        posicionesPendientes.Remove (1)
        
        ' Bloqueamos y expandimos
        Call bloquearExpansivoPos(posicion, posicionesPendientes)
    Loop

    'Seleccionamos el area donde vamos a trabajar
    Call modSeleccionArea.puntoArea(ME_Tools.areaSeleccionada, X_MINIMO_USABLE, Y_MINIMO_USABLE)
    Call modSeleccionArea.actualizarArea(ME_Tools.areaSeleccionada, X_MAXIMO_USABLE, Y_MAXIMO_USABLE)

    ME_Tools_Triggers.herramientaInternaTrigger = herramientasTriggers.insertar
    Call selectToolMultiple(Tools.tool_triggers, "Aplicar bloqueo a área")
End Sub

Public Function desBloquearLineaArea(area As tAreaSeleccionada, Optional ByVal heading As E_Heading = E_Heading.NONE, Optional ByVal lineaDoble As Boolean = True) As iComando
    
    Dim x As Integer
    Dim y As Integer
    Dim descripcion As String
    Dim comando As cComandoCompuesto
    
    Set comando = New cComandoCompuesto

    If area.abajo = area.arriba And area.derecha = area.izquierda Then
        descripcion = "DesBloquear (" & area.izquierda & "," & area.arriba & ")"
    Else
        descripcion = "DesBloquear desde (" & area.izquierda & "," & area.arriba & ") a (" & area.derecha & "," & area.abajo & ")"
    End If
    
    Call comando.crear(Nothing, descripcion)
    
    For x = area.izquierda To area.derecha
        For y = area.arriba To area.abajo
            Call comando.agregarComando(DesBloquearLinea(x, y, heading, lineaDoble))
        Next y
    Next x

    Set desBloquearLineaArea = comando

End Function

Public Function DesBloquearLinea(ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As E_Heading = NONE, Optional ByVal HagoSegundaLlamada As Boolean = True) As iComando
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long
    
    If InMapBounds(x, y) Then
        If heading = E_Heading.NONE Then
            Set DesBloquearLinea = DesBloquearTile(x, y)
        Else
        
            Dim comando As cComandoInsertarTrigger
            Dim comandoCompuesto As cComandoCompuesto
            Set comando = New cComandoInsertarTrigger
            
            If HagoSegundaLlamada Then
                Set comandoCompuesto = New cComandoCompuesto
                Call comandoCompuesto.SetNombre("Borrar Bloqueo doble en (" & x & "," & y & ")")
                Set DesBloquearLinea = comandoCompuesto

            Else
                Set DesBloquearLinea = comando
            End If
        
            tTrigger = mapdata(x, y).Trigger

            Select Case heading
                Case E_Heading.EAST
                    tTrigger = (tTrigger And Not eTriggers.BloqueoEste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then
                        Call comandoCompuesto.agregarComando(ME_Tools_Triggers.DesBloquearLinea(x + 1, y, WEST, False))
                    End If
                        
                Case E_Heading.WEST
                    tTrigger = (tTrigger And Not eTriggers.BloqueoOeste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then
                        Call comandoCompuesto.agregarComando(ME_Tools_Triggers.DesBloquearLinea(x - 1, y, EAST, False))
                    End If
                    
                Case E_Heading.NORTH
                    tTrigger = (tTrigger And Not eTriggers.BloqueoNorte)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then
                         Call comandoCompuesto.agregarComando(ME_Tools_Triggers.DesBloquearLinea(x, y - 1, SOUTH, False))
                    End If
                        
                Case E_Heading.SOUTH
                    tTrigger = (tTrigger And Not eTriggers.BloqueoSur)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then
                         Call comandoCompuesto.agregarComando(ME_Tools_Triggers.DesBloquearLinea(x, y + 1, NORTH, False))
                    End If
                        
            End Select
            
            If HagoSegundaLlamada Then
                Call comandoCompuesto.agregarComando(comando)
            End If
            
            Call comando.crear(CInt(x), CInt(y), tTrigger)
             
        End If
    End If
End Function
'*****************************************************************
