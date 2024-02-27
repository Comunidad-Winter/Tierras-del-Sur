Attribute VB_Name = "Me_Tools_Entidades"
Option Explicit

'Objetos
Public Enum eHerramientasEntidades
    ninguna = 0
    insertar = 1
    borrar = 2
End Enum

Private Const cantidadSubHerramientaEntidades As Byte = 2 ' De 1 a ...

Public Type tEntidadSeleccionada
    IndexEntidad As Integer
    accion As iAccionEditor
    posicion As Byte
End Type

Public Type tListaEntidadesSeleccionadas
    infoEntidades() As tEntidadSeleccionada
End Type

Public herramientaInternaEntidades As eHerramientasEntidades
Public entidadesSeleccionadas() As tListaEntidadesSeleccionadas
Public entidadesSeleccionadasBorrado As tListaEntidadesSeleccionadas

Public Sub iniciarToolEntidades()

    herramientaInternaEntidades = ninguna
    Call seleccionarEntidadBorrado(0)
    Call seleccionarEntidad(0, Nothing, 0)
    
End Sub

Public Sub seleccionarEntidad(IndexEntidad As Integer, accion As iAccionEditor, posicion As Byte)
    ReDim entidadesSeleccionadas(1 To 1, 1 To 1) As tListaEntidadesSeleccionadas
    ReDim entidadesSeleccionadas(1, 1).infoEntidades(1 To 1)
    
    
    With entidadesSeleccionadas(1, 1).infoEntidades(1)
            Set .accion = accion
            .IndexEntidad = IndexEntidad
            .posicion = posicion
    End With
End Sub

Public Sub seleccionarEntidadBorrado(posicion As Byte)
    ReDim entidadesSeleccionadasBorrado.infoEntidades(1 To 1)
        
    With entidadesSeleccionadasBorrado.infoEntidades(1)
            Set .accion = Nothing
            .IndexEntidad = 0
            .posicion = posicion
    End With
End Sub

Public Sub click_BorrarEntidad()
    Call seleccionarEntidadBorrado(0)
    
    herramientaInternaEntidades = eHerramientasEntidades.borrar
    Call ME_Tools.seleccionarTool(frmMain.cmdBorrarEntidad, Tools.tool_entidades)
End Sub

Public Sub click_InsertarEntidadEnPos()
    herramientaInternaEntidades = eHerramientasEntidades.insertar
    Call ME_Tools.seleccionarTool(Nothing, Tools.tool_entidades)
End Sub

Public Sub click_BorrarEntidadEnPos()
    herramientaInternaEntidades = eHerramientasEntidades.borrar
    Call ME_Tools.seleccionarTool(Nothing, Tools.tool_entidades)
End Sub

Public Sub click_InsertarEntidad()
    herramientaInternaEntidades = eHerramientasEntidades.insertar
    Call ME_Tools.seleccionarTool(frmMain.cmdInsertarEntidad, Tools.tool_entidades)
End Sub

'Objetos
Public Sub rotarHerramientaInternaEntidad(paraArriba As Boolean)

    
    If paraArriba Then
        herramientaInternaEntidades = herramientaInternaEntidades + 1
        If herramientaInternaEntidades > cantidadSubHerramientaEntidades Then herramientaInternaEntidades = 1
    Else
        herramientaInternaEntidades = herramientaInternaEntidades - 1
        If herramientaInternaEntidades < 1 Then herramientaInternaEntidades = cantidadSubHerramientaEntidades
    End If

    
    Call Me_Tools_Entidades.activarUltimaHerramienta

End Sub

Public Sub activarUltimaHerramienta()
    Select Case herramientaInternaEntidades
        Case eHerramientasEntidades.insertar
            Call Me_Tools_Entidades.click_InsertarEntidad
        Case eHerramientasEntidades.borrar
            Call Me_Tools_Entidades.click_BorrarEntidad
    End Select
End Sub

Public Sub cargarListaEntidades()

Dim loopEntidad As Integer

frmMain.lstEntidades.vaciar

For loopEntidad = LBound(EntidadesIndexadas) To UBound(EntidadesIndexadas)
    
    If Me_indexar_Entidades.existe(loopEntidad) Then
        With EntidadesIndexadas(loopEntidad)
            
            If .tipo = eTipoEntidadVida.puntos Then
                Call frmMain.lstEntidades.addString(loopEntidad, loopEntidad & " - " & .nombre)
            End If
        End With
    End If
Next

End Sub
