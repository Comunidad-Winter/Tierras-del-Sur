Attribute VB_Name = "modFianza"
Option Explicit


Private Const COSTO_FIANZA As Long = 100000


Public Sub pagarFianza(ByRef personaje As User, ByVal ejercito As String)
    Dim alineacion As eAlineaciones
    
    If personaje.flags.Muerto = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes estar vivo para poder pagar la fianza.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    If MapInfo(personaje.pos.map).Pk = True Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes estar en zona segura para poder pagar la fianza.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    alineacion = eAlineaciones.indefinido
    
    If ejercito = "INDIGO" Then
        alineacion = eAlineaciones.Real
    ElseIf ejercito = "ESCARLATA" Then
        alineacion = eAlineaciones.caos
    End If
        
    If alineacion = eAlineaciones.indefinido Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes elegir entre el ejército INDIGO y el ESCARLATA.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    If alineacion = personaje.faccion.alineacion Then
        EnviarPaquete Paquetes.mensajeinfo, "Ya perteneces al ejército seleccionado.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
        
    If personaje.Stats.GLD < COSTO_FIANZA Then
        EnviarPaquete Paquetes.mensajeinfo, "Para ingresar al ejército necesitas aportar " & FormatNumber(COSTO_FIANZA, 0, vbFalse, vbFalse, vbTrue) & " monedas de oro.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    personaje.faccion.alineacion = alineacion
    
    personaje.Stats.GLD = personaje.Stats.GLD - COSTO_FIANZA
    
    EnviarPaquete Paquetes.EnviarOro, Codify(personaje.Stats.GLD), personaje.UserIndex
    EnviarPaquete Paquetes.mensajeinfo, "¡Bienvenido!. Tu aporte será bien invertido por nuestros líderes.", personaje.UserIndex, ToIndex
    
    Call modPersonaje_TCP.actualizarNick(personaje)
End Sub
