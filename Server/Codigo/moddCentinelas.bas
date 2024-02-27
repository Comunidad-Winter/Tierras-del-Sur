Attribute VB_Name = "modCentinelas"
Option Explicit

Private Const LONGITUD_CODIGO = 4 'Longitud del codigo que se genera
Private Const TIEMPO_MAX_CENTINELA = 240 'Segundos. Tiempo que tiene para meter el codigo
Private Const MENSAJE_ALERTA_INTERVALO = 20 'Segundos. Tiempo entre cada mensaje de "Debe ingres..."
Private Const NPC_CENTINELA = 117 'Numero del NPC que representa al centinela.

Public Type tCentinela
    TiempoDes As Integer 'Tiempo que falta para ser eliminado
    UserID As Long 'ID del personaje al cual esta asignado el centinela
    npcIndex As Integer 'Criatura que representa al centinela
    codigo As String * LONGITUD_CODIGO 'Codigo del centinela
End Type

Public TiempoMin As Integer
Public Centinelas() As tCentinela
Public CentinelasTrabajando As Integer


'Determinar en que momento llamar a los centinelas, haciendo esta accion
Public Sub AntiMacrosL()

Dim codigo As String
Dim NpcPosN As WorldPos
Dim UserIndex As Integer
Dim loopX As Byte
Dim CentinelaIndex As Integer

Static ProxCent As Integer

'Los centinelas aparecen cada entre 8 y 25 minutos
' Y luego entre 15 y 30
If ProxCent = 0 Then ProxCent = RandomNumber(8, 25)

TiempoMin = TiempoMin + 1

If TiempoMin >= ProxCent Then

    'Momento de llamarlos
    TiempoMin = 0
    ProxCent = RandomNumber(15, 30)
    
    'No hay centinelas trabajando
    CentinelasTrabajando = 0
    
    'Recorro todos los trabajadores asignadoles un centinela
    TrabajadoresGroup.itIniciar
    
    Do While (TrabajadoresGroup.ithasNext)
        
        UserIndex = TrabajadoresGroup.itnext
        
        With UserList(UserIndex)
        
            If .flags.Trabajando Then
            
                CentinelasTrabajando = CentinelasTrabajando + 1
                
                ReDim Preserve Centinelas(CentinelasTrabajando)

                'Genero el codigo
                codigo = ""
                
                For loopX = 1 To LONGITUD_CODIGO
                    codigo = codigo & Chr$(RandomNumber(Asc("a"), Asc("z")))
                Next loopX
                
                'Asigno el codigo al centinela
                Centinelas(CentinelasTrabajando).codigo = codigo
               
               'Relaciono al centinela con el usuario
                Centinelas(CentinelasTrabajando).UserID = .id
                UserList(UserIndex).CentinelaID = CentinelasTrabajando
                
                'Donde va a respawenaer
                NpcPosN = .pos
                NpcPosN.y = NpcPosN.y - 1
                
                'Creo al centinela
                CentinelaIndex = SpawnNpc(NPC_CENTINELA, NpcPosN, True, False)
                
                'Se pudo crear?
                If CentinelaIndex > MAXNPCS Then
                    CentinelaIndex = 0
                    Call LogError("No se pudo crear un centinela en la posicion: (" & NpcPosN.map & " ; " & NpcPosN.x & " ; " & NpcPosN.y)
                Else
                    'FOR DEBUG
                    Call LogCentinela("Centinela " & CentinelaIndex & " en " & NpcList(CentinelaIndex).pos.map & " x " & NpcList(CentinelaIndex).pos.x & " y " & NpcList(CentinelaIndex).pos.y)
                End If

                Centinelas(CentinelasTrabajando).npcIndex = CentinelaIndex

                ' No permito que el usuario se mueva
                EnviarPaquete Paquetes.EnCentinelaPa, "", UserIndex, ToIndex
        
                ' Logueamos
                Call LogCentinelaMysql(UserList(UserIndex).id, codigo, "LLEGADA_CENTINELA")
                Call LogCentinela(" CENTINELA --> A " & UserList(UserIndex).Name & " Codigo " & codigo & " NPC " & CentinelaIndex)
            End If
        End With

    Loop
    
    ' Evitamos prender el timer si no hay nadie trabajando es raro pero bueno
    If CentinelasTrabajando > 0 Then
        frmMain.AntiMacrosCen.Enabled = True
    End If
End If
End Sub

Public Sub procesarCentinelas()

Dim loopC As Integer
Dim UserIndex As Integer
Dim UsuarioDesconectado As Boolean

Static ConteoCen As Integer

For loopC = 1 To UBound(Centinelas)

    'Cuando se va el centinela despues de haber puesto el codigo
    If Centinelas(loopC).TiempoDes > 0 Then
        Centinelas(loopC).TiempoDes = Centinelas(loopC).TiempoDes - 1
        If Centinelas(loopC).TiempoDes = 0 Then
            'Logueo la muerte natural
            Call LogCentinela("El usuario puso el codigo y muere el npc " & Centinelas(loopC).npcIndex)
            Call modCentinelas.eliminarCentinela(Centinelas(loopC))
        End If
    End If
    
    'En base al nombre obtiene si esta logueado o no
    If Centinelas(loopC).UserID > 0 Then
        UserIndex = IDIndex(Centinelas(loopC).UserID) 'Busco al personaje, entre los online, por su id
        
        UsuarioDesconectado = False
        
        If UserIndex = 0 Then
            UsuarioDesconectado = True
        Else
            ' Si el personaje no está relacionado al Socket.
            If UserList(UserIndex).ConnID = INVALID_SOCKET Then UsuarioDesconectado = True
        End If
        
        If Not UsuarioDesconectado Then
        
            If ConteoCen = TIEMPO_MAX_CENTINELA Then
            
                ' UserList(userIndex).flags.Ban = 1
                ' UserList(userIndex).flags.Banrazon = UserList(userIndex).flags.Banrazon & vbCrLf & "Uso de macro inasistido. Baneado por el centinela. " & Date
                ' UserList(userIndex).flags.Unban = "NUNCA"

                'Aviso y logueo
                ' EnviarPaquete Paquetes.MensajeServer, "Centinela ha baneado a " & UserList(userIndex).Name & ".", userIndex, ToAdmins
                Call LogCentinela("DESLOGUEO " & UserList(UserIndex).Name)
                Call LogCentinelaMysql(UserList(UserIndex).id, " ", "DESLOGUEADO")
                 
                'Mato al centinela
                Call modCentinelas.eliminarCentinela(Centinelas(loopC))
                
                'Cierro al personaje
                If Not CloseSocket(UserIndex) Then Call LogError("Procesar centinleas")
                
            ElseIf (ConteoCen Mod MENSAJE_ALERTA_INTERVALO) = 0 Or ConteoCen = 0 Then 'Voy mostrando cada X tiempo
                EnviarPaquete Paquetes.MensajeCompuesto, Chr$(2) & Centinelas(loopC).codigo, UserIndex
                'Si el centinela tiene un npc asignado hago que este hable
                If Centinelas(loopC).npcIndex > 0 Then
                    EnviarPaquete Paquetes.DescNpc2, ITS(str(NpcList(Centinelas(loopC).npcIndex).Char.charIndex)) & "Hola " & UserList(UserIndex).Name & ", yo soy el centinela anti-macros por favor tipea el comando '/CENTINELA " & Centinelas(loopC).codigo & "' antes de " & (TIEMPO_MAX_CENTINELA - ConteoCen) & " segundos.", UserIndex
                End If
            End If
        
        Else
            'El personaje deslogueo, mato al centinela
            Call LogCentinela("El personaje del centinela " & Centinelas(loopC).npcIndex & " cerro.")
            Call LogCentinelaMysql(Centinelas(loopC).UserID, " ", "USUARIO_CERRO")
            Call modCentinelas.eliminarCentinela(Centinelas(loopC))
        End If
        
    End If
Next loopC


'Termino el trabajo de los centinelas
If ConteoCen = TIEMPO_MAX_CENTINELA + 20 Then

    ReDim Centinelas(0)
    
    ConteoCen = 0
    frmMain.AntiMacrosCen.Enabled = False
    CentinelasTrabajando = 0
    
Else
    ConteoCen = ConteoCen + 1
End If

End Sub

Public Sub desactivarCentinela(ByRef centinela As tCentinela)
    centinela.TiempoDes = 10
    centinela.UserID = 0
    centinela.codigo = ""
End Sub

Public Sub eliminarCentinela(ByRef centinela As tCentinela)

centinela.TiempoDes = 0
centinela.UserID = 0
centinela.codigo = ""

If centinela.npcIndex > 0 Then
    LogCentinela ("Mato al Centinela NPC Index " & centinela.npcIndex)
    
    ' Quietamos el NPC
    Call QuitarNPC(centinela.npcIndex)

    ' Sacamos referencia
    centinela.npcIndex = 0
End If

End Sub


Public Sub ponerCodigo(Usuario As User, ByVal codigo As String)
    'Pone el codigo del centinela
    Dim TempInt As Integer
    Dim tempstr As String
    
    If Usuario.CentinelaID = 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "¡Tu no estas marcado por una centinela!", Usuario.UserIndex, ToIndex
        Exit Sub
    End If
            
    'Obtengo el codigo
    TempInt = Usuario.CentinelaID
    tempstr = Centinelas(TempInt).codigo
    
    codigo = Trim$(LCase$(codigo))
    
    'Puso e codigo correctamente?
     If tempstr = codigo Then
            
        LogCentinela "el personaje " & Usuario.Name & " puso el codigo correctamente"
        Call LogCentinelaMysql(Usuario.id, codigo, "INGRESO_CORRECTO")
        EnviarPaquete Paquetes.mensajeinfo, "¡Código ingresado exitosamente!", Usuario.UserIndex, ToIndex
            
        'Hago que hable
        If Centinelas(TempInt).npcIndex > 0 Then
            EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(Centinelas(TempInt).npcIndex).Char.charIndex) & "Gracias, espero no haberlo molestado.", Usuario.UserIndex, ToIndex, Usuario.pos.map
        End If
                
        EnviarPaquete Paquetes.EnCentinelaPa, "", Usuario.UserIndex, ToIndex
                        
        Call desactivarCentinela(Centinelas(TempInt))

        'Desrelacion el centinela con este usuario.
        Usuario.CentinelaID = 0
                
    Else
        EnviarPaquete Paquetes.mensajeinfo, "¡El código ingresado NO ES el correcto!", Usuario.UserIndex, ToIndex
        Call LogCentinelaMysql(Usuario.id, Left(codigo, 5), "INGRESO_INCORRECTO")
        LogCentinela "el personaje " & Usuario.Name & " puso el codigo inc-correctamente. Puso " & codigo & " y es " & Centinelas(TempInt).codigo
    End If
End Sub
