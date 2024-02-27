Attribute VB_Name = "mdParty"
'Modulo de parties
'por EL OSO (ositear@yahoo.com.ar)

Option Explicit

Public Const MAX_PARTIES = 300
'cantidad maxima de parties en el servidor

Public Const MINPARTYLEVEL = 15
'nivel minimo para crear party

Public Const PARTY_MAXMEMBERS = 5
'Cantidad maxima de gente en la party

Public Const PARTY_EXPERIENCIAPORGOLPE = True
'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)

Public Const MAXDISTANCIAINGRESOPARTY = 2
'distancia al leader para que este acepte el ingreso

Public Const PARTY_MAXDISTANCIA = 18
'maxima distancia a un exito para obtener su experiencia

'restan las muertes de los miembros?
Public Const CASTIGOS = False

Public Type tPartyMember
    UserIndex As Integer
    Experiencia As Long
    Porcentaje As Single
End Type

' Partyes!
Private Parties() As clsParty

Public Sub iniciar()
    ReDim Parties(1 To MAX_PARTIES) As clsParty
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SOPORTES PARA LAS PARTIES
'(Ver este modulo como una clase abstracta "PartyManager")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextParty() As Integer
Dim i As Integer
NextParty = -1
For i = 1 To MAX_PARTIES
    If Parties(i) Is Nothing Then
        NextParty = i
        Exit Function
    End If
Next i
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
     PuedeCrearParty = True
    
    If UserList(UserIndex).flags.Muerto = 1 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(3), UserIndex, ToIndex
        PuedeCrearParty = False
    End If
       
    If UserList(UserIndex).Stats.ELV >= MINPARTYLEVEL Or UserList(UserIndex).Stats.UserSkills(Liderazgo) >= 90 Then
        If UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) * UserList(UserIndex).Stats.UserSkills(Liderazgo) < 100 Then
        EnviarPaquete Paquetes.mensajeinfo, "Tu carisma y liderazgo no son suficientes para liderar una party.", UserIndex, ToIndex
        PuedeCrearParty = False
        End If
    Else
    EnviarPaquete Paquetes.mensajeinfo, "Tu nivel no es suficiente para liderar una party.", UserIndex, ToIndex
    PuedeCrearParty = False
    End If
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
Dim tInt As Integer
If UserList(UserIndex).PartyIndex = 0 Then
    If UserList(UserIndex).flags.Muerto = 0 Then
        If UserList(UserIndex).Stats.UserSkills(Liderazgo) >= 0 Then
            tInt = mdParty.NextParty
            If tInt = -1 Then
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(48), UserIndex, ToIndex
                Exit Sub
            Else
                Set Parties(tInt) = New clsParty
                If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(49), UserIndex, ToIndex
                    Set Parties(tInt) = Nothing
                    Exit Sub
                Else
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(50), UserIndex, ToIndex
                    EnviarPaquete Paquetes.FundoParty, "", UserIndex, ToIndex
                    UserList(UserIndex).PartyIndex = tInt
                    UserList(UserIndex).PartySolicitud = 0
                    Call OnlineParty(UserIndex)
                    If Not Parties(tInt).HacerLeader(UserIndex) Then
                        EnviarPaquete Paquetes.MensajeSimple2, Chr$(51), UserIndex, ToIndex
                    Else
                        EnviarPaquete Paquetes.MensajeSimple2, Chr$(52), UserIndex, ToIndex
                    End If
                End If
            End If
        Else
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(53), UserIndex, ToIndex
        End If
    Else
        EnviarPaquete Paquetes.MensajeSimple, Chr$(3), UserIndex, ToIndex
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(54), UserIndex, ToIndex
End If
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
'ESTO ES enviado por el PJ para solicitar el ingreso a la party
Dim tInt As Integer
    If UserList(UserIndex).PartyIndex > 0 Then
        'si ya esta en una party
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(55), UserIndex, ToIndex
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto = 1 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(3), UserIndex, ToIndex
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    tInt = UserList(UserIndex).flags.TargetUser
    If tInt > 0 Then
        If UserList(tInt).PartyIndex > 0 Then
            UserList(UserIndex).PartySolicitud = UserList(tInt).PartyIndex
            EnviarPaquete Paquetes.PPI, UserList(UserIndex).Name, tInt, ToIndex
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(1) & UserList(UserIndex).Name, tInt, ToIndex
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(56), UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.mensajeinfo, UserList(tInt).Name & " no es fundador de ninguna party.", UserIndex, ToIndex
            UserList(UserIndex).PartySolicitud = 0
            Exit Sub
        End If
    Else
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(57), UserIndex, ToIndex
        UserList(UserIndex).PartySolicitud = 0
    End If
End Sub

Public Sub SalirDeParty(ByVal UserIndex As Integer)
Dim PI As Integer
PI = UserList(UserIndex).PartyIndex
If PI > 0 Then
    If Parties(PI).SaleMiembro(UserIndex) Then
        'sale el leader
        Call Parties(PI).MandarMensajeAConsola(UserList(UserIndex).Name & "abandono la party", "")
        Set Parties(PI) = Nothing
    Else
        UserList(UserIndex).PartyIndex = 0
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(58), UserIndex, ToIndex
End If
End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
Dim PI As Integer

PI = UserList(leader).PartyIndex
If PI > 0 Then
    If PI = UserList(OldMember).PartyIndex Then
        If Parties(PI).EsPartyLeader(leader) Then
            If Parties(PI).SaleMiembro(OldMember) Then
                Call Parties(PI).MandarMensajeAConsola("pn", "")
                Set Parties(PI) = Nothing
            Else
                UserList(OldMember).PartyIndex = 0
            End If
        Else
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(105), leader, ToIndex
        End If
    Else
        EnviarPaquete Paquetes.mensajeinfo, UserList(OldMember).Name & " no pertenece a tu party.", leader, ToIndex
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(106), leader, ToIndex
End If
End Sub

Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
'el UI es el leader
Dim PI As Integer
Dim razon As String

PI = UserList(leader).PartyIndex

' ¿Tiene party¨?
If PI = 0 Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(106), leader, ToIndex
    Exit Sub
End If

If Parties(PI).EsPartyLeader(leader) = False Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(108), leader, ToIndex
    Exit Sub
End If

If UserList(NewMember).PartyIndex > 0 Then
    EnviarPaquete Paquetes.mensajeinfo, UserList(NewMember).Name & " ya es miembro de otra party.", leader, ToIndex
    Exit Sub
End If

If UserList(leader).flags.Muerto = 1 Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(107), leader, ToIndex
    Exit Sub
End If

If UserList(NewMember).flags.Muerto = 1 Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(107), leader, ToIndex
    Exit Sub
End If

If Not UserList(NewMember).PartySolicitud = PI Then
    EnviarPaquete Paquetes.mensajeinfo, UserList(NewMember).Name & " no ha solicitado ingresar a tu party.", leader, ToIndex
    Exit Sub
End If

If Not Parties(PI).PuedeEntrar(NewMember, razon) Then
    EnviarPaquete Paquetes.MensajeServer, razon, leader
    Exit Sub
End If

If Parties(PI).NuevoMiembro(NewMember) Then
    EnviarPaquete Paquetes.Integranteparty, "", NewMember, ToIndex
    EnviarPaquete Paquetes.Pni, UserList(NewMember).Name, leader, ToIndex
    Call OnlineParty(leader)
    Call Parties(PI).MandarMensajeAConsola(UserList(leader).Name & " ha aceptado a " & UserList(NewMember).Name & " en la party.", "")
    UserList(NewMember).PartyIndex = PI
    UserList(NewMember).PartySolicitud = 0
Else
    'no pudo entrar
    'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
End If
       
End Sub

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
Dim PI As Integer
    PI = UserList(UserIndex).PartyIndex
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)
    End If
End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)
Dim PI As Integer
Dim texto As String

    PI = UserList(UserIndex).PartyIndex
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(texto)
        EnviarPaquete Paquetes.OnParty, texto, UserIndex, ToIndex, 0
    End If
End Sub

Public Sub CParty(ByVal UserIndex As Integer)
Dim PI As Integer
Dim texto As String

    PI = UserList(UserIndex).PartyIndex
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(texto)
        EnviarPaquete Paquetes.OnParty, texto, UserIndex, ToIndex
    End If
End Sub

Public Sub ActualizaExperiencias()
'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
Dim i As Integer

If Not PARTY_EXPERIENCIAPORGOLPE Then
    haciendoBK = True
    EnviarPaquete Paquetes.Pausa, "", 0, ToAll
    For i = 1 To MAX_PARTIES
        If Not Parties(i) Is Nothing Then
            Call Parties(i).FlushExperiencia
        End If
    Next i
    EnviarPaquete Paquetes.MensajeServer, "La experiencia fue distribuida.", 0, ToAll, 0
    EnviarPaquete Paquetes.Pausa, "", 0, ToAll
    haciendoBK = False
End If
End Sub

Public Sub entregarOro(ByVal PartyIndex As Integer, ByVal cantidad As Long)
    Call Parties(PartyIndex).repartirOro(cantidad)
End Sub

Public Sub ObtenerExito(ByVal UserIndex As Integer, ByRef criatura As npc, ByVal Exp As Long)

    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
  
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, criatura)
End Sub
Public Sub Acomodar(ByVal UserIndex As Integer) 'EL YIND
Dim PI As Integer
PI = UserList(UserIndex).PartyIndex
If PI > 0 Then
    If Parties(PI).EsPartyLeader(UserIndex) Then
        Call Parties(PI).MandarAcomodar(UserIndex)
    Else
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(58), UserIndex, ToIndex
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(258), UserIndex, ToIndex
End If
End Sub

Public Function getMaximaDiferenciaNiveles(puntosLiderazgo As Integer) As Integer
    If puntosLiderazgo >= 90 Then
        getMaximaDiferenciaNiveles = 9
    ElseIf puntosLiderazgo >= 80 Then
        getMaximaDiferenciaNiveles = 6
    ElseIf puntosLiderazgo >= 75 Then
        getMaximaDiferenciaNiveles = 5
    Else
        getMaximaDiferenciaNiveles = 5
    End If
    
End Function


Private Sub MaxPermitido(UserIndex As Integer, ByRef Minp As Byte, ByRef Maxp As Byte)
Dim SkillsL As Integer

SkillsL = UserList(UserIndex).Stats.UserSkills(Liderazgo)
    
If SkillsL >= 90 Then
    Minp = 40
    Maxp = 60
ElseIf SkillsL >= 75 Then
    Minp = 45
    Maxp = 55
Else
    Minp = 100
    Maxp = 100
End If

End Sub
Public Sub AcomodarP(ByVal UserIndex As Integer, ByVal cadena As String)    'EL YIND
Dim PI As Integer
Dim ListaP As Variant
Dim suma As Integer
Dim i As Integer
Dim MaxPorcentajePermitido As Byte
Dim MinPorcentajePermitido As Byte

PI = UserList(UserIndex).PartyIndex

If PI > 0 Then
    If Parties(PI).EsPartyLeader(UserIndex) Then
    
        Call MaxPermitido(UserIndex, MinPorcentajePermitido, MaxPorcentajePermitido)
        
        ListaP = Split(cadena, "|")
        
         For i = 0 To UBound(ListaP) - 1
            If val(ListaP(i)) > MaxPorcentajePermitido Or val(ListaP(i)) < MinPorcentajePermitido Then
                EnviarPaquete Paquetes.mensajeinfo, "No tiene los suficientes puntos en liderazgo para acomodar de esta forma los porcentajes", UserIndex, ToIndex
                Exit Sub
            Else
                suma = suma + val(ListaP(i))
            End If
        Next i
        
        If suma <> 100 Then Exit Sub
        
        Call Parties(PI).Acomodar(cadena)
        
        Call OnlineParty(UserIndex)
    Else
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(58), UserIndex, ToIndex
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(58), UserIndex, ToIndex
End If
End Sub
Public Function CantMiembros(ByVal UserIndex As Integer) As Integer

CantMiembros = 0

If UserList(UserIndex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
End If

End Function
