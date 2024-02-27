VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CLASE PARTY
'por EL OSO (ositear@yahoo.com.ar)
'define los metodos de control de la entidad party.
'Las reglas del negocio de la parties se definen el el mdparty.bas

Option Explicit

Private p_members(1 To PARTY_MAXMEMBERS) As tPartyMember
'miembros

Private p_expTotal As Long
'Estadistica :D

Private p_Fundador As Integer

Private p_CantMiembros As Integer
'cantidad de miembros

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()
    p_expTotal = 0
    p_CantMiembros = 0
End Sub

'Destructor de clase
Public Sub Class_Terminate()

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MiExperiencia
' DateTime  : 18/02/2007 19:22
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function MiExperiencia(ByVal UserIndex As Integer) As Long
'Me dice cuanta experiencia tengo colectada ya en la party
Dim i As Integer

i = 0

While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex <> UserIndex
    i = i + 1
Wend
 
If i <= PARTY_MAXMEMBERS Then
    MiExperiencia = p_members(i).Experiencia
Else        'esto no deberia pasar :p
    MiExperiencia = -1
End If

End Function


Public Sub repartirOro(ByVal cantidad As Long)
'Se produjo un evento que da experiencia en la wp referenciada
Dim i As Integer
Dim UI As Integer
Dim expPerCapita As Double
Dim liderPos As WorldPos
Dim mapa As Integer

liderPos.map = UserList(p_members(1).UserIndex).pos.map

For i = 1 To PARTY_MAXMEMBERS
    UI = p_members(i).UserIndex
    If UI > 0 Then
        ' Siempre cerquitas del lider
        If liderPos.map = UserList(UI).pos.map And UserList(UI).flags.Muerto = 0 Then
            Call modUsuarios.agregarOro(UserList(UI), cantidad * p_members(i).Porcentaje)
        End If
    End If
Next i

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ObtenerExito
' DateTime  : 18/02/2007 19:22
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Friend Sub ObtenerExito(ByVal expGanada As Long, ByRef criatura As npc)
'Se produjo un evento que da experiencia en la wp referenciada
Dim i As Integer
Dim UI As Integer
Dim expPerCapita As Double
Dim liderPos As WorldPos
Dim mapa As Integer

If CantMiembros > 1 Then
    expGanada = expGanada * 1.1 ' La party da un bono de un 10%. #61.
End If

p_expTotal = p_expTotal + expGanada
liderPos.map = UserList(p_members(1).UserIndex).pos.map

For i = 1 To PARTY_MAXMEMBERS
    UI = p_members(i).UserIndex
    If UI > 0 Then
        ' Siempre cerquitas del lider
        If criatura.pos.map = UserList(UI).pos.map And liderPos.map = UserList(UI).pos.map And UserList(UI).flags.Muerto = 0 Then
            
            ' Distancia entre el usuario y la criatura
            If Distance(UserList(UI).pos.x, UserList(UI).pos.y, criatura.pos.x, criatura.pos.y) <= PARTY_MAXDISTANCIA Then
            
                'Regla: #30 Penalización si el nivel es menor
                If UserList(UI).Stats.ELV > criatura.Nivel Then
                    expPerCapita = (expGanada * p_members(i).Porcentaje) * PENALIZACION_CRIATURA_MENOR_NIVEL_USUARIO
                Else
                    expPerCapita = (expGanada * p_members(i).Porcentaje)
                End If
                  
                If expPerCapita > 0 Then
                    p_members(i).Experiencia = p_members(i).Experiencia + expPerCapita
                End If
                
                If PARTY_EXPERIENCIAPORGOLPE Then
                    Call modUsuarios.agregarExperiencia(UI, expPerCapita)
                    Call modPersonaje_TCP.actualizarExperiencia(UserList(UI))
                End If
            End If
        End If
    End If
Next i

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : DameUnPorc
' DateTime  : 18/02/2007 19:31
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DameUnPorc() As Single 'EL YIND
Dim userscp As Integer ' Cantidad de personajes con Porcentaje asignado.
Dim i As Integer

' Se fija si hay algún personaje que tenga porcentaje asignado
For i = 1 To UBound(p_members)
    If p_members(i).Porcentaje > 0 Then
        userscp = userscp + 1
    End If
Next i

Dim restos As Single

' Le voy a quitar un poco de experiencia a cada uno.
For i = 1 To UBound(p_members)
    If p_members(i).Porcentaje > 0 Then
        restos = FormatNumber(p_members(i).Porcentaje / (userscp + 1), 2)
        p_members(i).Porcentaje = p_members(i).Porcentaje - restos
        DameUnPorc = DameUnPorc + restos
    End If
Next i

If DameUnPorc = 0 Then
    DameUnPorc = 1
End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : DarPorc
' DateTime  : 18/02/2007 19:31
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub DarPorc(Porc As Single)  'EL YIND
Dim i As Integer
Dim userscp As Integer

For i = 1 To UBound(p_members)
    If p_members(i).UserIndex > 0 And p_members(i).Porcentaje > 0 Then
        userscp = userscp + 1
    End If
Next i
For i = 1 To UBound(p_members)
    If p_members(i).UserIndex > 0 And p_members(i).Porcentaje > 0 Then
        p_members(i).Porcentaje = p_members(i).Porcentaje + Porc / userscp
    End If
Next i

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Acomodar
' DateTime  : 18/02/2007 19:32
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Acomodar(cadena As String) 'EL YIND
    Dim ListaP As Variant
    Dim i As Integer
    Dim Actualizacion As String
    
    Actualizacion = "Porcentajes: "

    ListaP = Split(cadena, "|")
        For i = 1 To PARTY_MAXMEMBERS
            If p_members(i).UserIndex > 0 Then
                p_members(i).Porcentaje = val(ListaP(i - 1)) / 100
                Actualizacion = Actualizacion & UserList(p_members(i).UserIndex).Name & " " & (p_members(i).Porcentaje * 100) & "%"
            End If
        Next i
    Actualizacion = Actualizacion & "."
    Call MandarMensajeAConsola(Actualizacion, p_members(1).UserIndex)

End Sub
'---------------------------------------------------------------------------------------
' Procedure : MandarAcomodar
' DateTime  : 18/02/2007 19:22
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub MandarAcomodar(ByRef UserIndex As Integer) 'EL YIND
Dim i As Integer
Dim cadena As String

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
        cadena = cadena & UserList(p_members(i).UserIndex).Name & "|" & p_members(i).Porcentaje & "|"
    End If
Next i

cadena = cadena & ByteToString(UserList(UserIndex).Stats.UserSkills(Liderazgo))
EnviarPaquete 167, cadena, UserIndex, ToIndex

End Sub
'---------------------------------------------------------------------------------------
' Procedure : MandarMensajeAConsola
' DateTime  : 18/02/2007 19:22
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)

Dim i As Integer

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(39) & texto, p_members(i).UserIndex, ToIndex
    End If
Next i
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EsPartyLeader
' DateTime  : 18/02/2007 19:31
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EsPartyLeader(ByVal UserIndex As Integer) As Boolean
    EsPartyLeader = (UserIndex = p_Fundador)
End Function

'---------------------------------------------------------------------------------------
' Procedure : NuevoMiembro
' DateTime  : 18/02/2007 19:22
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function NuevoMiembro(ByVal UserIndex As Integer) As Boolean
Dim i As Integer

i = 1
While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex > 0
    i = i + 1
Wend

If i <= PARTY_MAXMEMBERS Then
    p_members(i).Experiencia = 0
    p_members(i).UserIndex = UserIndex
    p_members(i).Porcentaje = DameUnPorc
    NuevoMiembro = True
    p_CantMiembros = p_CantMiembros + 1
Else
    NuevoMiembro = False
End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : SaleMiembro
' DateTime  : 18/02/2007 19:21
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SaleMiembro(ByVal UserIndex As Integer) As Boolean
'el valor de retorno representa si se disuelve la party
Dim i As Integer
Dim j As Integer

i = 1
SaleMiembro = False
While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex <> UserIndex
    i = i + 1
Wend

If i = 1 Then
    'sale el founder, la party se disuelve
    SaleMiembro = True
    Call MandarMensajeAConsola("El lider disuelve la party.", "")
    For j = PARTY_MAXMEMBERS To 1 Step -1
        If p_members(j).UserIndex > 0 Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(37) & UserList(p_members(1).UserIndex).Name, p_members(j).UserIndex, ToIndex
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(38) & p_members(j).Experiencia, p_members(j).UserIndex, ToIndex
    
            If Not PARTY_EXPERIENCIAPORGOLPE Then
                Call modUsuarios.agregarExperiencia(p_members(j).UserIndex, p_members(j).Experiencia)
            End If
            
            EnviarPaquete Paquetes.ppe, "", p_members(j).UserIndex, ToIndex
            Call MandarMensajeAConsola(UserList(p_members(j).UserIndex).Name & " abandona la party.", "")
            UserList(p_members(j).UserIndex).PartyIndex = 0
            p_CantMiembros = p_CantMiembros - 1
            p_members(j).UserIndex = 0
            p_members(j).Experiencia = 0
            Call DarPorc(p_members(j).Porcentaje)
            p_members(j).Porcentaje = 0
        End If
    Next j
Else
    If i <= PARTY_MAXMEMBERS Then
        If Not PARTY_EXPERIENCIAPORGOLPE Then
            Call modUsuarios.agregarExperiencia(p_members(i).UserIndex, p_members(i).Experiencia)
        End If
        Call MandarMensajeAConsola(UserList(p_members(i).UserIndex).Name & " abandona la party.", "")
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(38) & p_members(i).Experiencia, p_members(i).UserIndex, ToIndex
        EnviarPaquete Paquetes.ppe, "", p_members(i).UserIndex, ToIndex
        EnviarPaquete Paquetes.Sefuedeparty, UserList(p_members(i).UserIndex).Name, p_members(1).UserIndex, ToIndex
        p_CantMiembros = p_CantMiembros - 1
        p_members(i).UserIndex = 0
        p_members(i).Experiencia = 0
           
        'Reasigno porcentajes
        Dim o As Integer

        
        For o = 1 To PARTY_MAXMEMBERS
        If p_members(o).UserIndex > 0 Then
        p_members(o).Porcentaje = p_members(o).Porcentaje + FormatNumber((p_members(i).Porcentaje / p_CantMiembros), 2)
        End If
        Next
        Dim suma As Single
        For o = 1 To PARTY_MAXMEMBERS
        If p_members(o).UserIndex > 0 Then
        suma = suma + p_members(o).Porcentaje
        End If
        Next
        If suma > 1 Then
        suma = FormatNumber(suma - 1, 2)
        p_members(1).Porcentaje = p_members(1).Porcentaje - suma
        End If
        p_members(i).Porcentaje = 0
        CompactMemberList
        Call OnlineParty(p_members(1).UserIndex)
        '''
    End If
End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : HacerLeader
' DateTime  : 18/02/2007 19:27
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function HacerLeader(ByVal UserIndex As Integer) As Boolean
Dim i As Integer
Dim OldLeader As Integer
Dim oldExp As Long
Dim oldPorc As Long
Dim UserIndexIndex As Integer
Dim maximaDiferenciaNiveles As Integer

UserIndexIndex = 0
HacerLeader = True

maximaDiferenciaNiveles = getMaximaDiferenciaNiveles(UserList(UserIndex).Stats.UserSkills(eSkills.Liderazgo))

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
        If p_members(i).UserIndex = UserIndex Then
            UserIndexIndex = i
            Exit For
        End If
        'chequeo la dif entre levels
        If Abs(UserList(UserIndex).Stats.ELV - UserList(p_members(i).UserIndex).Stats.ELV) > maximaDiferenciaNiveles Then
            HacerLeader = False
        End If
    End If
Next i

If Not HacerLeader Then Exit Function

If UserIndexIndex = 0 Then
    'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
    Call LogError("INCONSISTENCIA DE PARTIES")
  '  Call Senddata(ToAdmins, 0, 0, "Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!")
    HacerLeader = False
    Exit Function
End If


'aca esta todo bien y doy vuelta las collections
OldLeader = p_members(1).UserIndex
oldExp = p_members(1).Experiencia
oldPorc = p_members(1).Porcentaje


p_members(1).UserIndex = p_members(UserIndexIndex).UserIndex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia
p_members(1).Porcentaje = p_members(UserIndexIndex).Porcentaje

p_members(i).UserIndex = OldLeader
p_members(i).Experiencia = oldExp
p_members(i).Porcentaje = oldPorc

p_Fundador = p_members(1).UserIndex

End Function


'---------------------------------------------------------------------------------------
' Procedure : ObtenerMiembrosOnline
' DateTime  : 18/02/2007 19:21
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ObtenerMiembrosOnline(ByRef MemberList As String)
'esta funcion viola un poco la correcta forma de hacer la clase
'pero bue...
Dim i As Integer
Dim z As Byte

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
    z = z + 1
        MemberList = MemberList & p_members(i).Experiencia & ";" & p_members(i).Porcentaje * 100 & ";" & UserList(p_members(i).UserIndex).Name & ":"
    End If
Next i

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PuedeEntrar
' DateTime  : 18/02/2007 19:21
' Author    : Marce
' Purpose   : DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
'---------------------------------------------------------------------------------------
'
Public Function PuedeEntrar(ByVal UserIndex As Integer, ByRef razon As String) As Boolean

Dim esArmada As Boolean
Dim esCaos As Boolean
Dim MyLevel As Integer
Dim i As Integer
Dim UI As Integer
Dim puntosLiderazgo As Integer
Dim maximaDiferenciaNiveles As Byte

If distancia(UserList(p_members(1).UserIndex).pos, UserList(UserIndex).pos) > MAXDISTANCIAINGRESOPARTY Then
    razon = "El personaje se encuentra demasiado lejos tuyo."
    PuedeEntrar = False
    Exit Function
End If

If Not (p_members(PARTY_MAXMEMBERS).UserIndex = 0) Then
    razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS
    PuedeEntrar = False
    Exit Function
End If

esArmada = (UserList(UserIndex).faccion.ArmadaReal = 1)
esCaos = (UserList(UserIndex).faccion.FuerzasCaos = 1)
MyLevel = UserList(UserIndex).Stats.ELV
puntosLiderazgo = UserList(p_members(1).UserIndex).Stats.UserSkills(Liderazgo)

maximaDiferenciaNiveles = getMaximaDiferenciaNiveles(puntosLiderazgo)

PuedeEntrar = True

' Analizo, con respecto al resto de los integrantes, las reglas
For i = 1 To PARTY_MAXMEMBERS
    UI = p_members(i).UserIndex

    If UI > 0 Then
    
        ' Aspirante armada en party crimi
        If Not UserList(UI).faccion.alineacion = eAlineaciones.Neutro And Not UserList(UserIndex).faccion.alineacion = eAlineaciones.Neutro Then
            If Not UserList(UI).faccion.alineacion = UserList(UserIndex).faccion.alineacion Then
                razon = "No puedes compartir la party con integrantes del ejercito adversario."
                PuedeEntrar = False
            End If
        End If
                
        If UserList(UI).Stats.ELV <= 13 Then
            razon = "No pueden ingresar a la party personajes de nivel 13 o inferior."
            PuedeEntrar = False
        End If
        
        'diferencia entre niveles
        If (Abs(UserList(UI).Stats.ELV - UserList(UserIndex).Stats.ELV) > maximaDiferenciaNiveles) Then
            razon = "La diferencia con " & UserList(UI).Name & " es mayor a " & maximaDiferenciaNiveles & " niveles."
            PuedeEntrar = False
        End If
    
        If Not PuedeEntrar Then
            Exit For 'violate una programacion estructurada
        End If
    End If
Next i



End Function


'---------------------------------------------------------------------------------------
' Procedure : FlushExperiencia
' DateTime  : 18/02/2007 19:27
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub FlushExperiencia()
'esta funcion se invoca frente a cerradas del servidor. Flushea la experiencia
'acumulada a los usuarios.
Dim i As Integer

If PARTY_EXPERIENCIAPORGOLPE Then    'esto sirve SOLO cuando acumulamos la experiencia!
    Exit Sub
End If

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).UserIndex > 0 Then
        If p_members(i).Experiencia > 0 Then
            Call modUsuarios.agregarExperiencia(p_members(i).UserIndex, p_members(i).Experiencia)
            'Call AddtoVar(UserList(p_members(i).UserIndex).Stats.Exp, p_members(i).Experiencia, MAXEXP)
            'Call CheckUserLevel(UserList(p_members(i).UserIndex), p_members(i).UserIndex)
        'Else
            'TODO y esto!?!??! Que sentido tiene, ya se sabe que la experiencia es 0.
            'Y el usuario no deberia tener porque experiencia negativa...
            'If Abs(UserList(p_members(i).UserIndex).Stats.Exp) > Abs(p_members(i).Experiencia) Then
            '    UserList(p_members(i).UserIndex).Stats.Exp = UserList(p_members(i).UserIndex).Stats.Exp + p_members(i).Experiencia
            'Else
            '    UserList(p_members(i).UserIndex).Stats.Exp = 0
            'End If
        End If
        p_members(i).Experiencia = 0
    End If
Next i

End Sub




'---------------------------------------------------------------------------------------
' Procedure : CompactMemberList
' DateTime  : 18/02/2007 19:32
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub CompactMemberList()
Dim i As Integer
Dim freeIndex As Integer

i = 1
While i <= PARTY_MAXMEMBERS
    If p_members(i).UserIndex = 0 And freeIndex = 0 Then
        freeIndex = i
    ElseIf p_members(i).UserIndex > 0 And freeIndex > 0 Then
        p_members(freeIndex).Experiencia = p_members(i).Experiencia
        p_members(freeIndex).UserIndex = p_members(i).UserIndex
        p_members(freeIndex).Porcentaje = p_members(i).Porcentaje
        p_members(i).UserIndex = 0
        p_members(i).Experiencia = 0
        p_members(i).Porcentaje = 0
        'muevo el de la pos i a freeindex
        i = freeIndex
        freeIndex = 0
    End If
    i = i + 1
Wend

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CantMiembros
' DateTime  : 18/02/2007 19:32
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function CantMiembros() As Integer
CantMiembros = p_CantMiembros
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReasignoPorcentaje
' DateTime  : 18/02/2007 19:21
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Function ReasignoPorcentaje() As Single 'EL YIND
'Dim userscp As Integer
'Dim NewPorc As Single
'Dim i As Integer

'For i = 1 To UBound(p_members)
'    If p_members(i).userIndex > 0 Then
'        userscp = userscp + 1
'    End If
'Next i
'NewPorc = FormatNumber(1 / (userscp), 2)
'For i = 1 To UBound(p_members)
'    If p_members(i).userIndex > 0 Then
'        p_members(i).Porcentaje = p_members(i).Porcentaje - NewPorc / userscp
'    End If
'Next i
'DameUnPorc = NewPorc

'End Function
