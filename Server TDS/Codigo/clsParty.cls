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
'Argentum Online 0.11.20
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez




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
'el creador
'[Wizard 03/09/05]
Private P_Escuchando As String 'Defino a String por si acaso se buguea, nunca llegaria a enviar el mensaje a otro usuario.
'Quise decir, que luego buscaria el UI del nick, y no enviaria al Ui directamente porque si el otro cerrara y esto se bugueara otro personaje escucharia
'Seria msa facil hacerlo con un public, pero quise seguir el tipo del class
'[/Wizar]

Private p_CantMiembros As Integer
'cantidad de miembros

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()
    p_expTotal = 0
    p_CantMiembros = 0
    P_Escuchando = ""
End Sub

'Destructor de clase
Public Sub Class_Terminate()

End Sub


Public Function MiExperiencia(ByVal UserIndex As Integer) As Long
'Me dice cuanta experiencia tengo colectada ya en la party
Dim I As Integer
I = 0

While I <= PARTY_MAXMEMBERS And p_members(I).UserIndex <> UserIndex
    I = I + 1
Wend
 
If I <= PARTY_MAXMEMBERS Then
    MiExperiencia = p_members(I).Experiencia
Else        'esto no deberia pasar :p
    MiExperiencia = -1
End If

End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal mapa As Integer, X As Integer, Y As Integer)
'Se produjo un evento que da experiencia en la wp referenciada
Dim I As Integer
Dim UI As Integer
Dim expPerCapita As Double

p_expTotal = p_expTotal + ExpGanada

expPerCapita = CLng(ExpGanada / p_CantMiembros)

For I = 1 To PARTY_MAXMEMBERS
    UI = p_members(I).UserIndex
    If UI > 0 Then
        If mapa = UserList(UI).Pos.Map And UserList(UI).flags.Muerto = 0 Then
            If Distance(UserList(UI).Pos.X, UserList(UI).Pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
                p_members(I).Experiencia = p_members(I).Experiencia + expPerCapita
                If p_members(I).Experiencia < 0 Then
                    p_members(I).Experiencia = 0
                End If
                If PARTY_EXPERIENCIAPORGOLPE Then
                    Call AddtoVar(UserList(UI).Stats.Exp, expPerCapita, MAXEXP)
                    Call CheckUserLevel(UI)
                    Call SendUserStatsBox(UI)
                End If
            End If
        End If
    End If
Next I

End Sub

Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
Dim I As Integer
Dim espia As Integer

For I = 1 To PARTY_MAXMEMBERS
    If p_members(I).UserIndex > 0 Then
        Call Senddata(ToIndex, p_members(I).UserIndex, 0, "||" & texto)
    End If
Next I
    If P_Escuchando <> "" Then
        espia = NameIndex(P_Escuchando)
        Call Senddata(ToIndex, espia, 0, "||" & texto)
    End If
    
End Sub

Public Function EsPartyLeader(ByVal UserIndex As Integer) As Boolean
    EsPartyLeader = (UserIndex = p_Fundador)
End Function

Public Function NuevoMiembro(ByVal UserIndex As Integer) As Boolean
Dim I As Integer
I = 1
While I <= PARTY_MAXMEMBERS And p_members(I).UserIndex > 0
    I = I + 1
Wend

If I <= PARTY_MAXMEMBERS Then
    p_members(I).Experiencia = 0
    p_members(I).UserIndex = UserIndex
    NuevoMiembro = True
    p_CantMiembros = p_CantMiembros + 1
Else
    NuevoMiembro = False
End If

End Function

Public Function SaleMiembro(ByVal UserIndex As Integer) As Boolean
'el valor de retorno representa si se disuelve la party
Dim I As Integer
Dim j As Integer
I = 1
SaleMiembro = False
While I <= PARTY_MAXMEMBERS And p_members(I).UserIndex <> UserIndex
    I = I + 1
Wend

If I = 1 Then
    'sale el founder, la party se disuelve
    SaleMiembro = True
    Call MandarMensajeAConsola("El lider disuelve la party.", "")
    For j = PARTY_MAXMEMBERS To 1 Step -1
        If p_members(j).UserIndex > 0 Then
            Call Senddata(ToIndex, p_members(j).UserIndex, 0, "||Abandonas la party liderada por " & UserList(p_members(1).UserIndex).Name & FONTTYPE_PARTY)
    
            Call Senddata(ToIndex, p_members(j).UserIndex, 0, "||Durante la misma has conseguido " & p_members(j).Experiencia & " puntos de experiencia!" & FONTTYPE_PARTY)
    
            If Not PARTY_EXPERIENCIAPORGOLPE Then
                Call AddtoVar(UserList(p_members(j).UserIndex).Stats.Exp, p_members(j).Experiencia, MAXEXP)
                Call CheckUserLevel(p_members(j).UserIndex)
                Call SendUserStatsBox(p_members(j).UserIndex)
            End If
            Call MandarMensajeAConsola(UserList(p_members(j).UserIndex).Name & " abandona la party." & FONTTYPE_PARTY, "")
            UserList(p_members(j).UserIndex).PartyIndex = 0
            p_CantMiembros = p_CantMiembros - 1
            p_members(j).UserIndex = 0
            p_members(j).Experiencia = 0
        End If
    Next j
            '[Wizard 03/09/05] No hay mas coneccion entre el gm y la party
            If P_Escuchando <> "" Then
                UserList(NameIndex(P_Escuchando)).flags.EscuchaParty = 0
                P_Escuchando = ""
            End If '[/Wizard]
Else
    If I <= PARTY_MAXMEMBERS Then
        If Not PARTY_EXPERIENCIAPORGOLPE Then
            Call AddtoVar(UserList(p_members(I).UserIndex).Stats.Exp, p_members(I).Experiencia, MAXEXP)
            Call CheckUserLevel(p_members(I).UserIndex)
            Call SendUserStatsBox(p_members(I).UserIndex)
        End If
        Call MandarMensajeAConsola(UserList(p_members(I).UserIndex).Name & " abandona la party." & FONTTYPE_PARTY, "")
        Call Senddata(ToIndex, p_members(I).UserIndex, 0, "||Durante la misma has conseguido " & p_members(I).Experiencia & " puntos de experiencia!" & FONTTYPE_PARTY)
        p_CantMiembros = p_CantMiembros - 1
        p_members(I).UserIndex = 0
        p_members(I).Experiencia = 0
        CompactMemberList
    End If
End If

End Function

Public Function HacerLeader(ByVal UserIndex As Integer) As Boolean
Dim I As Integer
Dim OldLeader As Integer
Dim oldExp As Long
Dim UserIndexIndex As Integer

UserIndexIndex = 0
HacerLeader = True

For I = 1 To PARTY_MAXMEMBERS
    If p_members(I).UserIndex > 0 Then
        If p_members(I).UserIndex = UserIndex Then
            UserIndexIndex = I
            Exit For
        End If
        'chequeo la dif entre levels
        If Abs(UserList(UserIndex).Stats.ELV - UserList(p_members(I).UserIndex).Stats.ELV) > MAXPARTYDELTALEVEL Then
            Call Senddata(ToIndex, p_members(I).UserIndex, 0, "||No se puede hacer líder a " & UserList(UserIndex).Name & " pues la diferencia entre niveles de los miembros de la party sería mayor a " & MAXPARTYDELTALEVEL & FONTTYPE_PARTY)
            HacerLeader = False
        End If
    End If
Next I

If Not HacerLeader Then Exit Function

If UserIndexIndex = 0 Then
    'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
    Call LogError("INCONSISTENCIA DE PARTIES")
    Call Senddata(ToAdmins, 0, 0, "|| Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!")
    HacerLeader = False
    Exit Function
End If


'aca esta todo bien y doy vuelta las collections
OldLeader = p_members(1).UserIndex
oldExp = p_members(1).Experiencia

p_members(1).UserIndex = p_members(UserIndexIndex).UserIndex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia

p_members(I).UserIndex = OldLeader
p_members(I).Experiencia = oldExp

p_Fundador = p_members(1).UserIndex

'no need to compact
End Function


Public Sub ObtenerMiembrosOnline(ByRef MemberList As String)
'esta funcion viola un poco la correcta forma de hacer la clase
'pero bue...
Dim I As Integer

For I = 1 To PARTY_MAXMEMBERS
    If p_members(I).UserIndex > 0 Then
        MemberList = MemberList & UserList(p_members(I).UserIndex).Name & ": " & p_members(I).Experiencia & vbCrLf
    End If
Next I

MemberList = MemberList & vbCrLf & "Experiencia total: " & p_expTotal

End Sub


Public Function PuedeEntrar(ByVal UserIndex As Integer, ByRef razon As String) As Boolean
'DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
Dim esArmada As Boolean
Dim esCaos As Boolean
Dim MyLevel As Integer
Dim I As Integer
Dim rv As Boolean
Dim UI As Integer

rv = True
esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)
esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
MyLevel = UserList(UserIndex).Stats.ELV

rv = Distancia(UserList(p_members(1).UserIndex).Pos, UserList(UserIndex).Pos) <= MAXDISTANCIAINGRESOPARTY
If rv Then
    rv = (p_members(PARTY_MAXMEMBERS).UserIndex = 0)
    If rv Then
        For I = 1 To PARTY_MAXMEMBERS
            UI = p_members(I).UserIndex
            'pongo los casos que evitarian que pueda entrar
            'aspirante armada en party crimi
            If UI > 0 Then
                If esArmada And Criminal(UI) Then
                    razon = "Los miembros de la armada no entran a una party con criminales."
                    rv = False
                End If
                'aspirante caos en party ciuda
                If esCaos And Not Criminal(UI) Then
                    razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
                    rv = False
                End If
                'aspirante crimi en party armada
                If UserList(UI).Faccion.ArmadaReal = 1 And Criminal(UserIndex) Then
                    razon = "Los criminales no entran a partys con miembros de la armada real."
                    rv = False
                End If
                'aspirante ciuda en party caos
                If UserList(UI).Faccion.FuerzasCaos = 1 And Not Criminal(UserIndex) Then
                    razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
                    rv = False
                End If
                'diferencia entre niveles
                If Abs(UserList(UI).Stats.ELV - UserList(UserIndex).Stats.ELV) > MAXPARTYDELTALEVEL Then
                    razon = "La diferencia con " & UserList(UI).Name & " es mayor a " & MAXPARTYDELTALEVEL & " niveles."
                    rv = False
                End If
            
                If Not rv Then Exit For 'violate una programacion estructurada
            End If
        Next I
    Else
        razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS
    End If
Else
    razon = "Te encuentras muy lejos del fundador."
End If

PuedeEntrar = rv

End Function


Public Sub FlushExperiencia()
'esta funcion se invoca frente a cerradas del servidor. Flushea la experiencia
'acumulada a los usuarios.

Dim I As Integer
If Not PARTY_EXPERIENCIAPORGOLPE Then   'esto sirve SOLO cuando acumulamos la experiencia!
    For I = 1 To PARTY_MAXMEMBERS
        If p_members(I).UserIndex > 0 Then
            If p_members(I).Experiencia > 0 Then
                Call AddtoVar(UserList(p_members(I).UserIndex).Stats.Exp, p_members(I).Experiencia, MAXEXP)
                Call CheckUserLevel(p_members(I).UserIndex)
            Else
                If Abs(UserList(p_members(I).UserIndex).Stats.Exp) > Abs(p_members(I).Experiencia) Then
                    UserList(p_members(I).UserIndex).Stats.Exp = UserList(p_members(I).UserIndex).Stats.Exp + p_members(I).Experiencia
                Else
                    UserList(p_members(I).UserIndex).Stats.Exp = 0
                End If
            End If
            p_members(I).Experiencia = 0
            Call SendUserStatsBox(p_members(I).UserIndex)
        End If
    Next I
End If

End Sub




Private Sub CompactMemberList()
Dim I As Integer
Dim freeIndex As Integer
Dim tPlayer As tPartyMember
I = 1
While I <= PARTY_MAXMEMBERS
    If p_members(I).UserIndex = 0 And freeIndex = 0 Then
        freeIndex = I
    ElseIf p_members(I).UserIndex > 0 And freeIndex > 0 Then
        p_members(freeIndex).Experiencia = p_members(I).Experiencia
        p_members(freeIndex).UserIndex = p_members(I).UserIndex
        p_members(I).UserIndex = 0
        p_members(I).Experiencia = 0
        'muevo el de la pos i a freeindex
        I = freeIndex
        freeIndex = 0
    End If
    I = I + 1
Wend

End Sub

Public Function CantMiembros() As Integer
CantMiembros = p_CantMiembros
End Function
'[Wizard 03/09/05] Sub, que cambia el nombre del gm espiando
Public Sub Espiando(ByVal Name As String)
If Name <> "Noneone" Then
    P_Escuchando = Name
Else
    P_Escuchando = ""
End If

End Sub
Public Function GetSpy() '[Wizard 03/09/05] ReadOnly Function que dice el nombre del gme scuchadno
GetSpy = P_Escuchando
End Function
