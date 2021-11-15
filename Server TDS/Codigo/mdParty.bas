Attribute VB_Name = "mdParty"
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

'Modulo de parties
'por EL OSO (ositear@yahoo.com.ar)

'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Public Const MAX_PARTIES = 300
'cantidad maxima de parties en el servidor

Public Const MINPARTYLEVEL = 15
'nivel minimo para crear party

Public Const PARTY_MAXMEMBERS = 5
'Cantidad maxima de gente en la party

Public Const PARTY_EXPERIENCIAPORGOLPE = False
'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)

Public Const MAXPARTYDELTALEVEL = 7
'maxima diferencia de niveles permitida en una party

Public Const MAXDISTANCIAINGRESOPARTY = 2
'distancia al leader para que este acepte el ingreso

Public Const PARTY_MAXDISTANCIA = 18
'maxima distancia a un exito para obtener su experiencia

'restan las muertes de los miembros?
Public Const CASTIGOS = False

Public Type tPartyMember
    UserIndex As Integer
    Experiencia As Long
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SOPORTES PARA LAS PARTIES
'(Ver este modulo como una clase abstracta "PartyManager")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextParty() As Integer
Dim I As Integer
NextParty = -1
For I = 1 To MAX_PARTIES
    If Parties(I) Is Nothing Then
        NextParty = I
        Exit Function
    End If
Next I
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
     PuedeCrearParty = True
    
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y3")
        PuedeCrearParty = False
    End If
       
    If UserList(UserIndex).Stats.ELV >= MINPARTYLEVEL Then
        If UserList(UserIndex).Stats.UserAtributos(Carisma) * UserList(UserIndex).Stats.UserSkills(Liderazgo) < 100 Then
        Call Senddata(ToIndex, UserIndex, 0, "|| Tu carisma y liderazgo no son suficientes para liderar una party." & FONTTYPE_PARTY)
        PuedeCrearParty = False
        End If
    Else
    Call Senddata(ToIndex, UserIndex, 0, "|| Tu nivel no es suficiente para liderar una party." & FONTTYPE_PARTY)
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
                Call Senddata(ToIndex, UserIndex, 0, "Y303")
                Exit Sub
            Else
                Set Parties(tInt) = New clsParty
                If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                    Call Senddata(ToIndex, UserIndex, 0, "Y304")
                    Set Parties(tInt) = Nothing
                    Exit Sub
                Else
                    Call Senddata(ToIndex, UserIndex, 0, "Y305")
                     Call Senddata(ToIndex, UserIndex, 0, "SS")
                    UserList(UserIndex).PartyIndex = tInt
                    UserList(UserIndex).PartySolicitud = 0
                    If Not Parties(tInt).HacerLeader(UserIndex) Then
                        Call Senddata(ToIndex, UserIndex, 0, "Y306")
                    Else
                        Call Senddata(ToIndex, UserIndex, 0, "Y307")
                    End If
                End If
            End If
        Else
            Call Senddata(ToIndex, UserIndex, 0, "Y308")
        End If
    Else
        Call Senddata(ToIndex, UserIndex, 0, "Y3")
    End If
Else
    Call Senddata(ToIndex, UserIndex, 0, "Y309")
End If
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
'ESTO ES enviado por el PJ para solicitar el ingreso a la party
Dim tInt As Integer
    If UserList(UserIndex).PartyIndex > 0 Then
        'si ya esta en una party
        Call Senddata(ToIndex, UserIndex, 0, "Y310")
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call Senddata(ToIndex, UserIndex, 0, "Y3")
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    tInt = UserList(UserIndex).flags.TargetUser
    If tInt > 0 Then
        If UserList(tInt).PartyIndex > 0 Then
            UserList(UserIndex).PartySolicitud = UserList(tInt).PartyIndex
            Call Senddata(ToIndex, tInt, 0, "PNI" + UserList(UserIndex).Name)
            Call Senddata(ToIndex, tInt, 0, "||El Personaje " & UserList(UserIndex).Name & " desea ingresar a la party." & FONTTYPE_PARTY)
            Call Senddata(ToIndex, UserIndex, 0, "Y311")
        Else
            Call Senddata(ToIndex, UserIndex, 0, "|| " & UserList(tInt).Name & " no es fundador de ninguna party." & FONTTYPE_PARTY)
            UserList(UserIndex).PartySolicitud = 0
            Exit Sub
        End If
    Else
        Call Senddata(ToIndex, UserIndex, 0, "Y312")
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
    Call Senddata(ToIndex, UserIndex, 0, "Y313")
End If
End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
Dim PI As Integer
Dim razon As String
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
            Call Senddata(ToIndex, leader, 0, "Y360")
        End If
    Else
        Call Senddata(ToIndex, leader, 0, "|| " & UserList(OldMember).Name & " no pertenece a tu party." & FONTTYPE_PARTY)
    End If
Else
    Call Senddata(ToIndex, leader, 0, "Y361")
End If
End Sub

Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
'el UI es el leader
Dim PI As Integer
Dim razon As String

PI = UserList(leader).PartyIndex
If PI > 0 Then
    If Parties(PI).EsPartyLeader(leader) Then
        If UserList(NewMember).PartyIndex = 0 Then
            If Not UserList(leader).flags.Muerto = 1 Then
                If Not UserList(NewMember).flags.Muerto = 1 Then
                    If UserList(NewMember).PartySolicitud = PI Then
                        If Parties(PI).PuedeEntrar(NewMember, razon) Then
                            If Parties(PI).NuevoMiembro(NewMember) Then
                                Call Senddata(ToIndex, NewMember, 0, "GH")
                                Call Senddata(ToIndex, NewMember, 0, "PNI" & UserList(NewMember).Name)
                                Call Parties(PI).MandarMensajeAConsola(UserList(leader).Name & " ha aceptado a " & UserList(NewMember).Name & " en la party." & FONTTYPE_PARTY, "")
                                UserList(NewMember).PartyIndex = PI
                                UserList(NewMember).PartySolicitud = 0
                            Else
                                'no pudo entrar
                                'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                                Call Senddata(ToAdmins, leader, 0, "|| Servidor> CATASTROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S " & FONTTYPE_PARTY)
                            End If
                        Else
                            'no debe entrar
                            Call Senddata(ToIndex, leader, 0, "|| " & razon & FONTTYPE_PARTY)
                        End If
                    Else
                        Call Senddata(ToIndex, leader, 0, "|| " & UserList(NewMember).Name & " no ha solicitado ingresar a tu party." & FONTTYPE_PARTY)
                        Exit Sub
                    End If
                Else
                    Call Senddata(ToIndex, leader, 0, "Y362")
                    Exit Sub
                End If
            Else
                Call Senddata(ToIndex, leader, 0, "Y362")
                Exit Sub
            End If
        Else
            Call Senddata(ToIndex, leader, 0, "||" & UserList(NewMember).Name & " ya es miembro de otra party." & FONTTYPE_PARTY)
            ' ya tiene party el otro tipo
        End If
    Else
        Call Senddata(ToIndex, leader, 0, "Y363")
        Exit Sub
    End If
Else
    Call Senddata(ToIndex, leader, 0, "Y361")
    Exit Sub
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
        Call Senddata(ToIndex, UserIndex, 0, "FC" & texto)
    End If
End Sub

Public Sub CParty(ByVal UserIndex As Integer)
Dim PI As Integer
Dim texto As String

    PI = UserList(UserIndex).PartyIndex
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(texto)
        Call Senddata(ToIndex, UserIndex, 0, "FC" & texto & "~0~0~0~0~0~")
    Else
    Call Senddata(ToIndex, UserIndex, 0, "FC")
    End If
End Sub

Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
Dim PI As Integer

If OldLeader = NewLeader Then Exit Sub
PI = UserList(OldLeader).PartyIndex
If PI > 0 Then
    If PI = UserList(NewLeader).PartyIndex Then
        If UserList(NewLeader).flags.Muerto = 0 Then
            If Parties(PI).EsPartyLeader(OldLeader) Then
                If Parties(PI).HacerLeader(NewLeader) Then
                    Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
                Else
                    Call Senddata(ToIndex, OldLeader, 0, "Y364")
                End If
            Else
                Call Senddata(ToIndex, OldLeader, 0, "Y365")
            End If
        Else
            Call Senddata(ToIndex, OldLeader, 0, "Y357")
        End If
    Else
        Call Senddata(ToIndex, OldLeader, 0, "||" & UserList(NewLeader).Name & " no pertenece a tu party." & FONTTYPE_PARTY)
    End If
End If
End Sub

Public Sub ActualizaExperiencias()
'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
Dim I As Integer

If Not PARTY_EXPERIENCIAPORGOLPE Then
    haciendoBK = True
    Call Senddata(ToAll, 0, 0, "BKW")
    For I = 1 To MAX_PARTIES
        If Not Parties(I) Is Nothing Then
            Call Parties(I).FlushExperiencia
        End If
    Next I
    Call Senddata(ToAll, 0, 0, "||Servidor> La experiencia fue distribuida." & FONTTYPE_SERVER)
    Call Senddata(ToAll, 0, 0, "BKW")
    haciendoBK = False
End If
End Sub

Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, mapa As Integer, X As Integer, Y As Integer)
    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
  
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, mapa, X, Y)
End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
CantMiembros = 0
If UserList(UserIndex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
End If
End Function
'********************Misery_Ezequiel 28/05/05********************'
