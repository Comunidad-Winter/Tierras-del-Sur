Attribute VB_Name = "modClanes"
'Argentum Online 0.9.0.2
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

'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Public Guilds As New Collection

Public Sub ComputeVote(ByVal UserIndex As Integer, ByVal rdata As String)
Dim myGuild As cGuild

Set myGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If myGuild Is Nothing Then Exit Sub
If Not myGuild.Elections Then
   Call Senddata(ToIndex, UserIndex, 0, "Y13")
   Exit Sub
End If
If UserList(UserIndex).GuildInfo.YaVoto = 1 Then
   Call Senddata(ToIndex, UserIndex, 0, "Y14")
   Exit Sub
End If
If Not myGuild.IsMember(rdata) Then
   Call Senddata(ToIndex, UserIndex, 0, "Y15")
   Exit Sub
End If
Call myGuild.Votes.Add(rdata)
UserList(UserIndex).GuildInfo.YaVoto = 1
Call Senddata(ToIndex, UserIndex, 0, "Y16")
End Sub

Public Sub ResetUserVotes(ByRef myGuild As cGuild)
On Error GoTo errh
Dim k As Integer, Index As Integer
Dim UserFile As String

For k = 1 To myGuild.Members.Count
    Index = DameUserIndexConNombre(myGuild.Members(k))
    If Index <> 0 Then 'is online
        UserList(Index).GuildInfo.YaVoto = 0
    Else
        UserFile = CharPath & UCase$(myGuild.Members(k)) & ".chr"
        If FileExist(UserFile, vbNormal) Then
                Call WriteVar(UserFile, "GUILD", "YaVoto", 0)
        End If
    End If
Next k
errh:
End Sub

Public Sub DayElapsed()
On Error GoTo errh
Dim t%
Dim MemberIndex As Integer
Dim UserFile As String

For t% = 1 To Guilds.Count
    If Guilds(t%).DaysSinceLastElection < Guilds(t%).ElectionPeriod Then
        Guilds(t%).DaysSinceLastElection = Guilds(t%).DaysSinceLastElection + 1
    Else
       If Guilds(t%).Elections = False Then
            Guilds(t%).ResetVotes
            Call ResetUserVotes(Guilds(t%))
            Guilds(t%).Elections = True
            MemberIndex = DameGuildMemberIndex(Guilds(t%).GuildName)
            If MemberIndex <> 0 Then
                Call Senddata(ToGuildMembers, MemberIndex, 0, "Y201")
                Call Senddata(ToGuildMembers, MemberIndex, 0, "Y202")
                Call Senddata(ToGuildMembers, MemberIndex, 0, "Y203")
                Call Senddata(ToGuildMembers, MemberIndex, 0, "Y204")
            End If
        Else
            If Guilds(t%).Members.Count > 1 Then
                    'compute elections results
                    Dim leader$, newleaderindex As Integer, oldleaderindex As Integer
                    leader$ = Guilds(t%).NuevoLider
                    Guilds(t%).Elections = False
                    MemberIndex = DameGuildMemberIndex(Guilds(t%).GuildName)
                    newleaderindex = DameUserIndexConNombre(leader$)
                    oldleaderindex = DameUserIndexConNombre(Guilds(t%).leader)
                    If UCase$(leader$) <> UCase$(Guilds(t%).leader) Then
                        If oldleaderindex <> 0 Then
                            UserList(oldleaderindex).GuildInfo.EsGuildLeader = 0
                        Else
                            UserFile = CharPath & UCase$(Guilds(t%).leader) & ".chr"
                            If FileExist(UserFile, vbNormal) Then
                                    Call WriteVar(UserFile, "GUILD", "EsGuildLeader", 0)
                            End If
                        End If
                        If newleaderindex <> 0 Then
                            UserList(newleaderindex).GuildInfo.EsGuildLeader = 1
                            Call AddtoVar(UserList(newleaderindex).GuildInfo.VecesFueGuildLeader, 1, 10000)
                        Else
                            UserFile = CharPath & UCase$(leader$) & ".chr"
                            If FileExist(UserFile, vbNormal) Then
                                    Call WriteVar(UserFile, "GUILD", "EsGuildLeader", 1)
                            End If
                        End If
                        Guilds(t%).leader = leader$
                    End If
                    If MemberIndex <> 0 Then
                            Call Senddata(ToGuildMembers, MemberIndex, 0, "Y18")
                            Call Senddata(ToGuildMembers, MemberIndex, 0, "||El nuevo lider es " & leader$ & FONTTYPE_GUILD)
                    End If
                    If newleaderindex <> 0 Then
                        Call Senddata(ToIndex, newleaderindex, 0, "Y19")
                        Call GiveGuildPoints(400, newleaderindex)
                    End If
                    Guilds(t%).DaysSinceLastElection = 0
            End If
        End If
    End If
Next t%
Exit Sub
errh:
    Call LogError(Err.Description & " error en DayElapsed.")
End Sub

Public Sub GiveGuildPoints(ByVal Pts As Integer, ByVal UserIndex As Integer, Optional ByVal SendNotice As Boolean = True)
If SendNotice Then _
   Call Senddata(ToIndex, UserIndex, 0, "||¡¡¡Has recibido " & Pts & " guildpoints!!!" & FONTTYPE_GUILD)
Call AddtoVar(UserList(UserIndex).GuildInfo.GuildPoints, Pts, 9000000)
End Sub

Public Sub DropGuildPoints(ByVal Pts As Integer, ByVal UserIndex As Integer, Optional ByVal SendNotice As Boolean = True)
UserList(UserIndex).GuildInfo.GuildPoints = UserList(UserIndex).GuildInfo.GuildPoints - Pts
End Sub

Public Sub AcceptPeaceOffer(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim oGuild As cGuild

Set oGuild = FetchGuild(rdata)
If oGuild Is Nothing Then Exit Sub
If Not oGuild.IsEnemy(UserList(UserIndex).GuildInfo.GuildName) Then
    Call Senddata(ToIndex, UserIndex, 0, "Y184" & FONTTYPE_GUILD)
    Exit Sub
End If
Call oGuild.RemoveEnemy(UserList(UserIndex).GuildInfo.GuildName)
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub
Call oGuild.RemoveEnemy(rdata)
Call oGuild.RemoveProposition(rdata)

Dim MemberIndex As Integer
MemberIndex = DameUserIndexConNombre(rdata)

If MemberIndex <> 0 Then _
    Call Senddata(ToGuildMembers, MemberIndex, 0, "||El clan firmó la paz con " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_GUILD)
Call Senddata(ToGuildMembers, UserIndex, 0, "||El clan firmó la paz con " & rdata & FONTTYPE_GUILD)
End Sub

Public Sub SendPeaceRequest(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub
Dim Soli As cSolicitud

Set Soli = oGuild.GetPeaceRequest(rdata)
If Soli Is Nothing Then Exit Sub
Call Senddata(ToIndex, UserIndex, 0, "PEACEDE" & Soli.Desc)
End Sub


Public Sub RecievePeaceOffer(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim H$
H$ = UCase$(ReadField(1, rdata, 44))

If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(H$) Then Exit Sub
Dim oGuild As cGuild

Set oGuild = FetchGuild(H$)
If oGuild Is Nothing Then Exit Sub
If Not oGuild.IsEnemy(UserList(UserIndex).GuildInfo.GuildName) Then
    Call Senddata(ToIndex, UserIndex, 0, "Y184")
    Exit Sub
End If
If oGuild.IsAllie(UserList(UserIndex).GuildInfo.GuildName) Then
    Call Senddata(ToIndex, UserIndex, 0, "Y185")
    Exit Sub
End If
Dim peaceoffer As New cSolicitud

peaceoffer.Desc = ReadField(2, rdata, 44)
peaceoffer.UserName = UserList(UserIndex).GuildInfo.GuildName

If Not oGuild.IncludesPeaceOffer(peaceoffer.UserName) Then
    Call oGuild.PeacePropositions.Add(peaceoffer)
    Call Senddata(ToIndex, UserIndex, 0, "Y186")
Else
    Call Senddata(ToIndex, UserIndex, 0, "Y187")
End If
End Sub

Public Sub SendPeacePropositions(ByVal UserIndex As Integer)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub
Dim L%, k$

If oGuild.PeacePropositions.Count = 0 Then Exit Sub
k$ = "PEACEPR" & oGuild.PeacePropositions.Count & ","

For L% = 1 To oGuild.PeacePropositions.Count
    k$ = k$ & oGuild.PeacePropositions(L%).UserName & ","
Next L%
Call Senddata(ToIndex, UserIndex, 0, k$)
End Sub

Public Sub EacharMember(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo hayerror
Dim NameeChado As String
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
    Exit Sub
ElseIf (UCase(UserList(UserIndex).GuildRef.leader) <> UCase(UserList(UserIndex).Name)) And (UCase(rdata) <> UCase(UserList(UserIndex).Name)) Then
    Exit Sub
End If
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub

Dim MemberIndex As Integer
MemberIndex = DameUserIndexConNombre(rdata)

If MemberIndex <> 0 Then 'esta online
        If UserList(MemberIndex).GuildInfo.EsGuildLeader = 1 Then
            Call Senddata(ToGuildMembers, MemberIndex, 0, "Y188")
            Exit Sub
        End If
        Call Senddata(ToIndex, MemberIndex, 0, "Y189")
        Call AddtoVar(UserList(MemberIndex).GuildInfo.Echadas, 1, 1000)
        UserList(MemberIndex).GuildInfo.GuildPoints = 0
        UserList(MemberIndex).GuildInfo.GuildName = ""
        Call Senddata(ToGuildMembers, UserIndex, 0, "||" & rdata & " fue expulsado del clan." & FONTTYPE_GUILD)
        '[Wizard 03/09/05] Forma burda de actualizar el nick ahorrar lineas, anchodebanda y clonacion de pjs jajajaja pero = es feo.
        Call WarpUserChar(MemberIndex, UserList(MemberIndex).Pos.Map, UserList(MemberIndex).Pos.X, UserList(MemberIndex).Pos.Y, False)
        NameeChado = UserList(MemberIndex).Name
Else

        If rs.State = 0 Then
        sql = "select * from usuarios where NickB = '" & rdata & "'"
        rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
        Else
        rs.Update
        rs.Close
        End If
     
        
        If val(rs!EsGuildLeaderB) = 1 Then
            Call Senddata(ToGuildMembers, MemberIndex, 0, "Y188")
            Exit Sub
        End If
        Call Senddata(ToIndex, MemberIndex, 0, "Y189")
        rs!EchadasB = val(rs!EchadasB) + 1
        NameeChado = rs!nickB
        rs!GuildNameB = ""
        rs!guildPtsB = 0
        Call Senddata(ToGuildMembers, UserIndex, 0, "||" & rdata & " fue expulsado del clan." & FONTTYPE_GUILD)
        rs.Update
        rs.Close
End If
Call oGuild.RemoveMember(NameeChado)
Exit Sub
hayerror:
Call LogError("Error en echar mienbro de clan")
End Sub

Public Sub DenyRequest(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub

Dim Soli As cSolicitud
Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub

Dim MemberIndex As Integer
MemberIndex = DameUserIndexConNombre(Soli.UserName)
If MemberIndex <> 0 Then 'esta online
    Call Senddata(ToIndex, MemberIndex, 0, "Y191")
    Call AddtoVar(UserList(MemberIndex).GuildInfo.SolicitudesRechazadas, 1, 10000)
End If
Call oGuild.RemoveSolicitud(Soli.UserName)
End Sub


Public Sub AcceptClanMember(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub
Dim Soli As cSolicitud
Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub

Dim MemberIndex As Integer
MemberIndex = DameUserIndexConNombre(Soli.UserName)
'Veamos si la alineacion le permite entrar;)
'[Wizard]
Select Case oGuild.CAlineacion
    Case 1 'Neutro:)
        If UserList(MemberIndex).Faccion.ArmadaReal _
        <> 0 Or UserList(MemberIndex).Faccion _
        .FuerzasCaos <> 0 Then Exit Sub
    Case 2 'Real
        If UserList(MemberIndex).Faccion.ArmadaReal _
        = 0 Then Exit Sub
    Case 3 'Caos
        If UserList(MemberIndex).Faccion.FuerzasCaos _
        = 0 Then Exit Sub
    Case Else 'Comete un biscocho
        Call LogError("ERROR EN ACCEPTMEMBER; ALINEACION")
        Exit Sub
End Select





Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub


MemberIndex = DameUserIndexConNombre(Soli.UserName)
If MemberIndex <> 0 Then 'esta online
    If UserList(MemberIndex).GuildInfo.GuildName <> "" Then
        Call Senddata(ToIndex, UserIndex, 0, "Y192")
        Exit Sub
    End If
    UserList(MemberIndex).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName
    UserList(MemberIndex).GuildInfo.CAlineacion = oGuild.CAlineacion
    Call AddtoVar(UserList(MemberIndex).GuildInfo.ClanesParticipo, 1, 1000)
    Call Senddata(ToIndex, MemberIndex, 0, "Y193")
    Call Senddata(ToIndex, MemberIndex, 0, "||Ahora sos un miembro activo del clan " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_GUILD)
    Call GiveGuildPoints(25, MemberIndex)
    Errorestaen = "accept"
    '[Wizard 03/09/05] Forma burda de actualizar el nick ahorrar lineas, anchodebanda y clonacion de pjs jajajaja pero = es feo.
    Call WarpUserChar(MemberIndex, UserList(MemberIndex).Pos.Map, UserList(MemberIndex).Pos.X, UserList(MemberIndex).Pos.Y, False)
        
Else
    Call Senddata(ToIndex, UserIndex, 0, "Y194")
    Exit Sub
End If
Call oGuild.Members.Add(Soli.UserName)
Call oGuild.RemoveSolicitud(Soli.UserName)
Call Senddata(ToGuildMembers, UserIndex, 0, "TW" & SND_ACEPTADOCLAN)
Call Senddata(ToGuildMembers, UserIndex, 0, "||" & rdata & " ha sido aceptado en el clan." & FONTTYPE_GUILD)
End Sub

Public Sub SendPeticion(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
    
Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub
Dim Soli As cSolicitud

Set Soli = oGuild.GetSolicitud(rdata)
If Soli Is Nothing Then Exit Sub
Call Senddata(ToIndex, UserIndex, 0, "PETICIO" & Soli.Desc)
End Sub

Public Sub SolicitudIngresoClan(ByVal UserIndex As Integer, ByVal Data As String)
If EsNewbie(UserIndex) Then
   Call Senddata(ToIndex, UserIndex, 0, "Y195")
   Exit Sub
End If
Dim MiSol As New cSolicitud
MiSol.Desc = ReadField(2, Data, 44)
MiSol.UserName = UserList(UserIndex).Name

Dim clan$
clan$ = ReadField(1, Data, 44)

Dim oGuild As cGuild
Set oGuild = FetchGuild(clan$)

If oGuild Is Nothing Then Exit Sub
If oGuild.IsMember(UserList(UserIndex).Name) Then Exit Sub
If Not oGuild.SolicitudesIncludes(MiSol.UserName) Then
        Call AddtoVar(UserList(UserIndex).GuildInfo.Solicitudes, 1, 1000)
        Call oGuild.TestSolicitudBound
        Call oGuild.Solicitudes.Add(MiSol)
        Call Senddata(ToIndex, UserIndex, 0, "Y196")
        Exit Sub
Else
        Call Senddata(ToIndex, UserIndex, 0, "Y197")
End If
End Sub

Public Sub SendCharInfo(ByVal UserName As String, ByVal UserIndex As Integer)
'¿Existe el personaje?
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub



sql = "select * from usuarios where NickB = '" & UserName & "'"
rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText



Dim MiUser As User
MiUser.Name = UserName
MiUser.Raza = rs!razaB
MiUser.Clase = rs!claseb
MiUser.Genero = rs!generoB
MiUser.Stats.ELV = rs!elvb
MiUser.Stats.GLD = rs!gldb
MiUser.Stats.Banco = rs!bancob
MiUser.Reputacion.Promedio = rs!promedioB

Dim H$
H$ = "CHRINFO" & UserName & ","
H$ = H$ & MiUser.Raza & ","
H$ = H$ & MiUser.Clase & ","
H$ = H$ & MiUser.Genero & ","
H$ = H$ & MiUser.Stats.ELV & ","
H$ = H$ & MiUser.Stats.GLD & ","
H$ = H$ & MiUser.Stats.Banco & ","
H$ = H$ & MiUser.Reputacion.Promedio & ","
MiUser.GuildInfo.FundoClan = rs!FundoClanB
MiUser.GuildInfo.EsGuildLeader = rs!EsGuildLeaderB
MiUser.GuildInfo.Echadas = rs!EchadasB
MiUser.GuildInfo.Solicitudes = rs!SolicitudesB
MiUser.GuildInfo.SolicitudesRechazadas = rs!SolicitudesRechazadasB
MiUser.GuildInfo.VecesFueGuildLeader = rs!VecesFueGuildLeaderB
'MiUser.GuildInfo.YaVoto = val(GetVar(UserFile, "Guild", "YaVoto"))
MiUser.GuildInfo.ClanesParticipo = rs!ClanesParticipoB
H$ = H$ & MiUser.GuildInfo.FundoClan & ","
H$ = H$ & MiUser.GuildInfo.EsGuildLeader & ","
H$ = H$ & MiUser.GuildInfo.Echadas & ","
H$ = H$ & MiUser.GuildInfo.Solicitudes & ","
H$ = H$ & MiUser.GuildInfo.SolicitudesRechazadas & ","
H$ = H$ & MiUser.GuildInfo.VecesFueGuildLeader & ","
H$ = H$ & MiUser.GuildInfo.ClanesParticipo & ","
MiUser.GuildInfo.ClanFundado = rs!ClanFundadoB
MiUser.GuildInfo.GuildName = rs!GuildNameB
H$ = H$ & MiUser.GuildInfo.ClanFundado & ","
H$ = H$ & MiUser.GuildInfo.GuildName & ","
MiUser.Faccion.ArmadaReal = rs!EjercitoRealB
MiUser.Faccion.FuerzasCaos = rs!ejercitocaosb
MiUser.Faccion.CiudadanosMatados = rs!CiudMatadosB
MiUser.Faccion.CriminalesMatados = rs!CrimMatadosB
H$ = H$ & MiUser.Faccion.ArmadaReal & ","
H$ = H$ & MiUser.Faccion.FuerzasCaos & ","
H$ = H$ & MiUser.Faccion.CiudadanosMatados & ","
rs.Close
Call Senddata(ToIndex, UserIndex, 0, H$)
End Sub

Public Sub UpdateGuildNews(ByVal rdata As String, ByVal UserIndex As Integer)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub
oGuild.GuildNews = rdata
End Sub

Public Sub UpdateCodexAndDesc(ByVal rdata As String, ByVal UserIndex As Integer)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub
Call oGuild.UpdateCodexAndDesc(rdata)
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim cad$, t%
'<-------Lista de guilds ---------->
cad$ = "LEADERI" & Guilds.Count & "¬"
For t% = 1 To Guilds.Count
    cad$ = cad$ & Guilds(t%).GuildName & "¬"
Next t%
Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub
'<-------Lista de miembros ---------->
cad$ = cad$ & oGuild.Members.Count & "¬"
For t% = 1 To oGuild.Members.Count
    cad$ = cad$ & oGuild.Members.Item(t%) & "¬"
Next t%
'<------- Guild News -------->
Dim GN$
GN$ = Replace(oGuild.GuildNews, vbCrLf, "º")
cad$ = cad$ & GN$ & "¬"
'<------- Solicitudes ------->
cad$ = cad$ & oGuild.Solicitudes.Count & "¬"
For t% = 1 To oGuild.Solicitudes.Count
    cad$ = cad$ & oGuild.Solicitudes.Item(t%).UserName & "¬"
Next t%
Call Senddata(ToIndex, UserIndex, 0, cad$)
End Sub

Public Sub SetNewURL(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub
oGuild.URL = rdata
Call Senddata(ToIndex, UserIndex, 0, "Y198")
End Sub

Public Sub DeclareAllie(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(rdata) Then Exit Sub

Dim LeaderGuild As cGuild, enemyGuild As cGuild
Set LeaderGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If LeaderGuild Is Nothing Then Exit Sub
Set enemyGuild = FetchGuild(rdata)
If enemyGuild Is Nothing Then Exit Sub
If LeaderGuild.IsEnemy(enemyGuild.GuildName) Then
        Call Senddata(ToIndex, UserIndex, 0, "Y199")
Else
   If Not LeaderGuild.IsAllie(enemyGuild.GuildName) Then
        Call LeaderGuild.AlliedGuilds.Add(enemyGuild.GuildName)
        Call enemyGuild.AlliedGuilds.Add(LeaderGuild.GuildName)
        Call Senddata(ToGuildMembers, UserIndex, 0, "||Tu clan ha firmado una alianza con " & enemyGuild.GuildName & FONTTYPE_GUILD)
        Call Senddata(ToGuildMembers, UserIndex, 0, "TW" & SND_DECLAREWAR)
        Dim Index As Integer
        Index = DameGuildMemberIndex(enemyGuild.GuildName)
        If Index <> 0 Then
            Call Senddata(ToGuildMembers, Index, 0, "||" & LeaderGuild.GuildName & " firmo una alianza con tu clan." & FONTTYPE_GUILD)
            Call Senddata(ToGuildMembers, Index, 0, "TW" & SND_DECLAREWAR)
        End If
   Else
        Call Senddata(ToIndex, UserIndex, 0, "Y200")
   End If
End If
End Sub

Public Sub DeclareWar(ByVal UserIndex As Integer, ByVal rdata As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(rdata) Then Exit Sub

Dim LeaderGuild As cGuild, enemyGuild As cGuild
Set LeaderGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If LeaderGuild Is Nothing Then Exit Sub
Set enemyGuild = FetchGuild(rdata)
If enemyGuild Is Nothing Then Exit Sub
If Not LeaderGuild.IsEnemy(enemyGuild.GuildName) Then
        Call LeaderGuild.RemoveAllie(enemyGuild.GuildName)
        Call enemyGuild.RemoveAllie(LeaderGuild.GuildName)
        Call LeaderGuild.EnemyGuilds.Add(enemyGuild.GuildName)
        Call enemyGuild.EnemyGuilds.Add(LeaderGuild.GuildName)
        Call Senddata(ToGuildMembers, UserIndex, 0, "||Tu clan le declaró la guerra a " & enemyGuild.GuildName & FONTTYPE_GUILD)
        Call Senddata(ToGuildMembers, UserIndex, 0, "TW" & SND_DECLAREWAR)
        Dim Index As Integer
        Index = DameGuildMemberIndex(enemyGuild.GuildName)
        If Index <> 0 Then
            Call Senddata(ToGuildMembers, Index, 0, "||" & LeaderGuild.GuildName & " le declaradó la guerra a tu clan." & FONTTYPE_GUILD)
            Call Senddata(ToGuildMembers, Index, 0, "TW" & SND_DECLAREWAR)
        End If
Else
   Call Senddata(ToIndex, UserIndex, 0, "||Tu clan ya esta en guerra con " & enemyGuild.GuildName & FONTTYPE_GUILD)
End If
End Sub

Public Function DameGuildMemberIndex(ByVal GuildName As String) As Integer
Dim LoopC As Integer
LoopC = 1
GuildName = UCase$(GuildName)
Do Until UCase$(UserList(LoopC).GuildInfo.GuildName) = GuildName
    LoopC = LoopC + 1
    If LoopC > MaxUsers Then
        DameGuildMemberIndex = 0
        Exit Function
    End If
Loop
DameGuildMemberIndex = LoopC
End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer)
If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub

Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub

Dim k$
k$ = "GUILDNE" & oGuild.GuildNews & "¬"

Dim t%
k$ = k$ & oGuild.EnemyGuilds.Count & "¬"
For t% = 1 To oGuild.EnemyGuilds.Count
    k$ = k$ & oGuild.EnemyGuilds(t%) & "¬"
Next t%
k$ = k$ & oGuild.AlliedGuilds.Count & "¬"
For t% = 1 To oGuild.AlliedGuilds.Count
    k$ = k$ & oGuild.AlliedGuilds(t%) & "¬"
Next t%
Call Senddata(ToIndex, UserIndex, 0, k$)
If oGuild.Elections Then
    Call Senddata(ToIndex, UserIndex, 0, "Y201")
    Call Senddata(ToIndex, UserIndex, 0, "Y202")
    Call Senddata(ToIndex, UserIndex, 0, "Y203")
    Call Senddata(ToIndex, UserIndex, 0, "Y204")
End If
End Sub

Public Sub SendGuildsList(ByVal UserIndex As Integer)
Dim cad$, t%
cad$ = "GL" & Guilds.Count & ","
For t% = 1 To Guilds.Count
    cad$ = cad$ & Guilds(t%).GuildName & ","
Next t%
Call Senddata(ToIndex, UserIndex, 0, cad$)
End Sub

Public Function FetchGuild(ByVal GuildName As String) As Object
Dim k As Integer
For k = 1 To Guilds.Count
    If UCase$(Guilds.Item(k).GuildName) = UCase$(GuildName) Then
            Set FetchGuild = Guilds.Item(k)
            Exit Function
    End If
Next k
Set FetchGuild = Nothing
End Function

Public Sub LoadGuildsDB()
Dim file As String, cant As Integer
file = App.Path & "\Guilds\" & "GuildsInfo.inf"

If Not FileExist(file, vbNormal) Then Exit Sub
cant = val(GetVar(file, "INIT", "NroGuilds"))

Dim NewGuild As cGuild
Dim k%
For k% = 1 To cant
    Set NewGuild = New cGuild
    Call NewGuild.InitializeGuildFromDisk(k%)
    Call Guilds.Add(NewGuild)
Next k%
End Sub

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String)
On Error GoTo errhandler
Dim oGuild As cGuild

If Guilds.Count = 0 Then Exit Sub
Set oGuild = FetchGuild(GuildName)
If oGuild Is Nothing Then Exit Sub

Dim cad$
cad$ = "CLANDET"
cad$ = cad$ & oGuild.GuildName
cad$ = cad$ & "¬" & oGuild.Founder
cad$ = cad$ & "¬" & oGuild.FundationDate
cad$ = cad$ & "¬" & oGuild.leader
cad$ = cad$ & "¬" & oGuild.URL
cad$ = cad$ & "¬" & oGuild.Members.Count
cad$ = cad$ & "¬" & oGuild.DaysToNextElection
cad$ = cad$ & "¬" & oGuild.Gold
cad$ = cad$ & "¬" & oGuild.EnemyGuilds.Count
cad$ = cad$ & "¬" & oGuild.AlliedGuilds.Count
cad$ = cad$ & "¬" & oGuild.CAlineacion
Dim codex$
codex$ = oGuild.CodexLenght()

Dim k%
For k% = 0 To oGuild.CodexLenght()
    codex$ = codex$ & "¬" & oGuild.GetCodex(k%)
Next k%
cad$ = cad$ & "¬" & codex$ & oGuild.Description
Call Senddata(ToIndex, UserIndex, 0, cad$)
errhandler:
End Sub

Public Function CanCreateGuild(ByVal UserIndex As Integer) As Boolean
'[Misery_Ezequiel 26/06/05]
If UserList(UserIndex).Stats.ELV < 25 Then
    CanCreateGuild = False
    Call Senddata(ToIndex, UserIndex, 0, "Y205")
    Exit Function
End If
'[\]Misery_Ezequiel 26/06/05]
If UserList(UserIndex).Stats.UserSkills(Liderazgo) < 90 Then
    CanCreateGuild = False
    Call Senddata(ToIndex, UserIndex, 0, "Y206")
    Exit Function
End If
CanCreateGuild = True
End Function

Public Function ExisteGuild(ByVal Name As String) As Boolean
Dim k As Integer
Name = UCase$(Name)
For k = 1 To Guilds.Count
    If UCase$(Guilds(k).GuildName) = Name Then
            ExisteGuild = True
            Exit Function
    End If
Next k
End Function

Public Function CreateGuild(ByVal Name As String, ByVal Rep As Long, ByVal Index As Integer, ByVal GuildInfo As String) As Boolean

If Not CanCreateGuild(Index) Then
    CreateGuild = False
    Exit Function
End If

Dim miclan As New cGuild


If Not miclan.Initialize(GuildInfo, Name, Rep) Then
    CreateGuild = False
    Call Senddata(ToIndex, Index, 0, "Y207")
    Exit Function
End If

'[Wizard] Revisamos q el fundador respete la alineacion del clan
Select Case miclan.CAlineacion
    Case 1 'Neutro
        If UserList(Index).Faccion.FuerzasCaos = 1 Or UserList(Index).Faccion.ArmadaReal = 1 Then
        CreateGuild = False
        Call Senddata(ToIndex, Index, 0, "||Para fundar un clan neutro no puedes ser del Caos ni de la Real" & FONTTYPE_INFO)
        Exit Function
        End If
    Case 2 'Real
        If UserList(Index).Faccion.ArmadaReal = 0 Then
        CreateGuild = False
        Call Senddata(ToIndex, Index, 0, "||Para fundar un clan de la Armada Real debes ser de la misma" & FONTTYPE_INFO)
        Exit Function
        End If
    Case 3 'Caos
        If UserList(Index).Faccion.FuerzasCaos = 0 Then
        CreateGuild = False
        Call Senddata(ToIndex, Index, 0, "||Para fundar un clan del caos debes ser del mismo" & FONTTYPE_INFO)
        Exit Function
        End If
    Case Else 'ErrOR
        Call LogError("Error en la ALINEACION!!!!")
        CreateGuild = False: Exit Function
End Select

If ExisteGuild(miclan.GuildName) Then
    CreateGuild = False
    Call Senddata(ToIndex, Index, 0, "Y208")
    Exit Function
End If

Call miclan.Members.Add(UCase$(UserList(Index).Name))
Call Guilds.Add(miclan, miclan.GuildName)

UserList(Index).GuildInfo.FundoClan = 1
UserList(Index).GuildInfo.EsGuildLeader = 1

Call AddtoVar(UserList(Index).GuildInfo.VecesFueGuildLeader, 1, 10000)
Call AddtoVar(UserList(Index).GuildInfo.ClanesParticipo, 1, 10000)

UserList(Index).GuildInfo.ClanFundado = miclan.GuildName
UserList(Index).GuildInfo.GuildName = UserList(Index).GuildInfo.ClanFundado

Call GiveGuildPoints(5000, Index)
Call Senddata(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call Senddata(ToAll, 0, 0, "||¡¡¡" & UserList(Index).Name & " fundo el clan '" & UserList(Index).GuildInfo.GuildName & "'!!!" & FONTTYPE_GUILD)
'[Wizard 03/09/05] Forma burda de actualizar el nick ahorrar lineas, anchodebanda y clonacion de pjs jajajaja pero = es feo.
Call WarpUserChar(Index, UserList(Index).Pos.Map, UserList(Index).Pos.X, UserList(Index).Pos.Y, False)
CreateGuild = True
End Function

Public Sub SaveGuildsDB()
Dim j As Integer
Dim file As String

file = App.Path & "\Guilds\" & "GuildsInfo.inf"
If FileExist(file, vbNormal) Then Kill file
Call WriteVar(file, "INIT", "NroGuilds", str(Guilds.Count))
For j = 1 To Guilds.Count
    Call Guilds(j).SaveGuild(file, j)
Next j
End Sub
'********************Misery_Ezequiel 28/05/05********************'
