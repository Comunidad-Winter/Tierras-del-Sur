Attribute VB_Name = "ES"
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

Option Explicit

Public Sub CargarSpawnList()

    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC


End Sub

Function EsDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False
End Function

Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim Nomb As String
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    Nomb = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    If Left(Nomb, 1) = "*" Or Left(Nomb, 1) = "+" Then Nomb = Right(Nomb, Len(Nomb) - 1)
    If UCase$(Name) = Nomb Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Public Function TxtDimension(ByVal Name As String) As Long
Dim N As Integer, cad As String, Tam As Long
N = FreeFile(1)
Open Name For Input As #N
Tam = 0
Do While Not EOF(N)
    Tam = Tam + 1
    Line Input #N, cad
Loop
Close N
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim N As Integer, i As Integer
N = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #N

For i = 1 To UBound(ForbidenNames)
    Line Input #N, ForbidenNames(i)
Next i

Close N

End Sub

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsLeerInis

Leer.Abrir DatPath & "Hechizos.dat"
'j = Val(Leer.DarValor(

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.DarValor("INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.Value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = Leer.DarValor("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).Desc = Leer.DarValor("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.DarValor("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.DarValor("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.DarValor("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.DarValor("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(Leer.DarValor("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = val(Leer.DarValor("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.DarValor("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(Leer.DarValor("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Resis = val(Leer.DarValor("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.DarValor("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.DarValor("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.DarValor("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.DarValor("Hechizo" & Hechizo, "MaxMana"))

    Hechizos(Hechizo).SubeSta = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.DarValor("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.DarValor("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.DarValor("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.DarValor("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.DarValor("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = val(Leer.DarValor("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverParalisis"))
'[Misery_Ezequiel 26/06/05]
    Hechizos(Hechizo).AgiUpAndFuer = val(Leer.DarValor("Hechizo" & Hechizo, "AgiUpAndFuer"))
    Hechizos(Hechizo).MinAgiFuer = val(Leer.DarValor("Hechizo" & Hechizo, "MinAgiFuer"))
    Hechizos(Hechizo).MaxAgiFuer = val(Leer.DarValor("Hechizo" & Hechizo, "MaxAgiFuer"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.DarValor("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
'[\]Misery_Ezequiel 26/06/05]
    Hechizos(Hechizo).CuraVeneno = val(Leer.DarValor("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.DarValor("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(Leer.DarValor("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.DarValor("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.DarValor("Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = val(Leer.DarValor("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(Leer.DarValor("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(Leer.DarValor("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(Leer.DarValor("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).cant = val(Leer.DarValor("Hechizo" & Hechizo, "Cant"))
       Hechizos(Hechizo).Mimetiza = val(Leer.DarValor("hechizo" & Hechizo, "Mimetiza"))
    
    Hechizos(Hechizo).Materializa = val(Leer.DarValor("Hechizo" & Hechizo, "Materializa"))
    Hechizos(Hechizo).ItemIndex = val(Leer.DarValor("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.DarValor("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.DarValor("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = val(Leer.DarValor("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = val(Leer.DarValor("Hechizo" & Hechizo, "Target"))
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
    

       'marche
    Hechizos(Hechizo).NeedStaff = val(Leer.DarValor("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(val(Leer.DarValor("Hechizo" & Hechizo, "StaffAffected")))
    'marche
      Dim i As Integer

    For i = 1 To NUMCLASES
      Hechizos(Hechizo).ClaseProhibida(i) = Leer.DarValor("Hechizo" & Hechizo, "CP" & i)
    Next
Next
Exit Sub

errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

'Public Sub CargarHechizos()
'On Error GoTo errhandler
'
'If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
'
'Dim Hechizo As Integer
'
''obtiene el numero de hechizos
'NumeroHechizos = val(GetVar(DatPath & "Hechizos.dat", "INIT", "NumeroHechizos"))
'ReDim Hechizos(1 To NumeroHechizos) As tHechizo
'
'frmCargando.cargar.Min = 0
'frmCargando.cargar.max = NumeroHechizos
'frmCargando.cargar.Value = 0
'
''Llena la lista
'For Hechizo = 1 To NumeroHechizos
'
'    Hechizos(Hechizo).Nombre = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Nombre")
'    Hechizos(Hechizo).Desc = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Desc")
'    Hechizos(Hechizo).PalabrasMagicas = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "PalabrasMagicas")
'
'    Hechizos(Hechizo).HechizeroMsg = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "HechizeroMsg")
'    Hechizos(Hechizo).TargetMsg = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "TargetMsg")
'    Hechizos(Hechizo).PropioMsg = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "PropioMsg")
'
'    Hechizos(Hechizo).Tipo = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Tipo"))
'    Hechizos(Hechizo).WAV = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "WAV"))
'    Hechizos(Hechizo).FXgrh = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Fxgrh"))
'
'    Hechizos(Hechizo).loops = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Loops"))
'
'    Hechizos(Hechizo).Resis = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Resis"))
'
'    Hechizos(Hechizo).SubeHP = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeHP"))
'    Hechizos(Hechizo).MinHP = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinHP"))
'    Hechizos(Hechizo).MaxHP = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxHP"))
'
'    Hechizos(Hechizo).SubeMana = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeMana"))
'    Hechizos(Hechizo).MiMana = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinMana"))
'    Hechizos(Hechizo).MaMana = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxMana"))
'
'    Hechizos(Hechizo).SubeSta = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeSta"))
'    Hechizos(Hechizo).MinSta = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinSta"))
'    Hechizos(Hechizo).MaxSta = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxSta"))
'
'    Hechizos(Hechizo).SubeHam = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeHam"))
'    Hechizos(Hechizo).MinHam = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinHam"))
'    Hechizos(Hechizo).MaxHam = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxHam"))
'
'    Hechizos(Hechizo).SubeSed = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeSed"))
'    Hechizos(Hechizo).MinSed = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinSed"))
'    Hechizos(Hechizo).MaxSed = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxSed"))
'
'    Hechizos(Hechizo).SubeAgilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeAG"))
'    Hechizos(Hechizo).MinAgilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinAG"))
'    Hechizos(Hechizo).MaxAgilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxAG"))
'
'    Hechizos(Hechizo).SubeFuerza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeFU"))
'    Hechizos(Hechizo).MinFuerza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinFU"))
'    Hechizos(Hechizo).MaxFuerza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxFU"))
'
'    Hechizos(Hechizo).SubeCarisma = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeCA"))
'    Hechizos(Hechizo).MinCarisma = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinCA"))
'    Hechizos(Hechizo).MaxCarisma = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxCA"))
'
'
'    Hechizos(Hechizo).Invisibilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Invisibilidad"))
'    Hechizos(Hechizo).Paraliza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Paraliza"))
'    Hechizos(Hechizo).RemoverParalisis = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "RemoverParalisis"))
'
'    Hechizos(Hechizo).CuraVeneno = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "CuraVeneno"))
'    Hechizos(Hechizo).Envenena = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Envenena"))
'    Hechizos(Hechizo).Maldicion = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Maldicion"))
'    Hechizos(Hechizo).RemoverMaldicion = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "RemoverMaldicion"))
'    Hechizos(Hechizo).Bendicion = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Bendicion"))
'    Hechizos(Hechizo).Revivir = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Revivir"))
'
'    Hechizos(Hechizo).Ceguera = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Ceguera"))
'    Hechizos(Hechizo).Estupidez = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Estupidez"))
'
'    Hechizos(Hechizo).Invoca = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Invoca"))
'    Hechizos(Hechizo).NumNpc = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "NumNpc"))
'    Hechizos(Hechizo).Cant = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Cant"))
'
'
'    Hechizos(Hechizo).Materializa = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Materializa"))
'    Hechizos(Hechizo).ItemIndex = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "ItemIndex"))
'
'    Hechizos(Hechizo).MinSkill = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinSkill"))
'    Hechizos(Hechizo).ManaRequerido = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "ManaRequerido"))
'
'    'Barrin 30/9/03
'    Hechizos(Hechizo).StaRequerido = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "StaRequerido"))
'
'    Hechizos(Hechizo).Target = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Target"))
'    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
'Next
'Exit Sub
'
'errhandler:
' MsgBox "Error cargando hechizos.dat"
'End Sub

Sub LoadMotd()
Dim i As Integer

MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    MOTD(i).Formato = ""
Next i

End Sub

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True
Dim i As Integer

''''''''''''''lo pongo aca x sugernecia del yind
For i = 1 To LastNPC
    If Npclist(i).flags.NPCActive Then
        If Npclist(i).Contadores.TiempoExistencia > 0 Then
            Call MuereNpc(i, 0)
        End If
    End If
Next i
'[Misery_Ezequiel 12/07/05]

'[\]Misery_Ezequiel 12/07/05]
'''''''''''/'lo pongo aca x sugernecia del yind

Call Senddata(ToAll, 0, 0, "BKW")

Call SaveGuildsDB
Call LimpiarMundo
Call WorldSave

Call Senddata(ToAll, 0, 0, "BKW")

Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile
End Sub


Public Sub SaveMapData(ByVal N As Integer)

'Call LogTarea("Sub SaveMapData N:" & n)

Dim LoopC As Integer
Dim TempInt As Integer
Dim Y As Integer
Dim X As Integer
Dim SaveAs As String

SaveAs = App.Path & "\WorldBackUP\Map" & N & ".map"

If FileExist(SaveAs, vbNormal) Then
    Kill SaveAs
End If

If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) Then
    Kill Left$(SaveAs, Len(SaveAs) - 4) & ".inf"
End If

'Open .map file
Open SaveAs For Binary As #1
Seek #1, 1
SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
SaveAs = SaveAs & ".inf"
'Open .inf file
Open SaveAs For Binary As #2
Seek #2, 1
'map Header
        
Put #1, , MapInfo(N).MapVersion
Put #1, , MiCabecera
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'inf Header
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt

'Write .map file
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        
        '.map file
        Put #1, , MapData(N, X, Y).Blocked
        
        For LoopC = 1 To 4
            Put #1, , MapData(N, X, Y).Graphic(LoopC)
        Next LoopC
        
        'Lugar vacio para futuras expansiones
        Put #1, , MapData(N, X, Y).trigger
        
        Put #1, , TempInt
        
        '.inf file
        'Tile exit
        Put #2, , MapData(N, X, Y).TileExit.Map
        Put #2, , MapData(N, X, Y).TileExit.X
        Put #2, , MapData(N, X, Y).TileExit.Y
        
        'NPC
        If MapData(N, X, Y).NpcIndex > 0 Then
            Put #2, , Npclist(MapData(N, X, Y).NpcIndex).Numero
        Else
            Put #2, , 0
        End If
        'Object
        
        If MapData(N, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(N, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_FOGATA Then
                MapData(N, X, Y).OBJInfo.ObjIndex = 0
                MapData(N, X, Y).OBJInfo.Amount = 0
            End If
'            If ObjData(MapData(n, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_MANCHAS Then
'                MapData(n, X, Y).OBJInfo.ObjIndex = 0
'                MapData(n, X, Y).OBJInfo.Amount = 0
'            End If
        End If
        
        Put #2, , MapData(N, X, Y).OBJInfo.ObjIndex
        Put #2, , MapData(N, X, Y).OBJInfo.Amount
        
        'Empty place holders for future expansion
        Put #2, , TempInt
        Put #2, , TempInt
        
    Next X
Next Y

'Close .map file
Close #1

'Close .inf file
Close #2

'write .dat file
SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
Call WriteVar(SaveAs, "Mapa" & N, "Name", MapInfo(N).Name)
Call WriteVar(SaveAs, "Mapa" & N, "MusicNum", MapInfo(N).Music)
Call WriteVar(SaveAs, "Mapa" & N, "StartPos", MapInfo(N).StartPos.Map & "-" & MapInfo(N).StartPos.X & "-" & MapInfo(N).StartPos.Y)

Call WriteVar(SaveAs, "Mapa" & N, "Terreno", MapInfo(N).Terreno)
Call WriteVar(SaveAs, "Mapa" & N, "Zona", MapInfo(N).Zona)
Call WriteVar(SaveAs, "Mapa" & N, "Restringir", MapInfo(N).Restringir)
Call WriteVar(SaveAs, "Mapa" & N, "BackUp", str(MapInfo(N).BackUp))
'[Misery_Ezequiel 27/06/05]
Call WriteVar(SaveAs, "Mapa" & N, "Nivel", MapInfo(N).Nivel)
'[\]Misery_Ezequiel 27/06/05]
If MapInfo(N).Pk Then
    Call WriteVar(SaveAs, "Mapa" & N, "pk", "0")
Else
    Call WriteVar(SaveAs, "Mapa" & N, "pk", "1")
End If

End Sub

Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc


End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadObjCarpintero()
Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
ReDim Preserve ObjCarpintero(1 To N) As Integer
For lc = 1 To N
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc
End Sub

Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsLeerInis

Leer.Abrir DatPath & "Obj.dat"
'j = val(Leer.DarValor("INIT", "NumObjs"))  '

'obtiene el numero de obj
NumObjDatas = val(Leer.DarValor("INIT", "NumObjs"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.Value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).Name = Leer.DarValor("OBJ" & Object, "Name")
    
    ObjData(Object).GrhIndex = val(Leer.DarValor("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    '[Wizard]Le carga esto a todos.
    ObjData(Object).SkillM = val(Leer.DarValor("OBJ" & Object, "SkillM"))
    '{/Wizard}
    
    ObjData(Object).ObjType = val(Leer.DarValor("OBJ" & Object, "ObjType"))
    ObjData(Object).SubTipo = val(Leer.DarValor("OBJ" & Object, "Subtipo"))
    
    ObjData(Object).Newbie = val(Leer.DarValor("OBJ" & Object, "Newbie"))
    
    If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
        ObjData(Object).SkillDefe = val(Leer.DarValor("OBJ" & Object, "SkillD"))
        ObjData(Object).ShieldAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
    End If
    
    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
        ObjData(Object).CascoAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
        ObjData(Object).SkillTacticassT = val(Leer.DarValor("OBJ" & Object, "SkillTT"))
    End If
    
    ObjData(Object).Ropaje = val(Leer.DarValor("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.DarValor("OBJ" & Object, "HechizoIndex"))
    
    If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
            ObjData(Object).WeaponAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.DarValor("OBJ" & Object, "Apuñala"))
            ObjData(Object).QuitaEnergia = val(Leer.DarValor("OBJ" & Object, "Sta"))
            ObjData(Object).SkillCombate = val(Leer.DarValor("OBJ" & Object, "SkillC"))
            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
            ObjData(Object).proyectil = val(Leer.DarValor("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.DarValor("OBJ" & Object, "Municiones"))
             ' marche
            ObjData(Object).StaffPower = val(Leer.DarValor("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.DarValor("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = val(Leer.DarValor("OBJ" & Object, "Refuerzo"))
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
            ObjData(Object).SkillTacticass = val(Leer.DarValor("OBJ" & Object, "SkillT"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
    End If
    '[Misery_Ezequiel 26/06/05]
    If ObjData(Object).ObjType = OBJTYPE_ANILLOS Then
            ObjData(Object).DefensaMagicaMax = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMax"))
            ObjData(Object).DefensaMagicaMin = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMin"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
    End If
    '[\]Misery_Ezequiel 26/06/05]
    If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
        ObjData(Object).Snd1 = val(Leer.DarValor("OBJ" & Object, "SND1"))
        ObjData(Object).Snd2 = val(Leer.DarValor("OBJ" & Object, "SND2"))
        ObjData(Object).Snd3 = val(Leer.DarValor("OBJ" & Object, "SND3"))
        ObjData(Object).MinInt = val(Leer.DarValor("OBJ" & Object, "MinInt"))
    End If
    
    ObjData(Object).LingoteIndex = val(Leer.DarValor("OBJ" & Object, "LingoteIndex"))
    
    If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
        ObjData(Object).MinSkill = val(Leer.DarValor("OBJ" & Object, "MinSkill"))
    End If
    
    ObjData(Object).MineralIndex = val(Leer.DarValor("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(Leer.DarValor("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.DarValor("OBJ" & Object, "MinHP"))
  
    
    ObjData(Object).Mujer = val(Leer.DarValor("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.DarValor("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.DarValor("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.DarValor("OBJ" & Object, "MinAgu"))
    
    
    ObjData(Object).MinDef = val(Leer.DarValor("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.DarValor("OBJ" & Object, "MAXDEF"))
    
    ObjData(Object).Respawn = val(Leer.DarValor("OBJ" & Object, "ReSpawn"))
    
    ObjData(Object).RazaEnana = val(Leer.DarValor("OBJ" & Object, "RazaEnana"))
    
    ObjData(Object).Valor = val(Leer.DarValor("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(Leer.DarValor("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(Leer.DarValor("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(Leer.DarValor("OBJ" & Object, "Llave"))
            ObjData(Object).clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    End If
    
    
    If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
        ObjData(Object).IndexAbierta = val(Leer.DarValor("OBJ" & Object, "IndexAbierta"))
        ObjData(Object).IndexCerrada = val(Leer.DarValor("OBJ" & Object, "IndexCerrada"))
        ObjData(Object).IndexCerradaLlave = val(Leer.DarValor("OBJ" & Object, "IndexCerradaLlave"))
    End If
    
    
    'Puertas y llaves
    ObjData(Object).clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.DarValor("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.DarValor("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.DarValor("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.DarValor("OBJ" & Object, "ID")
    
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = Leer.DarValor("OBJ" & Object, "CP" & i)
    Next
            
    ObjData(Object).Resistencia = val(Leer.DarValor("OBJ" & Object, "Resistencia"))
    ObjData(Object).DefensaMagicaMax = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMin"))
    'Pociones
    If ObjData(Object).ObjType = 11 Then
        ObjData(Object).TipoPocion = val(Leer.DarValor("OBJ" & Object, "TipoPocion"))
        ObjData(Object).MaxModificador = val(Leer.DarValor("OBJ" & Object, "MaxModificador"))
        ObjData(Object).MinModificador = val(Leer.DarValor("OBJ" & Object, "MinModificador"))
        ObjData(Object).DuracionEfecto = val(Leer.DarValor("OBJ" & Object, "DuracionEfecto"))
    End If

    ObjData(Object).SkCarpinteria = val(Leer.DarValor("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(Leer.DarValor("OBJ" & Object, "Madera"))
          ObjData(Object).MaderaT = val(Leer.DarValor("OBJ" & Object, "MaderaT"))
    If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.DarValor("OBJ" & Object, "Paraliza"))
    End If
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.DarValor("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = val(Leer.DarValor("OBJ" & Object, "NoSeCae"))
    
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
    'frmCargando.cargar.
    
    'DoEvents
Next Object

Exit Sub

errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

'Sub LoadOBJData()
'
''Call LogTarea("Sub LoadOBJData")
'
'On Error GoTo errhandler
'
'If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
'
''*****************************************************************
''Carga la lista de objetos
''*****************************************************************
'Dim Object As Integer
'
''obtiene el numero de obj
'NumObjDatas = val(Leer.DarValor("INIT", "NumObjs"))
'
'frmCargando.cargar.Min = 0
'frmCargando.cargar.max = NumObjDatas
'frmCargando.cargar.Value = 0
'
'
'ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
'
''Llena la lista
'For Object = 1 To NumObjDatas
'
'    ObjData(Object).Name = Leer.DarValor("OBJ" & Object, "Name")
'
'    ObjData(Object).GrhIndex = val(Leer.DarValor("OBJ" & Object, "GrhIndex"))
'
'    ObjData(Object).ObjType = val(Leer.DarValor("OBJ" & Object, "ObjType"))
'    ObjData(Object).SubTipo = val(Leer.DarValor("OBJ" & Object, "Subtipo"))
'
'    ObjData(Object).Newbie = val(Leer.DarValor("OBJ" & Object, "Newbie"))
'
'    If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
'        ObjData(Object).ShieldAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
'        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
'        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
'        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
'        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
'    End If
'
'    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
'        ObjData(Object).CascoAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
'        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
'        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
'        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
'        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
'    End If
'
'    ObjData(Object).Ropaje = val(Leer.DarValor("OBJ" & Object, "NumRopaje"))
'    ObjData(Object).HechizoIndex = val(Leer.DarValor("OBJ" & Object, "HechizoIndex"))
'
'    If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
'            ObjData(Object).WeaponAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
'            ObjData(Object).Apuñala = val(Leer.DarValor("OBJ" & Object, "Apuñala"))
'            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
'            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
'            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
'            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
'            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
'            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
'            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
'            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
'            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
'            ObjData(Object).proyectil = val(Leer.DarValor("OBJ" & Object, "Proyectil"))
'            ObjData(Object).Municion = val(Leer.DarValor("OBJ" & Object, "Municiones"))
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
'            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
'            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
'            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
'            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
'            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
'            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
'            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
'            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
'            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
'            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
'        ObjData(Object).Snd1 = val(Leer.DarValor("OBJ" & Object, "SND1"))
'        ObjData(Object).Snd2 = val(Leer.DarValor("OBJ" & Object, "SND2"))
'        ObjData(Object).Snd3 = val(Leer.DarValor("OBJ" & Object, "SND3"))
'        ObjData(Object).MinInt = val(Leer.DarValor("OBJ" & Object, "MinInt"))
'    End If
'
'    ObjData(Object).LingoteIndex = val(Leer.DarValor("OBJ" & Object, "LingoteIndex"))
'
'    If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
'        ObjData(Object).MinSkill = val(Leer.DarValor("OBJ" & Object, "MinSkill"))
'    End If
'
'    ObjData(Object).MineralIndex = val(Leer.DarValor("OBJ" & Object, "MineralIndex"))
'
'    ObjData(Object).MaxHP = val(Leer.DarValor("OBJ" & Object, "MaxHP"))
'    ObjData(Object).MinHP = val(Leer.DarValor("OBJ" & Object, "MinHP"))
'
'
'    ObjData(Object).Mujer = val(Leer.DarValor("OBJ" & Object, "Mujer"))
'    ObjData(Object).Hombre = val(Leer.DarValor("OBJ" & Object, "Hombre"))
'
'    ObjData(Object).MinHam = val(Leer.DarValor("OBJ" & Object, "MinHam"))
'    ObjData(Object).MinSed = val(Leer.DarValor("OBJ" & Object, "MinAgu"))
'
'
'    ObjData(Object).MinDef = val(Leer.DarValor("OBJ" & Object, "MINDEF"))
'    ObjData(Object).MaxDef = val(Leer.DarValor("OBJ" & Object, "MAXDEF"))
'
'    ObjData(Object).Respawn = val(Leer.DarValor("OBJ" & Object, "ReSpawn"))
'
'    ObjData(Object).RazaEnana = val(Leer.DarValor("OBJ" & Object, "RazaEnana"))
'
'    ObjData(Object).Valor = val(Leer.DarValor("OBJ" & Object, "Valor"))
'
'    ObjData(Object).Crucial = val(Leer.DarValor("OBJ" & Object, "Crucial"))
'
'    ObjData(Object).Cerrada = val(Leer.DarValor("OBJ" & Object, "abierta"))
'    If ObjData(Object).Cerrada = 1 Then
'            ObjData(Object).Llave = val(Leer.DarValor("OBJ" & Object, "Llave"))
'            ObjData(Object).Clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
'    End If
'
'
'    If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
'        ObjData(Object).IndexAbierta = val(Leer.DarValor("OBJ" & Object, "IndexAbierta"))
'        ObjData(Object).IndexCerrada = val(Leer.DarValor("OBJ" & Object, "IndexCerrada"))
'        ObjData(Object).IndexCerradaLlave = val(Leer.DarValor("OBJ" & Object, "IndexCerradaLlave"))
'    End If
'
'
'    'Puertas y llaves
'    ObjData(Object).Clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
'
'    ObjData(Object).texto = Leer.DarValor("OBJ" & Object, "Texto")
'    ObjData(Object).GrhSecundario = val(Leer.DarValor("OBJ" & Object, "VGrande"))
'
'    ObjData(Object).Agarrable = val(Leer.DarValor("OBJ" & Object, "Agarrable"))
'    ObjData(Object).ForoID = Leer.DarValor("OBJ" & Object, "ID")
'
'
'    Dim i As Integer
'    For i = 1 To NUMCLASES
'        ObjData(Object).ClaseProhibida(i) = Leer.DarValor("OBJ" & Object, "CP" & i)
'    Next
'
'    ObjData(Object).Resistencia = val(Leer.DarValor("OBJ" & Object, "Resistencia"))
'
'    'Pociones
'    If ObjData(Object).ObjType = 11 Then
'        ObjData(Object).TipoPocion = val(Leer.DarValor("OBJ" & Object, "TipoPocion"))
'        ObjData(Object).MaxModificador = val(Leer.DarValor("OBJ" & Object, "MaxModificador"))
'        ObjData(Object).MinModificador = val(Leer.DarValor("OBJ" & Object, "MinModificador"))
'        ObjData(Object).DuracionEfecto = val(Leer.DarValor("OBJ" & Object, "DuracionEfecto"))
'    End If
'
'    ObjData(Object).SkCarpinteria = val(Leer.DarValor("OBJ" & Object, "SkCarpinteria"))
'
'    If ObjData(Object).SkCarpinteria > 0 Then _
'        ObjData(Object).Madera = val(Leer.DarValor("OBJ" & Object, "Madera"))
'
'    If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
'            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
'            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
'            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
'            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
'    End If
'
'    'Bebidas
'    ObjData(Object).MinSta = val(Leer.DarValor("OBJ" & Object, "MinST"))
'
'    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
'
'
'    DoEvents
'Next Object
'
'Exit Sub
'
'errhandler:
'    MsgBox "error cargando objetos"
'
'
'End Sub

'Sub LoadOBJData_Nuevo()
'
''Call LogTarea("Sub LoadOBJData")
'
'On Error GoTo errhandler
''On Error GoTo 0
'
'If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
'
''*****************************************************************
''Carga la lista de objetos
''*****************************************************************
'Dim Object As Integer
'
'Dim A As Long, S As Long
'
'A = INICarga(DatPath & "Obj.dat")
'Call INIConf(A, 0, "", 0)
'
''obtiene el numero de obj
''NumObjDatas = val(GetVar(DatPath & "Obj.dat", "INIT", "NumObjs"))
'S = INIBuscarSeccion(A, "INIT")
'NumObjDatas = INIDarClaveInt(A, S, "NumOBJs")
'
'frmCargando.cargar.Min = 0
'frmCargando.cargar.max = NumObjDatas
'frmCargando.cargar.Value = 0
'
'
'ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
'
''Llena la lista
'For Object = 1 To NumObjDatas
'    S = INIBuscarSeccion(A, "OBJ" & Object)
'
'    'ObjData(Object).Name = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Name")
'    ObjData(Object).Name = INIDarClaveStr(A, S, "Name")
'
'    'ObjData(Object).GrhIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "GrhIndex"))
'    ObjData(Object).GrhIndex = INIDarClaveInt(A, S, "GrhIndex")
'
'    'ObjData(Object).ObjType = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ObjType"))
'    'ObjData(Object).SubTipo = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Subtipo"))
'
'    ObjData(Object).ObjType = INIDarClaveInt(A, S, "ObjType")
'    ObjData(Object).SubTipo = INIDarClaveInt(A, S, "Subtipo")
'
'    'ObjData(Object).Newbie = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Newbie"))
'    ObjData(Object).Newbie = INIDarClaveInt(A, S, "Newbie")
'
'    If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
''        ObjData(Object).ShieldAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
''        ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
''        ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
''        ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
''        ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
'        ObjData(Object).ShieldAnim = INIDarClaveInt(A, S, "Anim")
'        ObjData(Object).LingH = INIDarClaveInt(A, S, "LingH")
'        ObjData(Object).LingP = INIDarClaveInt(A, S, "LingP")
'        ObjData(Object).LingO = INIDarClaveInt(A, S, "LingO")
'        ObjData(Object).SkHerreria = INIDarClaveInt(A, S, "SkHerreria")
'    End If
'
'    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
''        ObjData(Object).CascoAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
''        ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
''        ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
''        ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
''        ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
'        ObjData(Object).CascoAnim = INIDarClaveInt(A, S, "Anim")
'        ObjData(Object).LingH = INIDarClaveInt(A, S, "LingH")
'        ObjData(Object).LingP = INIDarClaveInt(A, S, "LingP")
'        ObjData(Object).LingO = INIDarClaveInt(A, S, "LingO")
'        ObjData(Object).SkHerreria = INIDarClaveInt(A, S, "SkHerreria")
'    End If
'
''    ObjData(Object).Ropaje = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "NumRopaje"))
''    ObjData(Object).HechizoIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "HechizoIndex"))
'    ObjData(Object).Ropaje = INIDarClaveInt(A, S, "NumRopaje")
'    ObjData(Object).HechizoIndex = INIDarClaveInt(A, S, "HechizoIndex")
'
'    If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
''            ObjData(Object).WeaponAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
''            ObjData(Object).Apuñala = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Apuñala"))
''            ObjData(Object).Envenena = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Envenena"))
''            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
''            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
''            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
''            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
''            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
''            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
''            ObjData(Object).Real = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Real"))
''            ObjData(Object).Caos = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Caos"))
''            ObjData(Object).proyectil = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Proyectil"))
''            ObjData(Object).Municion = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Municiones"))
'
'            ObjData(Object).WeaponAnim = INIDarClaveInt(A, S, "Anim")
'            ObjData(Object).Apuñala = INIDarClaveInt(A, S, "Apuñala")
'            ObjData(Object).Envenena = INIDarClaveInt(A, S, "Envenena")
'            ObjData(Object).MaxHIT = INIDarClaveInt(A, S, "MaxHIT")
'            ObjData(Object).MinHIT = INIDarClaveInt(A, S, "MinHIT")
'            ObjData(Object).LingH = INIDarClaveInt(A, S, "LingH")
'            ObjData(Object).LingP = INIDarClaveInt(A, S, "LingP")
'            ObjData(Object).LingO = INIDarClaveInt(A, S, "LingO")
'            ObjData(Object).SkHerreria = INIDarClaveInt(A, S, "SkHerreria")
'            ObjData(Object).Real = INIDarClaveInt(A, S, "Real")
'            ObjData(Object).Caos = INIDarClaveInt(A, S, "Caos")
'            ObjData(Object).proyectil = INIDarClaveInt(A, S, "Proyectil")
'            ObjData(Object).Municion = INIDarClaveInt(A, S, "Municiones")
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
''            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
''            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
''            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
''            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
''            ObjData(Object).Real = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Real"))
''            ObjData(Object).Caos = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Caos"))
'            ObjData(Object).LingH = INIDarClaveInt(A, S, "LingH")
'            ObjData(Object).LingP = INIDarClaveInt(A, S, "LingP")
'            ObjData(Object).LingO = INIDarClaveInt(A, S, "LingO")
'            ObjData(Object).SkHerreria = INIDarClaveInt(A, S, "SkHerreria")
'            ObjData(Object).Real = INIDarClaveInt(A, S, "Real")
'            ObjData(Object).Caos = INIDarClaveInt(A, S, "Caos")
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
''            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
''            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
''            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
''            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
'            ObjData(Object).LingH = INIDarClaveInt(A, S, "LingH")
'            ObjData(Object).LingP = INIDarClaveInt(A, S, "LingP")
'            ObjData(Object).LingO = INIDarClaveInt(A, S, "LingO")
'            ObjData(Object).SkHerreria = INIDarClaveInt(A, S, "SkHerreria")
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
''        ObjData(Object).Snd1 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND1"))
''        ObjData(Object).Snd2 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND1"))
''        ObjData(Object).Snd3 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND3"))
''        ObjData(Object).MinInt = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinInt"))
'        ObjData(Object).Snd1 = INIDarClaveInt(A, S, "SND1")
'        ObjData(Object).Snd2 = INIDarClaveInt(A, S, "SND2")
'        ObjData(Object).Snd3 = INIDarClaveInt(A, S, "SND3")
'        ObjData(Object).MinInt = INIDarClaveInt(A, S, "MinInt")
'    End If
'
'    'ObjData(Object).LingoteIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingoteIndex"))
'    ObjData(Object).LingoteIndex = INIDarClaveInt(A, S, "LingoteIndex")
'
'    If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
'        'ObjData(Object).MinSkill = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinSkill"))
'        ObjData(Object).MinSkill = INIDarClaveInt(A, S, "MinSkill")
'    End If
'
''    ObjData(Object).MineralIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MineralIndex"))
''
''    ObjData(Object).MaxHP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHP"))
''    ObjData(Object).MinHP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHP"))
''
''
''    ObjData(Object).Mujer = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Mujer"))
''    ObjData(Object).Hombre = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Hombre"))
''
''    ObjData(Object).MinHam = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHam"))
''    ObjData(Object).MinSed = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinAgu"))
'
'    ObjData(Object).MineralIndex = INIDarClaveInt(A, S, "MineralIndex")
'
'    ObjData(Object).MaxHP = INIDarClaveInt(A, S, "MaxHP")
'    ObjData(Object).MinHP = INIDarClaveInt(A, S, "MinHP")
'
'    ObjData(Object).Mujer = INIDarClaveInt(A, S, "Mujer")
'    ObjData(Object).Hombre = INIDarClaveInt(A, S, "Hombre")
'
'    ObjData(Object).MinHam = INIDarClaveInt(A, S, "MinHam")
'    ObjData(Object).MinSed = INIDarClaveInt(A, S, "MinAgu")
'
'
''    ObjData(Object).MinDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MINDEF"))
''    ObjData(Object).MaxDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MAXDEF"))
''
''    ObjData(Object).Respawn = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ReSpawn"))
''
''    ObjData(Object).RazaEnana = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "RazaEnana"))
''
''    ObjData(Object).Valor = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Valor"))
''
''    ObjData(Object).Crucial = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Crucial"))
''
''    ObjData(Object).Cerrada = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "abierta"))
'
'    ObjData(Object).MinDef = INIDarClaveInt(A, S, "MINDEF")
'    ObjData(Object).MaxDef = INIDarClaveInt(A, S, "MAXDEF")
'
'    ObjData(Object).Respawn = INIDarClaveInt(A, S, "ReSpawn")
'
'    ObjData(Object).RazaEnana = INIDarClaveInt(A, S, "RazaEnana")
'
'    ObjData(Object).Valor = INIDarClaveInt(A, S, "Valor")
'
'    ObjData(Object).Crucial = INIDarClaveInt(A, S, "Crucial")
'
'    ObjData(Object).Cerrada = INIDarClaveInt(A, S, "abierta")
'
'    If ObjData(Object).Cerrada = 1 Then
''            ObjData(Object).Llave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Llave"))
''            ObjData(Object).Clave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Clave"))
'            ObjData(Object).Llave = INIDarClaveInt(A, S, "Llave")
'            ObjData(Object).Clave = INIDarClaveInt(A, S, "Clave")
'    End If
'
'
'    If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
''        ObjData(Object).IndexAbierta = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexAbierta"))
''        ObjData(Object).IndexCerrada = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexCerrada"))
''        ObjData(Object).IndexCerradaLlave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexCerradaLlave"))
'        ObjData(Object).IndexAbierta = INIDarClaveInt(A, S, "IndexAbierta")
'        ObjData(Object).IndexCerrada = INIDarClaveInt(A, S, "IndexCerrada")
'        ObjData(Object).IndexCerradaLlave = INIDarClaveInt(A, S, "IndexCerradaLlave")
'    End If
'
'
'    'Puertas y llaves
''    ObjData(Object).Clave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Clave"))
''
''    ObjData(Object).texto = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Texto")
''    ObjData(Object).GrhSecundario = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "VGrande"))
''
''    ObjData(Object).Agarrable = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Agarrable"))
''    ObjData(Object).ForoID = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ID")
'    ObjData(Object).Clave = INIDarClaveInt(A, S, "Clave")
'
'    ObjData(Object).texto = INIDarClaveStr(A, S, "Texto")
'    ObjData(Object).GrhSecundario = INIDarClaveInt(A, S, "VGrande")
'
'    ObjData(Object).Agarrable = INIDarClaveInt(A, S, "Agarrable")
'    ObjData(Object).ForoID = INIDarClaveStr(A, S, "ID")
'
'
'    Dim i As Integer
'    For i = 1 To NUMCLASES
'        'ObjData(Object).ClaseProhibida(i) = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "CP" & i)
'        ObjData(Object).ClaseProhibida(i) = INIDarClaveStr(A, S, "CP" & i)
'    Next
'
'    'ObjData(Object).Resistencia = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Resistencia"))
'    ObjData(Object).Resistencia = INIDarClaveInt(A, S, "Resistencia")
'
'    'Pociones
'    If ObjData(Object).ObjType = 11 Then
''        ObjData(Object).TipoPocion = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "TipoPocion"))
''        ObjData(Object).MaxModificador = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxModificador"))
''        ObjData(Object).MinModificador = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinModificador"))
''        ObjData(Object).DuracionEfecto = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "DuracionEfecto"))
'        ObjData(Object).TipoPocion = INIDarClaveInt(A, S, "TipoPocion")
'        ObjData(Object).MaxModificador = INIDarClaveInt(A, S, "MaxModificador")
'        ObjData(Object).MinModificador = INIDarClaveInt(A, S, "MinModificador")
'        ObjData(Object).DuracionEfecto = INIDarClaveInt(A, S, "DuracionEfecto")
'
'    End If
'
''    ObjData(Object).SkCarpinteria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkCarpinteria"))
'    ObjData(Object).SkCarpinteria = INIDarClaveInt(A, S, "SkCarpinteria")
'
'    If ObjData(Object).SkCarpinteria > 0 Then
'        'ObjData(Object).Madera = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Madera"))
'        ObjData(Object).Madera = INIDarClaveInt(A, S, "Madera")
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
''            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
''            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
'            ObjData(Object).MaxHIT = INIDarClaveInt(A, S, "MaxHIT")
'            ObjData(Object).MinHIT = INIDarClaveInt(A, S, "MinHIT")
'    End If
'
'    If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
''            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
''            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
'            ObjData(Object).MaxHIT = INIDarClaveInt(A, S, "MaxHIT")
'            ObjData(Object).MinHIT = INIDarClaveInt(A, S, "MinHIT")
'    End If
'
'    'Bebidas
'    'ObjData(Object).MinSta = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinST"))
'    ObjData(Object).MinSta = INIDarClaveInt(A, S, "MinST")
'
'    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
'
'
'    'DoEvents
'Next Object
'
'
'Call INIDescarga(A)
'
'Exit Sub
'
'errhandler:
'
'Call INIDescarga(A)
'
'    MsgBox "error cargando objetos: " & Err.number & " : " & Err.Description
'
'
'End Sub



Sub LoadUserStats(UserIndex As Integer, UserFile As String)



Dim LoopC As Integer



    UserList(UserIndex).Stats.UserAtributos(1) = rs!AT1
    UserList(UserIndex).Stats.UserAtributos(2) = rs!AT2
    UserList(UserIndex).Stats.UserAtributos(3) = rs!AT3
    UserList(UserIndex).Stats.UserAtributos(4) = rs!AT4
    UserList(UserIndex).Stats.UserAtributos(5) = rs!AT5
    
For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next

'Veces echado por el anticheat
'UserList(UserIndex).Stats.Veceshechado = rs!VecesCheat
'Voto caos
    UserList(UserIndex).Stats.VotC = rs!vot

'Cargacion de Skill

    UserList(UserIndex).Stats.UserSkills(1) = rs!SK1
    UserList(UserIndex).Stats.UserSkills(2) = rs!SK2
    UserList(UserIndex).Stats.UserSkills(3) = rs!SK3
    UserList(UserIndex).Stats.UserSkills(4) = rs!SK4
    UserList(UserIndex).Stats.UserSkills(5) = rs!SK5
    UserList(UserIndex).Stats.UserSkills(6) = rs!SK6
    UserList(UserIndex).Stats.UserSkills(7) = rs!SK7
    UserList(UserIndex).Stats.UserSkills(8) = rs!SK8
    UserList(UserIndex).Stats.UserSkills(9) = rs!SK9
    UserList(UserIndex).Stats.UserSkills(10) = rs!SK10
    UserList(UserIndex).Stats.UserSkills(11) = rs!SK11
    UserList(UserIndex).Stats.UserSkills(12) = rs!SK12
    UserList(UserIndex).Stats.UserSkills(13) = rs!SK13
    UserList(UserIndex).Stats.UserSkills(14) = rs!SK14
    UserList(UserIndex).Stats.UserSkills(15) = rs!SK15
    UserList(UserIndex).Stats.UserSkills(16) = rs!SK16
    UserList(UserIndex).Stats.UserSkills(17) = rs!SK17
    UserList(UserIndex).Stats.UserSkills(18) = rs!SK18
    UserList(UserIndex).Stats.UserSkills(19) = rs!SK19
    UserList(UserIndex).Stats.UserSkills(20) = rs!SK20
    UserList(UserIndex).Stats.UserSkills(21) = rs!SK21


'Carga los echizos!

UserList(UserIndex).Stats.UserHechizos(1) = rs!H1
UserList(UserIndex).Stats.UserHechizos(2) = rs!H2
UserList(UserIndex).Stats.UserHechizos(3) = rs!H3
UserList(UserIndex).Stats.UserHechizos(4) = rs!H4
UserList(UserIndex).Stats.UserHechizos(5) = rs!H5
UserList(UserIndex).Stats.UserHechizos(6) = rs!H6
UserList(UserIndex).Stats.UserHechizos(7) = rs!H7
UserList(UserIndex).Stats.UserHechizos(8) = rs!H8
UserList(UserIndex).Stats.UserHechizos(9) = rs!H9
UserList(UserIndex).Stats.UserHechizos(10) = rs!H10
UserList(UserIndex).Stats.UserHechizos(11) = rs!H11
UserList(UserIndex).Stats.UserHechizos(12) = rs!H12
UserList(UserIndex).Stats.UserHechizos(13) = rs!H13
UserList(UserIndex).Stats.UserHechizos(14) = rs!H14
UserList(UserIndex).Stats.UserHechizos(15) = rs!H15
UserList(UserIndex).Stats.UserHechizos(16) = rs!H16
UserList(UserIndex).Stats.UserHechizos(17) = rs!H17
UserList(UserIndex).Stats.UserHechizos(18) = rs!H18
UserList(UserIndex).Stats.UserHechizos(19) = rs!H19
UserList(UserIndex).Stats.UserHechizos(20) = rs!H20
UserList(UserIndex).Stats.UserHechizos(21) = rs!H21
UserList(UserIndex).Stats.UserHechizos(22) = rs!H22
UserList(UserIndex).Stats.UserHechizos(23) = rs!H23
UserList(UserIndex).Stats.UserHechizos(24) = rs!H24
UserList(UserIndex).Stats.UserHechizos(25) = rs!H25
UserList(UserIndex).Stats.UserHechizos(26) = rs!H26
UserList(UserIndex).Stats.UserHechizos(27) = rs!H27
UserList(UserIndex).Stats.UserHechizos(28) = rs!H28
UserList(UserIndex).Stats.UserHechizos(29) = rs!H29
UserList(UserIndex).Stats.UserHechizos(30) = rs!H30
UserList(UserIndex).Stats.UserHechizos(31) = rs!H31
UserList(UserIndex).Stats.UserHechizos(32) = rs!H32
UserList(UserIndex).Stats.UserHechizos(33) = rs!H33
UserList(UserIndex).Stats.UserHechizos(34) = rs!H34
UserList(UserIndex).Stats.UserHechizos(35) = rs!H35


UserList(UserIndex).Stats.GLD = rs!gldb
UserList(UserIndex).Stats.Banco = rs!bancob

UserList(UserIndex).Stats.MET = rs!METB
UserList(UserIndex).Stats.MaxHP = rs!MaxHPB
UserList(UserIndex).Stats.MinHP = rs!MinHPB

UserList(UserIndex).Stats.FIT = rs!FITB
UserList(UserIndex).Stats.MinSta = rs!MinSTAB
UserList(UserIndex).Stats.MaxSta = rs!MaxStaB

UserList(UserIndex).Stats.MaxMAN = rs!MaxMANb
UserList(UserIndex).Stats.MinMAN = rs!MinMANB

UserList(UserIndex).Stats.MaxHIT = rs!MaxHITB
UserList(UserIndex).Stats.MinHIT = rs!MinHITB

UserList(UserIndex).Stats.MaxAGU = rs!MaxAGUB
UserList(UserIndex).Stats.MinAGU = rs!minAGUB

UserList(UserIndex).Stats.MaxHam = rs!MaxHAMB
UserList(UserIndex).Stats.MinHam = rs!MinHAMB

UserList(UserIndex).Stats.SkillPts = rs!SkillPtsLibresB

UserList(UserIndex).Stats.Exp = rs!EXPB
UserList(UserIndex).Stats.ELU = rs!ELUB
UserList(UserIndex).Stats.ELV = rs!elvb


UserList(UserIndex).Stats.UsuariosMatados = rs!UserMuertesB
UserList(UserIndex).Stats.CriminalesMatados = rs!CrimMuertesB
UserList(UserIndex).Stats.NPCsMuertos = rs!NpcsMuertesB

UserList(UserIndex).flags.PertAlCons = rs!PERTENECEB
UserList(UserIndex).flags.PertAlConsCaos = rs!PERTENECECAOSB


End Sub

Sub LoadUserReputacion(UserIndex As Integer, UserFile As String)

UserList(UserIndex).Reputacion.AsesinoRep = rs!AsesinoB
UserList(UserIndex).Reputacion.BandidoRep = rs!BandidoB
UserList(UserIndex).Reputacion.BurguesRep = rs!BurguesiaB
UserList(UserIndex).Reputacion.LadronesRep = rs!LadronesB
UserList(UserIndex).Reputacion.NobleRep = rs!NoblesB
UserList(UserIndex).Reputacion.PlebeRep = rs!PlebeB
UserList(UserIndex).Reputacion.Promedio = rs!promedioB

End Sub


Sub LoadUserInit(UserIndex As Integer, UserFile As String)


Dim LoopC As Integer
Dim ln As String
Dim ln2 As String
Dim Cantidad As Long

     
UserList(UserIndex).Faccion.ArmadaReal = rs!EjercitoRealB
UserList(UserIndex).Faccion.FuerzasCaos = rs!ejercitocaosb
UserList(UserIndex).Faccion.CiudadanosMatados = rs!CiudMatadosB
UserList(UserIndex).Faccion.CriminalesMatados = rs!CrimMatadosB
UserList(UserIndex).Faccion.RecibioArmaduraCaos = rs!rArCaosB
UserList(UserIndex).Faccion.RecibioArmaduraReal = rs!rArRealB
UserList(UserIndex).Faccion.RecibioExpInicialCaos = rs!rExCaosB
UserList(UserIndex).Faccion.RecibioExpInicialReal = rs!rExRealB
UserList(UserIndex).Faccion.RecompensasCaos = rs!recCaosB
UserList(UserIndex).Faccion.RecompensasReal = rs!recRealB
'oooooooooooooooooooooo
UserList(UserIndex).Stats.OroGanado = rs!OG
UserList(UserIndex).Stats.OroPerdido = rs!OP
UserList(UserIndex).Stats.RetosGanadoS = rs!RG
UserList(UserIndex).Stats.RetosPerdidosB = rs!RP
'ooooooooooooooo
UserList(UserIndex).flags.Muerto = rs!MuertoB
UserList(UserIndex).flags.Escondido = rs!EscondidoB
'[Wizard 03/09/05]
UserList(UserIndex).LastIP = rs!LastIPB
'[/Wizard]

UserList(UserIndex).flags.Hambre = rs!HambreB
UserList(UserIndex).flags.Sed = rs!SedB
UserList(UserIndex).flags.Desnudo = rs!DesnudoB

UserList(UserIndex).flags.Envenenado = rs!EnvenenadoB
UserList(UserIndex).flags.Paralizado = rs!ParalizadoB
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If

UserList(UserIndex).flags.Navegando = rs!NavegandoB

UserList(UserIndex).Counters.Pena = rs!penab
On Error Resume Next
UserList(UserIndex).flags.Penasas = rs!penasasb
UserList(UserIndex).Email = rs!EmailB

'Barrin 2/10/03
'UserList(UserIndex).Apadrinados = rs!ApadrinadosB

UserList(UserIndex).Genero = rs!generoB
UserList(UserIndex).Clase = rs!claseb
UserList(UserIndex).Raza = rs!razaB
UserList(UserIndex).Hogar = rs!HogarB
UserList(UserIndex).Char.Heading = rs!HeadingB


UserList(UserIndex).OrigChar.Head = rs!Headb
UserList(UserIndex).OrigChar.Body = rs!bodyb
UserList(UserIndex).OrigChar.WeaponAnim = rs!armab
UserList(UserIndex).OrigChar.ShieldAnim = rs!escudob
UserList(UserIndex).OrigChar.CascoAnim = rs!Cascob
UserList(UserIndex).OrigChar.Heading = rs!HeadingB




UserList(UserIndex).Desc = rs!DescB
'WARNING
UserList(UserIndex).Pos.Map = rs!mapb
UserList(UserIndex).Pos.X = rs!xb
UserList(UserIndex).Pos.Y = rs!yb


Dim loopd As Integer

'[MARCHE]--------------------------------------------------------------------
'***********************************************************************************

UserList(UserIndex).BancoInvent.Object(1).ObjIndex = val(ReadField(1, rs!Bobj1, 45))
UserList(UserIndex).BancoInvent.Object(1).Amount = val(ReadField(2, rs!Bobj1, 45))
UserList(UserIndex).BancoInvent.Object(2).ObjIndex = val(ReadField(1, rs!Bobj2, 45))
UserList(UserIndex).BancoInvent.Object(2).Amount = val(ReadField(2, rs!Bobj2, 45))
UserList(UserIndex).BancoInvent.Object(3).ObjIndex = val(ReadField(1, rs!Bobj3, 45))
UserList(UserIndex).BancoInvent.Object(3).Amount = val(ReadField(2, rs!Bobj3, 45))
UserList(UserIndex).BancoInvent.Object(4).ObjIndex = val(ReadField(1, rs!Bobj4, 45))
UserList(UserIndex).BancoInvent.Object(4).Amount = val(ReadField(2, rs!Bobj4, 45))
UserList(UserIndex).BancoInvent.Object(5).ObjIndex = val(ReadField(1, rs!Bobj5, 45))
UserList(UserIndex).BancoInvent.Object(5).Amount = val(ReadField(2, rs!Bobj5, 45))
UserList(UserIndex).BancoInvent.Object(6).ObjIndex = val(ReadField(1, rs!Bobj6, 45))
UserList(UserIndex).BancoInvent.Object(6).Amount = val(ReadField(2, rs!Bobj6, 45))
UserList(UserIndex).BancoInvent.Object(7).ObjIndex = val(ReadField(1, rs!Bobj7, 45))
UserList(UserIndex).BancoInvent.Object(7).Amount = val(ReadField(2, rs!Bobj7, 45))
UserList(UserIndex).BancoInvent.Object(8).ObjIndex = val(ReadField(1, rs!Bobj8, 45))
UserList(UserIndex).BancoInvent.Object(8).Amount = val(ReadField(2, rs!Bobj8, 45))
UserList(UserIndex).BancoInvent.Object(9).ObjIndex = val(ReadField(1, rs!Bobj9, 45))
UserList(UserIndex).BancoInvent.Object(9).Amount = val(ReadField(2, rs!Bobj9, 45))
UserList(UserIndex).BancoInvent.Object(10).ObjIndex = val(ReadField(1, rs!Bobj10, 45))
UserList(UserIndex).BancoInvent.Object(10).Amount = val(ReadField(2, rs!Bobj10, 45))
UserList(UserIndex).BancoInvent.Object(11).ObjIndex = val(ReadField(1, rs!Bobj11, 45))
UserList(UserIndex).BancoInvent.Object(11).Amount = val(ReadField(2, rs!Bobj11, 45))
UserList(UserIndex).BancoInvent.Object(12).ObjIndex = val(ReadField(1, rs!Bobj12, 45))
UserList(UserIndex).BancoInvent.Object(12).Amount = val(ReadField(2, rs!Bobj12, 45))
UserList(UserIndex).BancoInvent.Object(13).ObjIndex = val(ReadField(1, rs!Bobj13, 45))
UserList(UserIndex).BancoInvent.Object(13).Amount = val(ReadField(2, rs!Bobj13, 45))
UserList(UserIndex).BancoInvent.Object(14).ObjIndex = val(ReadField(1, rs!Bobj14, 45))
UserList(UserIndex).BancoInvent.Object(14).Amount = val(ReadField(2, rs!Bobj14, 45))
UserList(UserIndex).BancoInvent.Object(15).ObjIndex = val(ReadField(1, rs!Bobj15, 45))
UserList(UserIndex).BancoInvent.Object(15).Amount = val(ReadField(2, rs!Bobj15, 45))
UserList(UserIndex).BancoInvent.Object(16).ObjIndex = val(ReadField(1, rs!Bobj16, 45))
UserList(UserIndex).BancoInvent.Object(16).Amount = val(ReadField(2, rs!Bobj16, 45))
UserList(UserIndex).BancoInvent.Object(17).ObjIndex = val(ReadField(1, rs!Bobj17, 45))
UserList(UserIndex).BancoInvent.Object(17).Amount = val(ReadField(2, rs!Bobj17, 45))
UserList(UserIndex).BancoInvent.Object(18).ObjIndex = val(ReadField(1, rs!Bobj18, 45))
UserList(UserIndex).BancoInvent.Object(18).Amount = val(ReadField(2, rs!Bobj18, 45))
UserList(UserIndex).BancoInvent.Object(19).ObjIndex = val(ReadField(1, rs!Bobj19, 45))
UserList(UserIndex).BancoInvent.Object(19).Amount = val(ReadField(2, rs!Bobj19, 45))
UserList(UserIndex).BancoInvent.Object(20).ObjIndex = val(ReadField(1, rs!Bobj20, 45))
UserList(UserIndex).BancoInvent.Object(20).Amount = val(ReadField(2, rs!Bobj20, 45))
UserList(UserIndex).BancoInvent.Object(21).ObjIndex = val(ReadField(1, rs!Bobj21, 45))
UserList(UserIndex).BancoInvent.Object(21).Amount = val(ReadField(2, rs!Bobj21, 45))
UserList(UserIndex).BancoInvent.Object(22).ObjIndex = val(ReadField(1, rs!Bobj22, 45))
UserList(UserIndex).BancoInvent.Object(22).Amount = val(ReadField(2, rs!Bobj22, 45))
UserList(UserIndex).BancoInvent.Object(23).ObjIndex = val(ReadField(1, rs!Bobj23, 45))
UserList(UserIndex).BancoInvent.Object(23).Amount = val(ReadField(2, rs!Bobj23, 45))
UserList(UserIndex).BancoInvent.Object(24).ObjIndex = val(ReadField(1, rs!Bobj24, 45))
UserList(UserIndex).BancoInvent.Object(24).Amount = val(ReadField(2, rs!Bobj24, 45))
UserList(UserIndex).BancoInvent.Object(25).ObjIndex = val(ReadField(1, rs!Bobj25, 45))
UserList(UserIndex).BancoInvent.Object(25).Amount = val(ReadField(2, rs!Bobj25, 45))
UserList(UserIndex).BancoInvent.Object(26).ObjIndex = val(ReadField(1, rs!Bobj26, 45))
UserList(UserIndex).BancoInvent.Object(26).Amount = val(ReadField(2, rs!Bobj26, 45))
UserList(UserIndex).BancoInvent.Object(27).ObjIndex = val(ReadField(1, rs!Bobj27, 45))
UserList(UserIndex).BancoInvent.Object(27).Amount = val(ReadField(2, rs!Bobj27, 45))
UserList(UserIndex).BancoInvent.Object(28).ObjIndex = val(ReadField(1, rs!Bobj28, 45))
UserList(UserIndex).BancoInvent.Object(28).Amount = val(ReadField(2, rs!Bobj28, 45))
UserList(UserIndex).BancoInvent.Object(29).ObjIndex = val(ReadField(1, rs!Bobj29, 45))
UserList(UserIndex).BancoInvent.Object(29).Amount = val(ReadField(2, rs!Bobj29, 45))
UserList(UserIndex).BancoInvent.Object(30).ObjIndex = val(ReadField(1, rs!Bobj30, 45))
UserList(UserIndex).BancoInvent.Object(30).Amount = val(ReadField(2, rs!Bobj30, 45))
UserList(UserIndex).BancoInvent.Object(31).ObjIndex = val(ReadField(1, rs!Bobj31, 45))
UserList(UserIndex).BancoInvent.Object(31).Amount = val(ReadField(2, rs!Bobj31, 45))
UserList(UserIndex).BancoInvent.Object(32).ObjIndex = val(ReadField(1, rs!Bobj32, 45))
UserList(UserIndex).BancoInvent.Object(32).Amount = val(ReadField(2, rs!Bobj32, 45))
UserList(UserIndex).BancoInvent.Object(33).ObjIndex = val(ReadField(1, rs!Bobj33, 45))
UserList(UserIndex).BancoInvent.Object(33).Amount = val(ReadField(2, rs!Bobj33, 45))
UserList(UserIndex).BancoInvent.Object(34).ObjIndex = val(ReadField(1, rs!Bobj34, 45))
UserList(UserIndex).BancoInvent.Object(34).Amount = val(ReadField(2, rs!Bobj34, 45))
UserList(UserIndex).BancoInvent.Object(35).ObjIndex = val(ReadField(1, rs!Bobj35, 45))
UserList(UserIndex).BancoInvent.Object(35).Amount = val(ReadField(2, rs!Bobj35, 45))
UserList(UserIndex).BancoInvent.Object(36).ObjIndex = val(ReadField(1, rs!Bobj36, 45))
UserList(UserIndex).BancoInvent.Object(36).Amount = val(ReadField(2, rs!Bobj36, 45))
UserList(UserIndex).BancoInvent.Object(37).ObjIndex = val(ReadField(1, rs!Bobj37, 45))
UserList(UserIndex).BancoInvent.Object(37).Amount = val(ReadField(2, rs!Bobj37, 45))
UserList(UserIndex).BancoInvent.Object(38).ObjIndex = val(ReadField(1, rs!Bobj38, 45))
UserList(UserIndex).BancoInvent.Object(38).Amount = val(ReadField(2, rs!Bobj38, 45))
UserList(UserIndex).BancoInvent.Object(39).ObjIndex = val(ReadField(1, rs!Bobj39, 45))
UserList(UserIndex).BancoInvent.Object(39).Amount = val(ReadField(2, rs!Bobj39, 45))
UserList(UserIndex).BancoInvent.Object(40).ObjIndex = val(ReadField(1, rs!Bobj40, 45))
UserList(UserIndex).BancoInvent.Object(40).Amount = val(ReadField(2, rs!Bobj40, 45))

'------------------------------------------------------------------------------------
'[/MARCHE]*****************************************************************************


'Lista de objetos

'TERMINAR
'For LoopC = 1 To MAX_INVENTORY_SLOTS
UserList(UserIndex).Invent.Object(1).ObjIndex = val(ReadField(1, rs!iOBJ1, 45))
UserList(UserIndex).Invent.Object(1).Amount = val(ReadField(2, rs!iOBJ1, 45))
UserList(UserIndex).Invent.Object(1).Equipped = val(ReadField(3, rs!iOBJ1, 45))
UserList(UserIndex).Invent.Object(2).ObjIndex = val(ReadField(1, rs!iOBJ2, 45))
UserList(UserIndex).Invent.Object(2).Amount = val(ReadField(2, rs!iOBJ2, 45))
UserList(UserIndex).Invent.Object(2).Equipped = val(ReadField(3, rs!iOBJ2, 45))
UserList(UserIndex).Invent.Object(3).ObjIndex = val(ReadField(1, rs!iOBJ3, 45))
UserList(UserIndex).Invent.Object(3).Amount = val(ReadField(2, rs!iOBJ3, 45))
UserList(UserIndex).Invent.Object(3).Equipped = val(ReadField(3, rs!iOBJ3, 45))
UserList(UserIndex).Invent.Object(4).ObjIndex = val(ReadField(1, rs!iOBJ4, 45))
UserList(UserIndex).Invent.Object(4).Amount = val(ReadField(2, rs!iOBJ4, 45))
UserList(UserIndex).Invent.Object(4).Equipped = val(ReadField(3, rs!iOBJ4, 45))
UserList(UserIndex).Invent.Object(5).ObjIndex = val(ReadField(1, rs!iOBJ5, 45))
UserList(UserIndex).Invent.Object(5).Amount = val(ReadField(2, rs!iOBJ5, 45))
UserList(UserIndex).Invent.Object(5).Equipped = val(ReadField(3, rs!iOBJ5, 45))
UserList(UserIndex).Invent.Object(6).ObjIndex = val(ReadField(1, rs!iOBJ6, 45))
UserList(UserIndex).Invent.Object(6).Amount = val(ReadField(2, rs!iOBJ6, 45))
UserList(UserIndex).Invent.Object(6).Equipped = val(ReadField(3, rs!iOBJ6, 45))
UserList(UserIndex).Invent.Object(7).ObjIndex = val(ReadField(1, rs!iOBJ7, 45))
UserList(UserIndex).Invent.Object(7).Amount = val(ReadField(2, rs!iOBJ7, 45))
UserList(UserIndex).Invent.Object(7).Equipped = val(ReadField(3, rs!iOBJ7, 45))
UserList(UserIndex).Invent.Object(8).ObjIndex = val(ReadField(1, rs!iOBJ8, 45))
UserList(UserIndex).Invent.Object(8).Amount = val(ReadField(2, rs!iOBJ8, 45))
UserList(UserIndex).Invent.Object(8).Equipped = val(ReadField(3, rs!iOBJ8, 45))
UserList(UserIndex).Invent.Object(9).ObjIndex = val(ReadField(1, rs!iOBJ9, 45))
UserList(UserIndex).Invent.Object(9).Amount = val(ReadField(2, rs!iOBJ9, 45))
UserList(UserIndex).Invent.Object(9).Equipped = val(ReadField(3, rs!iOBJ9, 45))
UserList(UserIndex).Invent.Object(10).ObjIndex = val(ReadField(1, rs!iOBJ10, 45))
UserList(UserIndex).Invent.Object(10).Amount = val(ReadField(2, rs!iOBJ10, 45))
UserList(UserIndex).Invent.Object(10).Equipped = val(ReadField(3, rs!iOBJ10, 45))
UserList(UserIndex).Invent.Object(11).ObjIndex = val(ReadField(1, rs!iOBJ11, 45))
UserList(UserIndex).Invent.Object(11).Amount = val(ReadField(2, rs!iOBJ11, 45))
UserList(UserIndex).Invent.Object(11).Equipped = val(ReadField(3, rs!iOBJ11, 45))
UserList(UserIndex).Invent.Object(12).ObjIndex = val(ReadField(1, rs!iOBJ12, 45))
UserList(UserIndex).Invent.Object(12).Amount = val(ReadField(2, rs!iOBJ12, 45))
UserList(UserIndex).Invent.Object(12).Equipped = val(ReadField(3, rs!iOBJ12, 45))
UserList(UserIndex).Invent.Object(13).ObjIndex = val(ReadField(1, rs!iOBJ13, 45))
UserList(UserIndex).Invent.Object(13).Amount = val(ReadField(2, rs!iOBJ13, 45))
UserList(UserIndex).Invent.Object(13).Equipped = val(ReadField(3, rs!iOBJ13, 45))
UserList(UserIndex).Invent.Object(14).ObjIndex = val(ReadField(1, rs!iOBJ14, 45))
UserList(UserIndex).Invent.Object(14).Amount = val(ReadField(2, rs!iOBJ14, 45))
UserList(UserIndex).Invent.Object(14).Equipped = val(ReadField(3, rs!iOBJ14, 45))

UserList(UserIndex).Invent.Object(15).ObjIndex = val(ReadField(1, rs!iOBJ15, 45))
UserList(UserIndex).Invent.Object(15).Amount = val(ReadField(2, rs!iOBJ15, 45))
UserList(UserIndex).Invent.Object(15).Equipped = val(ReadField(3, rs!iOBJ15, 45))
UserList(UserIndex).Invent.Object(16).ObjIndex = val(ReadField(1, rs!iOBJ16, 45))
UserList(UserIndex).Invent.Object(16).Amount = val(ReadField(2, rs!iOBJ16, 45))
UserList(UserIndex).Invent.Object(16).Equipped = val(ReadField(3, rs!iOBJ16, 45))
UserList(UserIndex).Invent.Object(17).ObjIndex = val(ReadField(1, rs!iOBJ17, 45))
UserList(UserIndex).Invent.Object(17).Amount = val(ReadField(2, rs!iOBJ17, 45))
UserList(UserIndex).Invent.Object(17).Equipped = val(ReadField(3, rs!iOBJ17, 45))
UserList(UserIndex).Invent.Object(18).ObjIndex = val(ReadField(1, rs!iOBJ18, 45))
UserList(UserIndex).Invent.Object(18).Amount = val(ReadField(2, rs!iOBJ18, 45))
UserList(UserIndex).Invent.Object(18).Equipped = val(ReadField(3, rs!iOBJ18, 45))
UserList(UserIndex).Invent.Object(19).ObjIndex = val(ReadField(1, rs!iOBJ19, 45))
UserList(UserIndex).Invent.Object(19).Amount = val(ReadField(2, rs!iOBJ19, 45))
UserList(UserIndex).Invent.Object(19).Equipped = val(ReadField(3, rs!iOBJ19, 45))
UserList(UserIndex).Invent.Object(20).ObjIndex = val(ReadField(1, rs!iOBJ20, 45))
UserList(UserIndex).Invent.Object(20).Amount = val(ReadField(2, rs!iOBJ20, 45))
UserList(UserIndex).Invent.Object(20).Equipped = val(ReadField(3, rs!iOBJ20, 45))




'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = rs!WeaponEqpSlotB
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = rs!ArmourEqpSlotB
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = rs!EscudoEqpSlotB
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = rs!CascoEqpSlotB
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = rs!BarcoSlotB
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = rs!MunicionSlotB
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto herramienta
UserList(UserIndex).Invent.HerramientaEqpSlot = rs!HerramientaSlotB
If UserList(UserIndex).Invent.HerramientaEqpSlot > 0 Then
    UserList(UserIndex).Invent.HerramientaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.HerramientaEqpSlot).ObjIndex
End If

'[Wizard 07/09/05] Baje esto, para poder hacerlo completo(Por el Inventario del barco)
If UserList(UserIndex).flags.Muerto = 0 Then 'Si esta vivo...
    If UserList(UserIndex).flags.Navegando = 0 Then 'Si no navega...
        UserList(UserIndex).Char = UserList(UserIndex).OrigChar
    Else 'Navega....
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
        UserList(UserIndex).Char.Head = 0
        UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
    End If
Else 'Esta muerto!
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    If UserList(UserIndex).flags.Navegando = 1 Then 'Ta navegando
        UserList(UserIndex).Char.Body = iFragataFantasmal
        UserList(UserIndex).Char.Head = 0
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos <> 0 Then 'Es caos
        UserList(UserIndex).Char.Body = iCuerpoMuertoCrimi
        UserList(UserIndex).Char.Head = iCabezaMuertoCrimi
    Else 'No navega y no es caos: Casper blanquito^^
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
    End If
End If
'/Wizard
'
' [Marche 20-4-04
UserList(UserIndex).NroMacotas = rs!NroMascotasB

'WARNING
Select Case rs!NroMascotasB
Case "0"
Case "1"
 UserList(UserIndex).MascotasType(1) = rs!mas1
Case "2"
 UserList(UserIndex).MascotasType(1) = rs!mas1
 UserList(UserIndex).MascotasType(2) = rs!mas2
Case "3"
 UserList(UserIndex).MascotasType(1) = rs!mas1
 UserList(UserIndex).MascotasType(2) = rs!mas2
 UserList(UserIndex).MascotasType(3) = rs!mas3
End Select

UserList(UserIndex).flags.Banrazon = 1
UserList(UserIndex).NroMacotas = rs!NroMascotasB
' que lio

UserList(UserIndex).GuildInfo.FundoClan = rs!FundoClanB
UserList(UserIndex).GuildInfo.EsGuildLeader = rs!EsGuildLeaderB
UserList(UserIndex).GuildInfo.Echadas = rs!EchadasB
UserList(UserIndex).GuildInfo.Solicitudes = rs!SolicitudesB
UserList(UserIndex).GuildInfo.SolicitudesRechazadas = rs!SolicitudesRechazadasB
UserList(UserIndex).GuildInfo.VecesFueGuildLeader = rs!VecesFueGuildLeaderB
UserList(UserIndex).GuildInfo.YaVoto = rs!YaVotoB
UserList(UserIndex).GuildInfo.ClanesParticipo = rs!ClanesParticipoB
UserList(UserIndex).GuildInfo.GuildPoints = rs!GuildPointsB
UserList(UserIndex).GuildInfo.ClanFundado = rs!ClanFundadoB
UserList(UserIndex).GuildInfo.GuildName = rs!GuildNameB
'[Wizard] SI tiene clan cargamos la alineacion
'desde el objeto Guild de su clan! yaaaaaaa que
'podria traer problemas, cargara lgo q no existe.
If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub
Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
UserList(UserIndex).GuildInfo.CAlineacion = oGuild.CAlineacion


End Sub





Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

'Sub CargarBackUp_Nuevo()
'
''Call LogTarea("Sub CargarBackUp")
'
'If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
'
'Dim Map As Integer
'Dim LoopC As Integer
'Dim X As Integer
'Dim Y As Integer
'Dim DummyInt As Integer
'Dim TempInt As Integer
'Dim SaveAs As String
'Dim NpcFile As String
'Dim Porc As Long
'Dim FileNamE As String
'Dim c$
'
'Dim archmap As String, archinf As String
'
'On Error GoTo man
'
'
'NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
'frmCargando.cargar.Min = 0
'frmCargando.cargar.max = NumMaps
'frmCargando.cargar.Value = 0
'
'MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
'
'ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
'ReDim MapInfo(1 To NumMaps) As MapInfo
'
'For Map = 1 To NumMaps
'
'    FileNamE = App.Path & "\WorldBackUp\Map" & Map & ".map"
'
'    If FileExist(FileNamE, vbNormal) Then
'        archmap = App.Path & "\WorldBackUp\Map" & Map & ".map"
'        archinf = App.Path & "\WorldBackUp\Map" & Map & ".inf"
'        c$ = App.Path & "\WorldBackUp\Map" & Map & ".dat"
'    Else
'        archmap = App.Path & MapPath & "Mapa" & Map & ".map"
'        archinf = App.Path & MapPath & "Mapa" & Map & ".inf"
'        c$ = App.Path & MapPath & "Mapa" & Map & ".dat"
'    End If
'
'        Call CargarUnMapa(Map, archmap, archinf)
'
'          frmCargando.cargar.Value = frmCargando.cargar.Value + 1
'
'          DoEvents
'Next Map
'
'FrmStat.Visible = False
'
'Exit Sub
'
'man:
'    MsgBox ("Error durante la carga de mapas.")
'    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
'
'
'
'End Sub

Sub CargarBackUp_Nuevo2()

'Call LogTarea("Sub CargarBackUp")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim SaveAs As String
Dim npcfile As String
Dim Porc As Long
Dim FileNamE As String
Dim c$
    
On Error GoTo man

 
NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.Value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo

Dim buffer(1 To ((YMaxMapSize - YMinMapSize + 1) * (XMaxMapSize - XMinMapSize + 1))) As TileMap
Dim buffer2(1 To ((YMaxMapSize - YMinMapSize + 1) * (XMaxMapSize - XMinMapSize + 1))) As TileInf
Dim idx As Integer

For Map = 1 To NumMaps

    FileNamE = App.Path & "\WorldBackUp\Map" & Map & ".map"
    
    If FileExist(FileNamE, vbNormal) Then
        Open App.Path & "\WorldBackUp\Map" & Map & ".map" For Binary As #1
        Open App.Path & "\WorldBackUp\Map" & Map & ".inf" For Binary As #2
        c$ = App.Path & "\WorldBackUp\Map" & Map & ".dat"
    Else
        Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
        Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
        c$ = App.Path & MapPath & "Mapa" & Map & ".dat"
    End If
    
    Seek #1, 1
    Seek #2, 1
    'map Header
    Get #1, , MapInfo(Map).MapVersion
    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    'inf Header
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    'Load arrays
                   
    Get #1, , buffer
    Get #2, , buffer2
    
    
    idx = 1
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            MapData(Map, X, Y).Blocked = buffer(idx).bloqueado
            MapData(Map, X, Y).Graphic(1) = buffer(idx).grafs(1)
            MapData(Map, X, Y).Graphic(2) = buffer(idx).grafs(2)
            MapData(Map, X, Y).Graphic(3) = buffer(idx).grafs(3)
            MapData(Map, X, Y).Graphic(4) = buffer(idx).grafs(4)
            MapData(Map, X, Y).trigger = buffer(idx).trigger
            
            MapData(Map, X, Y).TileExit.Map = buffer2(idx).dest_mapa
            MapData(Map, X, Y).TileExit.X = buffer2(idx).dest_x
            MapData(Map, X, Y).TileExit.Y = buffer2(idx).dest_y
            
            MapData(Map, X, Y).NpcIndex = buffer2(idx).npc
            If MapData(Map, X, Y).NpcIndex > 0 Then
                
                If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                Else
                        npcfile = DatPath & "NPCs.dat"
                End If
                
                'Si el npc debe hacer respawn en la pos
                'original la guardamos
                If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                Else
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                End If
                
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                
                Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
            End If

            If buffer2(idx).obj_ind > 0 And buffer2(idx).obj_ind <= UBound(ObjData) Then
                MapData(Map, X, Y).OBJInfo.ObjIndex = buffer2(idx).obj_ind
                MapData(Map, X, Y).OBJInfo.Amount = buffer2(idx).obj_cant
            Else
                MapData(Map, X, Y).OBJInfo.ObjIndex = 0
                MapData(Map, X, Y).OBJInfo.Amount = 0
            End If
            
            idx = idx + 1
        Next X
    Next Y

    Close #1
    Close #2
    MapInfo(Map).Name = GetVar(c$, "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(c$, "Mapa" & Map, "MusicNum")
    MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
    If val(GetVar(c$, "Mapa" & Map, "Pk")) = 0 Then
          MapInfo(Map).Pk = True
    Else
          MapInfo(Map).Pk = False
    End If
    MapInfo(Map).Restringir = GetVar(c$, "Mapa" & Map, "Restringir")
    MapInfo(Map).BackUp = val(GetVar(c$, "Mapa" & Map, "BackUp"))
    MapInfo(Map).Terreno = GetVar(c$, "Mapa" & Map, "Terreno")
    MapInfo(Map).Zona = GetVar(c$, "Mapa" & Map, "Zona")
    
    '[Misery_Ezequiel 27/06/05]
    MapInfo(Map).Nivel = GetVar(c$, "Mapa" & Map, "Nivel")
    '[\]Misery_Ezequiel 27/06/05]
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1

   DoEvents
Next Map

FrmStat.Visible = False

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas.")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

  

End Sub

Sub CargarBackUp()

'Call LogTarea("Sub CargarBackUp")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim SaveAs As String
Dim npcfile As String
Dim Porc As Long
Dim FileNamE As String
Dim c$
    
On Error GoTo man

 
NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.Value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    
    FileNamE = App.Path & "\WorldBackUp\Map" & Map & ".map"
    
    If FileExist(FileNamE, vbNormal) Then
        Open App.Path & "\WorldBackUp\Map" & Map & ".map" For Binary As #1
        Open App.Path & "\WorldBackUp\Map" & Map & ".inf" For Binary As #2
        c$ = App.Path & "\WorldBackUp\Map" & Map & ".dat"
    Else
        Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
        Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
        c$ = App.Path & MapPath & "Mapa" & Map & ".dat"
    End If
    
        Seek #1, 1
        Seek #2, 1
        'map Header
        Get #1, , MapInfo(Map).MapVersion
        Get #1, , MiCabecera
        Get #1, , TempInt
        Get #1, , TempInt
        Get #1, , TempInt
        Get #1, , TempInt
        'inf Header
        Get #2, , TempInt
        Get #2, , TempInt
        Get #2, , TempInt
        Get #2, , TempInt
        Get #2, , TempInt
        'Load arrays
        'DoEvents
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize
                    '.dat file
                    Get #1, , MapData(Map, X, Y).Blocked
                    
                    'Get GRH number
                    For LoopC = 1 To 4
                        Get #1, , MapData(Map, X, Y).Graphic(LoopC)
                    Next LoopC
                    
                    'Space holder for future expansion
                    Get #1, , MapData(Map, X, Y).trigger
                    Get #1, , TempInt
                    
                                        
                    '.inf file
                    Get #2, , MapData(Map, X, Y).TileExit.Map
                    Get #2, , MapData(Map, X, Y).TileExit.X
                    Get #2, , MapData(Map, X, Y).TileExit.Y
                    
                    'Get and make NPC
                    Get #2, , MapData(Map, X, Y).NpcIndex
                    If MapData(Map, X, Y).NpcIndex > 0 Then
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        
                        If Npclist(MapData(Map, X, Y).NpcIndex).Numero > 499 Then
                            npcfile = DatPath & "NPCs-HOSTILES.dat"
                        Else
                            npcfile = DatPath & "NPCs.dat"
                        End If
                        
                        Dim fl As Byte
                        fl = val(GetVar(npcfile, "NPC" & Npclist(MapData(Map, X, Y).NpcIndex).Numero, "PosOrig"))
                        If fl = 1 Then
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                        Else
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = 0
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = 0
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = 0
                        End If
        
                        Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                        
                        
                        'Si existe el backup lo cargamos
                        If Npclist(MapData(Map, X, Y).NpcIndex).flags.BackUp = 1 Then
                                'cargamos el nuevo del backup
                                Call CargarNpcBackUp(MapData(Map, X, Y).NpcIndex, Npclist(MapData(Map, X, Y).NpcIndex).Numero)
                                
                        End If
                        
                        Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                    End If

                    'Get and make Object
                    Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
                    Get #2, , MapData(Map, X, Y).OBJInfo.Amount
        
                    'Space holder for future expansion (Objects, ect.
                    Get #2, , DummyInt
                    Get #2, , DummyInt
            Next X
        Next Y
        Close #1
        Close #2
          MapInfo(Map).Name = GetVar(c$, "Mapa" & Map, "Name")
          MapInfo(Map).Music = GetVar(c$, "Mapa" & Map, "MusicNum")
          MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
          MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
          MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
          If val(GetVar(c$, "Mapa" & Map, "Pk")) = 0 Then
                MapInfo(Map).Pk = True
          Else
                MapInfo(Map).Pk = False
          End If
          MapInfo(Map).Restringir = GetVar(c$, "Mapa" & Map, "Restringir")
          MapInfo(Map).BackUp = val(GetVar(c$, "Mapa" & Map, "BackUp"))
          MapInfo(Map).Terreno = GetVar(c$, "Mapa" & Map, "Terreno")
          MapInfo(Map).Zona = GetVar(c$, "Mapa" & Map, "Zona")
          '[Misery_Ezequiel 27/06/05]
          MapInfo(Map).Nivel = GetVar(c$, "Mapa" & Map, "Nivel")
          '[\]Misery_Ezequiel 27/06/05]
          frmCargando.cargar.Value = frmCargando.cargar.Value + 1
          
          DoEvents
Next Map

FrmStat.Visible = False

Exit Sub
man:
    MsgBox ("Error durante la carga de mapas.")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
End Sub

Sub LoadMapData()
'Call LogTarea("Sub LoadMapData")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim npcfile As String

On Error GoTo man

NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.Value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    DoEvents
    
    
    Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
    Seek #1, 1
    
    'inf
    Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
    Seek #2, 1
    
     'map Header
    Get #1, , MapInfo(Map).MapVersion
    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt

    'inf Header
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
        
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            '.dat file
            Get #1, , MapData(Map, X, Y).Blocked
            
            For LoopC = 1 To 4
                Get #1, , MapData(Map, X, Y).Graphic(LoopC)
            Next LoopC
            
            Get #1, , MapData(Map, X, Y).trigger
            Get #1, , TempInt
            
                                
            '.inf file
            Get #2, , MapData(Map, X, Y).TileExit.Map
            Get #2, , MapData(Map, X, Y).TileExit.X
            Get #2, , MapData(Map, X, Y).TileExit.Y
            
            'Get and make NPC
            Get #2, , MapData(Map, X, Y).NpcIndex
            If MapData(Map, X, Y).NpcIndex > 0 Then
                
                If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                Else
                        npcfile = DatPath & "NPCs.dat"
                End If
                
                'Si el npc debe hacer respawn en la pos
                'original la guardamos
                If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                Else
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                End If
                
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                
                Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
            End If

            'Get and make Object
            Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
            Get #2, , MapData(Map, X, Y).OBJInfo.Amount

            'Space holder for future expansion (Objects, ect.
            Get #2, , DummyInt
            Get #2, , DummyInt
        
        Next X
    Next Y

   
    Close #1
    Close #2

  
    MapInfo(Map).Name = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "MusicNum")
    MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    
    If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Terreno")

    MapInfo(Map).Zona = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Zona")
    
    MapInfo(Map).Frio = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Frio"))
    
    MapInfo(Map).Restringir = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Restringir")
    
    MapInfo(Map).BackUp = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "BACKUP"))
    '[Misery_Ezequiel 27/06/05]
    MapInfo(Map).Nivel = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Nivel")
    '[\]Misery_Ezequiel 27/06/05]
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
Next Map


Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
End Sub

'Sub LoadMapData_Nuevo()
'
'
''Call LogTarea("Sub LoadMapData")
'
'If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas."
'
'Dim Map As Integer
'Dim LoopC As Integer
'Dim X As Integer
'Dim Y As Integer
'Dim DummyInt As Integer
'Dim TempInt As Integer
'Dim NpcFile As String
'
'Dim archmap As String, archinf As String
'
'On Error GoTo man
'
'NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
'
'frmCargando.cargar.Min = 0
'frmCargando.cargar.max = NumMaps
'frmCargando.cargar.Value = 0
'
'MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
'
'ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
'ReDim MapInfo(1 To NumMaps) As MapInfo
'
'For Map = 1 To NumMaps
'    DoEvents
'
'    archmap = App.Path & MapPath & "Mapa" & Map & ".map"
'    archinf = App.Path & MapPath & "Mapa" & Map & ".inf"
'
'    Call CargarUnMapa(Map, archmap, archinf)
'
'    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
'Next Map
'
'
'Exit Sub
'
'man:
'    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
'    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
'
'
'End Sub


Sub LoadSini()

Dim Temporal As Long
Dim Temporal1 As Long
Dim LoopC As Integer

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

'Misc
CrcSubKey = val(GetVar(IniPath & "Server.ini", "INIT", "CrcSubKey"))

ServerIp = GetVar(IniPath & "Server.ini", "INIT", "ServerIp")
Temporal = InStr(1, ServerIp, ".")
Temporal1 = (Mid(ServerIp, 1, Temporal - 1) And &H7F) * 16777216
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + Mid(ServerIp, 1, Temporal - 1) * 65536
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + Mid(ServerIp, 1, Temporal - 1) * 256
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))

MixedKey = (Temporal1 + ServerIp) Xor &H65F64B42

Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
CamaraLenta = val(GetVar(IniPath & "Server.ini", "INIT", "CamaraLenta"))
ServerSoloGMs = val(GetVar(IniPath & "server.ini", "init", "ServerSoloGMs"))
UsandoSistemaPadrinos = val(GetVar(IniPath & "Server.ini", "INIT", "UsandoSistemaPadrinos"))
CantidadPorPadrino = val(GetVar(IniPath & "Server.ini", "INIT", "CantidadPorPadrino"))
Antish = val(GetVar(IniPath & "server.ini", "init", "ANTISH"))
'no olvidar
'ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
'ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
'ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
'TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
'TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))

'ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
'ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
'ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
'TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
'TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))

MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))

ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))
EnTesting = val(GetVar(IniPath & "server.ini", "INIT", "Testing"))
EncriptarProtocolosCriticos = val(GetVar(IniPath & "server.ini", "INIT", "Encriptar"))


'If ClientsCommandsQueue <> 0 Then
'        frmMain.CmdExec.Enabled = True
'Else
'        frmMain.CmdExec.Enabled = False
'End If

'Start pos
StartPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.Text = IntervaloSed

IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

'[Misery_Ezequiel 11/07/05]
IntervaloFrioDeNieve = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrioDeNieve"))
'[\]Misery_Ezequiel 11/07/05]

IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

frmMain.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval

frmMain.CmdExec.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec"))
FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))

IntervaloAutoReiniciar = val(GetVar(IniPath & "server.ini", "INTERVALOS", "IntervaloAutoReiniciar"))


'Ressurect pos
ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

#If (UsarQueSocket = 1) Then
'Busqueda eficiente :D
'ReDim Preserve WSAPISockChache(1 To MaxUsers + 10)
'WSAPISockChacheCant = 0
#End If

Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

'[Misery_Ezequiel 10/07/05]
Arghâl.Map = GetVar(DatPath & "Ciudades.dat", "Arghâl", "Mapa")
Arghâl.X = GetVar(DatPath & "Ciudades.dat", "Arghâl", "X")
Arghâl.Y = GetVar(DatPath & "Ciudades.dat", "Arghâl", "Y")
'[\]Misery_Ezequiel 10/07/05]

Call MD5sCarga

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''NUEVA FOMA DE GUARDADO''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo errhandler
Dim OldUserHead As Long

If rs.State = 0 Then
Else
rs.Update
rs.Close
End If

'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If UserList(UserIndex).Clase = "" Or UserList(UserIndex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).Name)
    Exit Sub
End If

If UserList(UserIndex).flags.Mimetizado = 1 Then
   UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
    UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
End If
Dim variable As String
If UserList(UserIndex).flags.Muerto = 1 Then
OldUserHead = UserList(UserIndex).Char.Head
UserList(UserIndex).Char.Head = iCabezaMuerto
Else
variable = ",BodyB='" & str(UserList(UserIndex).Char.Body) & "'"
variable = variable & ",Vot='" & str(UserList(UserIndex).Stats.VotC) & "'"
End If
Dim nromascotas As Integer

'///////////////MASCOTAS//////////////////////////////
nromascotas = 0
Dim i As Integer
For i = 1 To MAXMASCOTAS
 If UserList(UserIndex).MascotasIndex(i) > 0 Then
     If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
     nromascotas = nromascotas + 1
            Select Case nromascotas
            Case "1"
            variable = variable & ",mas1='" & UserList(UserIndex).MascotasType(i) & "'"
            Case "2"
            variable = variable & ",mas2='" & UserList(UserIndex).MascotasType(i) & "'"
            Case "3"
            variable = variable & ",mas3='" & UserList(UserIndex).MascotasType(i) & "'"
           End Select
    End If
End If
Next i
variable = variable & ",NroMascotasB='" & str(nromascotas) & "'"
'////////////MASCOTAS////////////////////////////////////
variable = variable & ",VecesCheat='" & UserList(UserIndex).Stats.Veceshechado & "'"
variable = variable & ",BanrazB='" & UserList(UserIndex).flags.Banrazon & "'"


'ATRIBUTOS
For i = 1 To 5
variable = variable & "," & "AT" & i & "='" & val(UserList(UserIndex).Stats.UserAtributosBackUP(i)) & "'"
Next
'SKILLS
For i = 1 To 21
variable = variable & "," & "SK" & i & "='" & val(UserList(UserIndex).Stats.UserSkills(i)) & "'"
Next
'Objetos de boveda
For i = 1 To 40
variable = variable & "," & "Bobj" & i & "='" & UserList(UserIndex).BancoInvent.Object(i).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(i).Amount & "'"
Next
'Objetos del inventario
For i = 1 To 20
variable = variable & "," & "iOBJ" & i & "='" & UserList(UserIndex).Invent.Object(i).ObjIndex & "-" & UserList(UserIndex).Invent.Object(i).Amount & "-" & UserList(UserIndex).Invent.Object(i).Equipped & "'"
Next

For i = 1 To 35
variable = variable & "," & "H" & i & "='" & UserList(UserIndex).Stats.UserHechizos(i) & "'"
Next

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6

variable = variable & ",PROMEDIOB='" & val(L) & "'"
If UserList(UserIndex).LastIP <> UserList(UserIndex).ip Then
variable = variable & ",LastIPB='" & UserList(UserIndex).ip & "'" & ",LastIPB2='" & UserList(UserIndex).LastIP & "'"
End If

sql = ("UPDATE usuarios SET Fecha='" & Date & "'," & "MuertoB='" & UserList(UserIndex).flags.Muerto & "',EscondidoB='" & val(UserList(UserIndex).flags.Escondido) & "',HambreB='" & val(UserList(UserIndex).flags.Hambre) & "',SedB='" & val(UserList(UserIndex).flags.Sed) & "',DesnudoB='" & val(UserList(UserIndex).flags.Desnudo) & "',banB='" & val(UserList(UserIndex).flags.Ban) & "',NavegandoB='" & val(UserList(UserIndex).flags.Navegando) & "'," & _
"EnvenenadoB='" & val(UserList(UserIndex).flags.Envenenado) & "',ParalizadoB='" & val(UserList(UserIndex).flags.Paralizado) & "',PERTENECEB='" & val(UserList(UserIndex).flags.PertAlCons) & "',PERTENECECAOSB='" & val(UserList(UserIndex).flags.PertAlConsCaos) & "',banB='" & val(UserList(UserIndex).flags.Ban) & "',penab='" & val(UserList(UserIndex).Counters.Pena) & "'," & _
"penasasb='" & UserList(UserIndex).flags.Penasas & "',EjercitoRealB='" & val(UserList(UserIndex).Faccion.ArmadaReal) & "',EjercitoCaosB='" & val(UserList(UserIndex).Faccion.FuerzasCaos) & "',CiudMatadosB='" & val(UserList(UserIndex).Faccion.CiudadanosMatados) & "',CrimMatadosB='" & val(UserList(UserIndex).Faccion.CriminalesMatados) & "',rArCaosB='" & val(UserList(UserIndex).Faccion.RecibioArmaduraCaos) & _
"', rArRealB='" & val(UserList(UserIndex).Faccion.RecibioArmaduraReal) & "',rExRealB='" & val(UserList(UserIndex).Faccion.RecibioExpInicialReal) & "',recCaosB='" & val(UserList(UserIndex).Faccion.RecompensasCaos) & "',recRealB='" & val(UserList(UserIndex).Faccion.RecompensasReal) & "',EsGuildLeaderB='" & val(UserList(UserIndex).GuildInfo.EsGuildLeader) & "'" & _
", EchadasB='" & val(UserList(UserIndex).GuildInfo.Echadas) & "',SolicitudesB='" & val(UserList(UserIndex).GuildInfo.Solicitudes) & "',SolicitudesRechazadasB='" & val(UserList(UserIndex).GuildInfo.SolicitudesRechazadas) & "',VecesFueGuildLeaderB='" & val(UserList(UserIndex).GuildInfo.VecesFueGuildLeader) & "',YaVotoB='" & val(UserList(UserIndex).GuildInfo.YaVoto) & "'" & _
", FundoClanB='" & val(UserList(UserIndex).GuildInfo.FundoClan) & "',GuildNameB='" & UserList(UserIndex).GuildInfo.GuildName & "',ClanFundadoB='" & UserList(UserIndex).GuildInfo.ClanFundado & "',ClanesParticipoB='" & str(UserList(UserIndex).GuildInfo.ClanesParticipo) & "',guildPtsB='" & str(UserList(UserIndex).GuildInfo.GuildPoints) & "'" & _
", EmailB='" & UserList(UserIndex).Email & "',generoB='" & UserList(UserIndex).Genero & "',razaB='" & UserList(UserIndex).Raza & "',HogarB='" & UserList(UserIndex).Hogar & "',claseb='" & UserList(UserIndex).Clase & "'" & _
", PasswordB='" & UserList(UserIndex).Password & "',DescB='" & UserList(UserIndex).Desc & "',HeadingB='" & str(UserList(UserIndex).Char.Heading) & "',OG='" & str(UserList(UserIndex).Stats.OroGanado) & "',OP='" & str(UserList(UserIndex).Stats.OroPerdido) & "'" & _
", RG='" & str(UserList(UserIndex).Stats.RetosGanadoS) & "',RP='" & str(UserList(UserIndex).Stats.RetosPerdidosB) & "',Headb='" & str(UserList(UserIndex).OrigChar.Head) & "',armab='" & str(UserList(UserIndex).Char.WeaponAnim) & "',escudob='" & str(UserList(UserIndex).Char.ShieldAnim) & "'" & _
", Cascob='" & str(UserList(UserIndex).Char.CascoAnim) & "',mapb='" & UserList(UserIndex).Pos.Map & "',yb='" & UserList(UserIndex).Pos.Y & "',xb='" & UserList(UserIndex).Pos.X & "'" & _
", gldb='" & str(UserList(UserIndex).Stats.GLD) & "',bancob='" & str(UserList(UserIndex).Stats.Banco) & "',METB='" & str(UserList(UserIndex).Stats.MET) & "',MaxHPB='" & str(UserList(UserIndex).Stats.MaxHP) & "',MinHPB='" & str(UserList(UserIndex).Stats.MinHP) & "'" & _
", FITB='" & str(UserList(UserIndex).Stats.FIT) & "',MaxStaB='" & str(UserList(UserIndex).Stats.MaxSta) & "',MinSTAB='" & str(UserList(UserIndex).Stats.MinSta) & "',MaxMANb='" & str(UserList(UserIndex).Stats.MaxMAN) & "',MinMANB='" & str(UserList(UserIndex).Stats.MinMAN) & "'" & _
", MaxHITB='" & str(UserList(UserIndex).Stats.MaxHIT) & "',MinHITB='" & str(UserList(UserIndex).Stats.MinHIT) & "',MaxAGUB='" & str(UserList(UserIndex).Stats.MaxAGU) & "',minAGUB='" & str(UserList(UserIndex).Stats.MinAGU) & "',MaxHAMB='" & str(UserList(UserIndex).Stats.MaxHam) & "'" & _
", MinHAMB='" & str(UserList(UserIndex).Stats.MinHam) & "',SkillPtsLibresB='" & str(UserList(UserIndex).Stats.SkillPts) & "',EXPB='" & str(UserList(UserIndex).Stats.Exp) & "',elvb='" & str(UserList(UserIndex).Stats.ELV) & "',ELUB='" & str(UserList(UserIndex).Stats.ELU) & "'" & _
", UserMuertesB='" & val(UserList(UserIndex).Stats.UsuariosMatados) & "',CrimMuertesB='" & val(UserList(UserIndex).Stats.CriminalesMatados) & "',NpcsMuertesB='" & val(UserList(UserIndex).Stats.NPCsMuertos) & "',CantidadItemsB='" & val(UserList(UserIndex).BancoInvent.NroItems) & "',WeaponEqpSlotB='" & str(UserList(UserIndex).Invent.WeaponEqpSlot) & "'" & _
", ArmourEqpSlotB='" & str(UserList(UserIndex).Invent.ArmourEqpSlot) & "',CascoEqpSlotB='" & str(UserList(UserIndex).Invent.CascoEqpSlot) & "',EscudoEqpSlotB='" & str(UserList(UserIndex).Invent.EscudoEqpSlot) & "',BarcoSlotB='" & str(UserList(UserIndex).Invent.BarcoSlot) & "',MunicionSlotB='" & str(UserList(UserIndex).Invent.MunicionEqpSlot) & "'" & _
", HerramientaSlotB='" & str(UserList(UserIndex).Invent.HerramientaEqpSlot) & "',AsesinoB='" & val(UserList(UserIndex).Reputacion.AsesinoRep) & "',BandidoB='" & val(UserList(UserIndex).Reputacion.BandidoRep) & "',BurguesiaB='" & val(UserList(UserIndex).Reputacion.BurguesRep) & "',LadronesB='" & val(UserList(UserIndex).Reputacion.LadronesRep) & "'" & _
", NoblesB='" & val(UserList(UserIndex).Reputacion.NobleRep) & "',PlebeB='" & val(UserList(UserIndex).Reputacion.PlebeRep) & "',rExCaosB='" & val(UserList(UserIndex).Faccion.RecibioExpInicialCaos) & "'" & variable & _
" WHERE NickB='" & UserFile & "'")

 conn.Execute (sql)

Exit Sub
errhandler:
conn.Close
conn.Open constr
Call LogError("Error en SaveUser de " & UserList(UserIndex).Name & Err.Description & " ")

End Sub




Function Criminal(ByVal UserIndex As Integer) As Boolean

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Criminal = (L < 0)

End Function

Sub BackUPnPc(NpcIndex As Integer)

'Call LogTarea("Sub BackUPnPc NpcIndex:" & NpcIndex)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

If NpcNumero > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Inflacion", val(Npclist(NpcIndex).Inflacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))

'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.Def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados))

'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
   Next
End If
End Sub

Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Call LogTarea("Sub CargarNpcBackUp NpcIndex:" & NpcIndex & " NpcNumber:" & NpcNumber)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

Dim npcfile As String

If NpcNumber > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))

Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))

Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
    Next LoopC
End If

Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))

Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile
End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile
End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile
End Sub

'Public Sub CargarUnMapa(Map As Integer, archmap As String, archinf As String)
'Dim Dm As Long
'Dim TM As TileMap, TI As TileInf
'Dim X As Integer, Y As Integer
'Dim LoopC As Integer
'Dim NpcFile As String
'
'    Dm = MAPCargaMapa(archmap, archinf)
'    If Dm = 0 Then
'        Debug.Print "kk " & Map
'    End If
'
'    For Y = YMinMapSize To YMaxMapSize
'        For X = XMinMapSize To XMaxMapSize
'            Call MAPLeeMapa(Dm, TM, TI)
'
'            '.dat file
'            'Get #1, , MapData(Map, X, Y).Blocked
'            MapData(Map, X, Y).Blocked = TM.bloqueado
'
''            For LoopC = 1 To 4
''                'Get #1, , MapData(Map, X, Y).Graphic(LoopC)
''                MapData(Map, X, Y).Graphic(LoopC) = TM.grafs(LoopC)
''                If TM.grafs(LoopC) <> 0 Then
''                    TM.grafs(LoopC) = TM.grafs(LoopC)
''                End If
''            Next LoopC
'            MapData(Map, X, Y).Graphic(1) = TM.grafs1
'            MapData(Map, X, Y).Graphic(2) = TM.grafs2
'            MapData(Map, X, Y).Graphic(3) = TM.grafs3
'            MapData(Map, X, Y).Graphic(4) = TM.grafs4
'
'            'Get #1, , MapData(Map, X, Y).trigger
'            'Get #1, , TempInt
'            MapData(Map, X, Y).trigger = TM.trigger
'            If TM.trigger <> 0 Then
'                TM.trigger = TM.trigger
'            End If
'
'            '.inf file
'            'Get #2, , MapData(Map, X, Y).TileExit.Map
'            'Get #2, , MapData(Map, X, Y).TileExit.X
'            'Get #2, , MapData(Map, X, Y).TileExit.Y
'
'            MapData(Map, X, Y).TileExit.Map = TI.dest_mapa
'            MapData(Map, X, Y).TileExit.X = TI.dest_x
'            MapData(Map, X, Y).TileExit.Y = TI.dest_y
'
'            'Get and make NPC
'            'Get #2, , MapData(Map, X, Y).NpcIndex
'            MapData(Map, X, Y).NpcIndex = TI.npc
'
'            If MapData(Map, X, Y).NpcIndex > 0 Then
'
'                If MapData(Map, X, Y).NpcIndex > 499 Then
'                        NpcFile = DatPath & "NPCs-HOSTILES.dat"
'                Else
'                        NpcFile = DatPath & "NPCs.dat"
'                End If
'
'                'Si el npc debe hacer respawn en la pos
'                'original la guardamos
'                If val(GetVar(NpcFile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
'                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
'                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
'                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
'                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
'                Else
'                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
'                End If
'
'                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
'                Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
'                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
'
'                Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
'            End If
'
'            'Get and make Object
'            'Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
'            'Get #2, , MapData(Map, X, Y).OBJInfo.Amount
'            MapData(Map, X, Y).OBJInfo.ObjIndex = TI.obj_ind
'            MapData(Map, X, Y).OBJInfo.Amount = TI.obj_cant
'
'            'Space holder for future expansion (Objects, ect.
'            'Get #2, , DummyInt
'            'Get #2, , DummyInt
'
'        Next X
'    Next Y
'
'    Call MAPCierraMapa(Dm)
'    ''Close #1
'    ''Close #2
'
'
'    MapInfo(Map).Name = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Name")
'    MapInfo(Map).Music = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "MusicNum")
'    MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
'    MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
'    MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
'
'    If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Pk")) = 0 Then
'        MapInfo(Map).Pk = True
'    Else
'        MapInfo(Map).Pk = False
'    End If
'
'
'    MapInfo(Map).Terreno = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Terreno")
'
'    MapInfo(Map).Zona = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Zona")
'
'    MapInfo(Map).Restringir = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Restringir")
'
'    MapInfo(Map).BackUp = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "BACKUP"))
'
'End Sub

Public Sub CargaApuestas()
Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))
End Sub
