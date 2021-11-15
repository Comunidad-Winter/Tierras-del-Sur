Attribute VB_Name = "Mod_General"
'Argentum Online 0.9.0.9

'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Public bO As Integer
Public bK As Long
Public bRK As Long
Public ss As Boolean
Public gh As Boolean
Public iplst As String
Public banners As String

Public bInvMod As Boolean  'El inventario se modificó?
Public bFogata As Boolean
Public CheatEn As Integer
'[Misery_Ezequiel 10/07/05]
Public bNieva() As Byte ' Array para determinar si
'debemos mostrar la animacion de la nieve
'[\]Misery_Ezequiel 10/07/05]
'**** Used By MP3 Playing. *****
    Public IMC   As IMediaControl
    Dim IBA   As IBasicAudio
    Dim IME   As IMediaEvent
    Dim IMPos As IMediaPosition
Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia
Public Activado As Boolean
Public nuevoc As Boolean
Public clantext1 As String
Public clantext2 As String
Public clantext3 As String

Public clantext5 As String
Private lFrameLimiter As Long
Public clantext4 As String
Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public sHKeys() As String

Const KEYEVENTF_KEYUP = &H2
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Function DirGraficos() As String
DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function SD(ByVal N As Integer) As Integer
'Suma digitos
Dim auxint As Integer
Dim digit As Byte
Dim suma As Integer
auxint = N
Do
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10
Loop While (auxint <> 0)
SD = suma
End Function

Public Function SDM(ByVal N As Integer) As Integer
'Suma digitos cada digito menos dos
Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer
auxint = N
Do
    digit = (auxint Mod 10)
    digit = digit - 1
    suma = suma + digit
    auxint = auxint \ 10
Loop While (auxint <> 0)
SDM = suma
End Function

Public Function Complex(ByVal N As Integer) As Integer
If N Mod 2 <> 0 Then
    Complex = N * SD(N)
Else
    Complex = N * SDM(N)
End If
End Function

Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Sub PlayWaveAPI(File As String)
On Error Resume Next

End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
Randomize Timer
RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "armas.dat"
DoEvents

NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

For loopc = 1 To NumWeaponAnims
    InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:
Versiones(1) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Graficos", "Val"))
Versiones(2) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Wavs", "Val"))
Versiones(3) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Midis", "Val"))
Versiones(4) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Init", "Val"))
Versiones(5) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Mapas", "Val"))
Versiones(6) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "E", "Val"))
Versiones(7) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub
errorH:
MsgBox ("Error cargando versiones")
End Sub

Sub CargarAnimEscudos()
On Error Resume Next
Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "escudos.dat"
DoEvents

NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

For loopc = 1 To NumEscudosAnims
    InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
Next loopc
End Sub

Sub Addtostatus(RichTextBox As RichTextBox, Text As String, red As Byte, green As Byte, blue As Byte, Bold As Byte, italic As Byte)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

frmCargando.Status.SelStart = Len(RichTextBox.Text)
frmCargando.Status.SelLength = 0
frmCargando.Status.SelColor = RGB(red, green, blue)

If Bold Then
    frmCargando.Status.SelBold = True
Else
    frmCargando.Status.SelBold = False
End If

If italic Then
    frmCargando.Status.SelItalic = True
Else
    frmCargando.Status.SelItalic = False
End If

frmCargando.Status.SelText = Chr(13) & Chr(10) & Text

End Sub

Public Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, Optional red As Integer = -1, Optional green As Integer, Optional blue As Integer, Optional Bold As Boolean, Optional italic As Boolean, Optional bCrLf As Boolean)
  Static meguardo As Integer
    With RichTextBox
        If (Len(.Text)) > 3000 Then
        .Text = vbCrLf & right(.Text, Len(.Text) - 500)
        End If
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = IIf(Bold, True, False)
        .SelItalic = IIf(italic, True, False)

        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        RichTextBox.Refresh
    End With
End Sub

Private Function Hex2Dec(ByVal h As String) As Long
Dim I As Long, N As Long, V As Long, C As Long
N = 0
For I = Len(h) To 1 Step -1
    C = Asc(UCase$(Mid$(h, I, 1)))
    If C >= Asc("A") And C <= Asc("F") Then
        V = C - Asc("A") + 10
    ElseIf C >= Asc("0") And C <= Asc("9") Then
        V = C - Asc("0")
    Else
        V = 0
    End If
    N = N + (16 ^ (Len(h) - I)) * V
Next I
Hex2Dec = N
End Function

'Sub AddtoRichTextBox(RichTextBox As RichTextBox, txt As String, Optional RED As Integer = -1, Optional GREEN As Integer, Optional BLUE As Integer, Optional Bold As Boolean, Optional Italic As Boolean, Optional bCrLf As Boolean)
'Dim i As Long
'Dim N As Long
'Dim Tag As String
'Dim t() As String
'Dim Dale As Boolean
'Dim ColorStack As New Collection
'
'With RichTextBox
'
'If (Len(.Text)) > 2000 Then .Text = ""
'.SelStart = Len(.Text)
'.SelLength = 0
'
''If Not IsMissing(Bold) Then .SelBold = IIf(Bold, True, False)
''If Not IsMissing(Italic) Then .SelItalic = IIf(Italic, True, False)
''If Not IsMissing(RED) And Not IsMissing(GREEN) And Not IsMissing(BLUE) Then .SelColor = RGB(RED, GREEN, BLUE)
'.SelBold = IIf(Bold, True, False)
'.SelItalic = IIf(Italic, True, False)
'If Not RED = -1 Then .SelColor = RGB(RED, GREEN, BLUE)
'
'If InStr(1, txt, "<") > 0 Then
'    i = 1
'    Dale = True
'
'    Do While Dale
'        N = InStr(i, txt, "<")
'        If N > 0 Then
'            .SelText = Mid(txt, i, N - i)
'
'            i = N + 1
'            N = InStr(i, txt, ">")
'            If N > 0 Then
'                Tag = Mid(txt, i, N - i)
'                i = N + 1
'                t = Split(Tag, " ")
'
'                If Len(Tag) > 0 Then
'                    Select Case UCase(t(0))
'                    Case "B"
'                        .SelBold = True
'                    Case "/B"
'                        .SelBold = False
'                    Case "K"
'                        .SelItalic = True
'                    Case "/K"
'                        .SelItalic = False
'                    Case "U"
'                        .SelUnderline = True
'                    Case "/U"
'                        .SelUnderline = False
'                    Case "C"
'                        If UBound(t) > 0 Then
'                            ColorStack.Add .SelColor
'                            .SelColor = IIf(Left(t(1), 1) = "#", Hex2Dec(t(1)), Val(t(1)))
'                        End If
'                    Case "/C"
'                        If ColorStack.Count > 0 Then
'                            .SelColor = ColorStack.Item(ColorStack.Count)
'                            ColorStack.Remove ColorStack.Count
'                        End If
'                    End Select
'                End If
'            Else
'                Dale = False
'            End If
'        Else
'            .SelText = Mid(txt, i)
'            Dale = False
'        End If
'    Loop
'    If Not bCrLf Then .SelText = vbCrLf
'Else
'    .SelText = IIf(bCrLf, txt, txt & vbCrLf)
'End If
'
'.Refresh
'
'End With
'
'End Sub

Sub AddtoTextBox(TextBox As TextBox, Text As String)
'******************************************
'Adds text to a text box at the bottom.
'Automatically scrolls to new text.
'******************************************
TextBox.SelStart = Len(TextBox.Text)
TextBox.SelLength = 0
TextBox.SelText = Chr(13) & Chr(10) & Text
End Sub

Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
Dim loopc As Integer

For loopc = 1 To LastChar
    If CharList(loopc).Active = 1 Then
        MapData(CharList(loopc).Pos.X, CharList(loopc).Pos.Y).CharIndex = loopc
    End If
Next loopc
End Sub

Sub SaveGameini()
'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim I As Integer

cad = LCase$(cad)
For I = 1 To Len(cad)
    car = Asc(Mid$(cad, I, 1))
    If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
Next I
AsciiValidos = True
End Function

Function CheckUserData(checkemail As Boolean) As Boolean
'Validamos los datos del user
Dim loopc As Integer
Dim CharAscii As Integer

'If IPdelServidor = frmMain.Socket1.LocalAddress Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If IPdelServidor = "localhost" Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If IPdelServidor = frmMain.Socket1.LocalName Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If IPdelServidor = "" Then
'    MsgBox ("IP del server incorrecto")
'    Exit Function
'End If
'
'If PuertoDelServidor = "" Then
'    MsgBox ("Puerto invalido.")
'    Exit Function
'End If
If checkemail Then
 If UserEmail = "" Then
    MsgBox ("Direccion de email invalida")
    Exit Function
 End If
End If
If UserPassword = "" Then
    MsgBox ("Ingrese un password.")
    Exit Function
End If
For loopc = 1 To Len(UserPassword)
    CharAscii = Asc(Mid$(UserPassword, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Password invalido.")
        Exit Function
    End If
Next loopc
If UserName = "" Then
    MsgBox ("Nombre invalido.")
    Exit Function
End If
If Len(UserName) > 30 Then
    MsgBox ("El nombre debe tener menos de 30 letras.")
    Exit Function
End If
For loopc = 1 To Len(UserName)

    CharAscii = Asc(Mid$(UserName, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Nombre invalido.")
        Exit Function
    End If
Next loopc
CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next
    Dim mifrm As Form
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
'if backspace allow
If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If
'Only allow space,numbers,letters and special characters
If KeyAscii < 32 Or KeyAscii = 44 Then
    LegalCharacter = False
    Exit Function
End If
If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If
'Check for bad special characters in between
If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If
'else everything is cool
LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
'Set Connected
Connected = True
Call SaveGameini
'Unload the connect form
Unload frmConnect
frmMain.Label8.Caption = UserName
'Load main form
frmMain.Visible = True
frmMain.piquete.enabled = True
frmMain.SoundFX.enabled = True
CheatEn = 0
End Sub

Sub CargarTip()
Dim N As Integer
N = RandomNumber(1, UBound(Tips))
If N > UBound(Tips) Then N = UBound(Tips)
frmtip.tip.Caption = Tips(N)
End Sub

Sub MoveNorth()
'[Misery_Ezequiel 05/06/05]
If Comerciando = True Or UserMeditar = True Then Exit Sub
'[\]Misery_Ezequiel 05/06/05]
If Cartel Then Cartel = False
If LegalPos(UserPos.X, UserPos.Y - 1) = True And UserParalizado = False Then
    Call SendData("M" & NORTH)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        Call MoveCharbyHead(UserCharIndex, NORTH)
        Call MoveScreen(NORTH)
        DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> NORTH Then
            Call SendData("CHEA" & NORTH)
    End If
End If
End Sub

Sub MoveEast()
'[Misery_Ezequiel 05/06/05]
If Comerciando = True Or UserMeditar = True Then Exit Sub
'[\]Misery_Ezequiel 05/06/05]
If Cartel Then Cartel = False
If LegalPos(UserPos.X + 1, UserPos.Y) = True And UserParalizado = False Then
    Call SendData("M" & EAST)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        Call MoveCharbyHead(UserCharIndex, EAST)
        Call MoveScreen(EAST)
        Call DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> EAST Then
            Call SendData("CHEA" & EAST)
    End If
End If
End Sub

Sub MoveSouth()
'[Misery_Ezequiel 05/06/05]
If Comerciando = True Or UserMeditar = True Then Exit Sub
'[\]Misery_Ezequiel 05/06/05]
If Cartel Then Cartel = False
If LegalPos(UserPos.X, UserPos.Y + 1) = True And UserParalizado = False Then
    Call SendData("M" & SOUTH)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        MoveCharbyHead UserCharIndex, SOUTH
        MoveScreen SOUTH
        DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> SOUTH Then
            Call SendData("CHEA" & SOUTH)
    End If
End If
End Sub

Sub MoveWest()
'[Misery_Ezequiel 05/06/05]
If Comerciando = True Or UserMeditar = True Then Exit Sub
'[\]Misery_Ezequiel 05/06/05]
If Cartel Then Cartel = False
If LegalPos(UserPos.X - 1, UserPos.Y) = True And UserParalizado = False Then
    Call SendData("M" & WEST)
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            MoveCharbyHead UserCharIndex, WEST
            MoveScreen WEST
            DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> WEST Then
            Call SendData("CHEA" & WEST)
    End If
End If
End Sub

Sub RandomMove()
'[Misery_Ezequiel 05/06/05]
If Comerciando = True Or UserMeditar = True Then Exit Sub
'[\]Misery_Ezequiel 05/06/05]
Dim j As Integer
j = RandomNumber(1, 4)
Select Case j
    Case 1
        Call MoveEast
    Case 2
        Call MoveNorth
    Case 3
        Call MoveWest
    Case 4
        Call MoveSouth
End Select
End Sub

Sub CheckKeys()
On Error Resume Next
'*****************************************************************
'Checks keys and respond
'*****************************************************************
Static KeyTimer As Integer
'Makes sure keys aren't being pressed to fast

If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If

If UserMeditar Then Exit Sub  '[Wizard 03/09/05]=> Esto hay q probarlo tengo mis dudas.
'Don't allow any these keys during movement..
If UserMoving = 0 Then
    If Not UserEstupido Then
            'Move Up
             frmMain.Coord2.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
            If GetKeyState(vbKeyUp) < 0 Then
                If Istrabajando = True Then Call DejarDeTrabajar
                If frmMain.TrainingMacro.enabled Then frmMain.DesactivarMacroHechizos
                Call MoveNorth
                Exit Sub
                  
            End If
            'Move Right
            If GetKeyState(vbKeyRight) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If frmMain.TrainingMacro.enabled Then frmMain.DesactivarMacroHechizos
                If Istrabajando = True Then Call DejarDeTrabajar
                Call MoveEast
                Exit Sub
                frmMain.Coord2.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
            End If
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                If frmMain.TrainingMacro.enabled Then frmMain.DesactivarMacroHechizos
                If Istrabajando = True Then Call DejarDeTrabajar
                Call MoveSouth
                'frmMain.Coord.Caption = Mapa(UserMap)
                Exit Sub
                   frmMain.Coord2.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
            End If
            'Move left
            If GetKeyState(vbKeyLeft) < 0 And GetKeyState(vbKeyShift) >= 0 Then
                If frmMain.TrainingMacro.enabled Then frmMain.DesactivarMacroHechizos
                If Istrabajando = True Then Call DejarDeTrabajar
                Call MoveWest
                'frmMain.Coord.Caption = Mapa(UserMap)
                Exit Sub
                   frmMain.Coord2.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
            End If
    Else
        Dim kp As Boolean
        kp = (GetKeyState(vbKeyUp) < 0) Or _
        GetKeyState(vbKeyRight) < 0 Or _
        GetKeyState(vbKeyDown) < 0 Or _
        GetKeyState(vbKeyLeft) < 0
        If kp Then Call RandomMove
        If frmMain.TrainingMacro.enabled Then frmMain.DesactivarMacroHechizos
        frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"
    End If
End If
End Sub

Sub MoveScreen(Heading As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case Heading
    Case NORTH
        Y = -1
    Case EAST
        X = 1
    Case SOUTH
        Y = 1
    Case WEST
        X = -1
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
Exit Sub
Stop
    '[CODE 001]:MatuX'
        ' Frame checker para el cheat ese
        Select Case FramesPerSecCounter
            Case 18 To 19
                lFrameModLimiter = 60
            Case 17
                lFrameModLimiter = 60
            Case 16
                lFrameModLimiter = 120
            Case 15
                lFrameModLimiter = 240
            Case 14
                lFrameModLimiter = 480
            Case 15
                lFrameModLimiter = 960
            Case 14
                lFrameModLimiter = 1920
            Case 13
                lFrameModLimiter = 3840
            Case 12
            Case 11
            Case 10
            Case 9
            Case 8
            Case 7
            Case 6
            Case 5
            Case 4
            Case 3
            Case 2
            Case 1
                lFrameModLimiter = 60 * 256
            Case 0
        End Select
    '[END]'

    Call DoFogataFx
End If
End Sub

Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active And loopc < UBound(CharList)
    loopc = loopc + 1
Loop
NextOpenChar = loopc
End Function

Public Function DirMapas() As String
DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Sub SwitchMap(Map As Integer)
Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim TempInt As Integer
      
Open DirMapas & "Mapa" & Map & ".map" For Binary As #1
Seek #1, 1
        
'map Header
Get #1, , MapInfo.MapVersion
Get #1, , MiCabecera
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'Load arrays
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        '.dat file
        Get #1, , MapData(X, Y).Blocked
        For loopc = 1 To 4
            Get #1, , MapData(X, Y).Graphic(loopc).GrhIndex
            'Set up GRH
            If MapData(X, Y).Graphic(loopc).GrhIndex > 0 Then
                InitGrh MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).GrhIndex
            End If
        Next loopc
        Get #1, , MapData(X, Y).Trigger
        Get #1, , TempInt
        'Erase NPCs
        If MapData(X, Y).CharIndex > 0 Then
            Call EraseChar(MapData(X, Y).CharIndex)
        End If
        
        'Erase OBJs
        MapData(X, Y).ObjGrh.GrhIndex = 0

    Next X
Next Y
Close #1
MapInfo.Name = ""
MapInfo.Music = ""
CurMap = Map
End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
Dim I As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0
For I = 1 To Len(Text)
    CurChar = Mid(Text, I, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = I
    End If
Next I
FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If
End Function

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean
If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Sub WriteClientVer()
Dim hFile As Integer
    
hFile = FreeFile()
Open App.Path & "\init\Ver.bin" For Binary Access Write As #hFile
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)
Put #hFile, , CInt(App.Major)
Put #hFile, , CInt(App.Minor)
Put #hFile, , CInt(App.Revision)
Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
Dim I As Integer
For I = 1 To UBound(ServersLst)
    If ServersLst(I).Ip = Ip Then
        IsIp = True
        Exit Function
    End If
Next I
End Function

Public Sub InitServersList(ByVal Lst As String)
On Error Resume Next
Dim NumServers As Integer
Dim I As Integer, Cont As Integer
I = 1

Do While (ReadField(I, RawServersList, Asc(";")) <> "")
    I = I + 1
    Cont = Cont + 1
Loop

ReDim ServersLst(1 To Cont) As tServerInfo

For I = 1 To Cont
    Dim cur$
    cur$ = ReadField(I, RawServersList, Asc(";"))
    ServersLst(I).Ip = ReadField(1, cur$, Asc(":"))
    ServersLst(I).Puerto = ReadField(2, cur$, Asc(":"))
    ServersLst(I).desc = ReadField(4, cur$, Asc(":"))
    ServersLst(I).PassRecPort = ReadField(3, cur$, Asc(":"))
Next I
CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer
If CurServer <> 0 Then
    CurServerPasRecPort = ServersLst(CurServer).PassRecPort
Else
    CurServerPasRecPort = CInt(frmConnect.PortTxt)
End If
End Function

Public Function CurServerIp() As String
If CurServer <> 0 Then
    CurServerIp = ServersLst(CurServer).Ip
Else
    CurServerIp = frmConnect.IPTxt
End If
End Function

Public Function CurServerPort() As Integer
If CurServer <> 0 Then
    CurServerPort = ServersLst(CurServer).Puerto
Else
    CurServerPort = CInt(frmConnect.PortTxt)
End If
End Function

Sub Main()
On Error Resume Next
ChDir App.Path

Call WriteClientVer
Call LeerLineaComandos


DoEvents
If App.PrevInstance Then
    Call MsgBox("Tierras Del Sur ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    End
End If


Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Integer

ChDrive App.Path
ChDir App.Path

'Obtengo mi MD5 hash
Dim fMD5HushYo As String * 32
fMD5HushYo = MD5File(App.Path & "\" & App.EXEName & ".exe")
'fMD5HushYo = "669423534a1063dcf5e0ba9d70983a2c"
MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 53)

'Cargamos el archivo de configuracion inicial
If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
    Config_Inicio = LeerGameIni()
End If

If FileExist(App.Path & "\init\ao.dat", vbNormal) Then
    Open App.Path & "\init\ao.dat" For Binary As #53
        Get #53, , RenderMod
    Close #53

    Musica = IIf(RenderMod.bNoMusic = 1, 1, 0)
    Fx = IIf(RenderMod.bNoSound = 1, 1, 0)
    
    'RenderMod.iImageSize = 0
    Select Case RenderMod.iImageSize
        Case 4
            RenderMod.iImageSize = 0
        Case 3
            RenderMod.iImageSize = 1
        Case 2
            RenderMod.iImageSize = 2
        Case 1
            RenderMod.iImageSize = 3
        Case 0
            RenderMod.iImageSize = 4
    End Select
End If

tipf = Config_Inicio.tip

frmCargando.Show
frmCargando.Refresh

UserParalizado = False
'[Wizard] Anti doble aO.
frmCargando.Label1 = "Verificando si hay otro cliente abierto..."
frmCargando.Caption = "TemporalCaption"
frmCargando.Caption = "Tierras Del Sur   "
'------------------------------------------

frmConnect.Version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
frmCargando.Label1 = "Buscando servidores...."

frmMain.Inet1.URL = "http://www.aotds.com.ar/iplist.txt"
RawServersList = frmMain.Inet1.OpenURL

'If RawServersList = "" Then
'    frmMain.Inet1.URL = "http://www.argentum-online.com.ar/admin/iplist2.txt"
'End If

#If UsarWrench = 1 Then
frmMain.Socket1.Startup
#Else
#End If
Debug.Print RawServersList
If RawServersList = "" Then
    ServersRecibidos = False
    ReDim ServersLst(1)
Else
    If Mid(RawServersList, 1, 9) = "<!DOCTYPE" Then
    frmMain.Inet1.URL = "http://200.32.10.47/iplist.txt"
    RawServersList = frmMain.Inet1.OpenURL
    End If
    ServersRecibidos = True
End If

Call InitServersList(RawServersList)

'IPdelServidor =
'PuertoDelServidor = 7666

frmCargando.Label1 = "Encontrado"
frmCargando.Label1 = "Iniciando constantes..."

ReDim Ciudades(1 To NUMCIUDADES) As String
Ciudades(1) = "Ullathorpe"
Ciudades(2) = "Nix"
Ciudades(3) = "Banderbill"
'[Misery_Ezequiel 10/07/05]
Ciudades(4) = "Arghâl"
'[\]Misery_Ezequiel 10/07/05]

ReDim CityDesc(1 To NUMCIUDADES) As String
CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
'********eLwE 15/05/05********
ReDim Mapa(1 To NumMapas) As String
Mapa(1) = "Ullathorpe"
Mapa(2) = "Sur Ullathorpe"
Mapa(3) = "Bosque"
Mapa(4) = "Suburbios Infectados"
Mapa(5) = "Norte Ullathorpe"
Mapa(6) = "Bosque de Duendes"
Mapa(7) = "Bosque"
Mapa(8) = "Bosque"
Mapa(9) = "Bosque"
Mapa(10) = "Bosque de Tigres"
Mapa(11) = "Este de Ullathorpe"
Mapa(12) = "Bosque"
Mapa(13) = "Bosque Infectado"
Mapa(14) = "Bosque"
Mapa(15) = "Desierto Renkil"
Mapa(16) = "Oasis"
Mapa(17) = "Desierto Rinkel"
Mapa(18) = "Bosque"
Mapa(19) = "Bosque"
Mapa(20) = "Muelle del Desierto"
Mapa(21) = "Desierto Rinkel"
Mapa(22) = "Bosque"
Mapa(23) = "Bosque"
Mapa(24) = "Bosque Fantasmal"
Mapa(25) = "Bosque Místico"
Mapa(26) = "Bosque"
Mapa(27) = "Bosque"
Mapa(28) = "Bosque"
Mapa(29) = "Bosque"
Mapa(30) = "Bosque"
Mapa(31) = "Bosque"
Mapa(32) = "Bosque"
Mapa(33) = "Minas Thyr"
Mapa(34) = "Nix"
Mapa(35) = "Mansión de Nix"
Mapa(36) = "Fuerte Orco"
Mapa(37) = "Newbie Dungeon"
Mapa(38) = "Bosque Dorck"
Mapa(39) = "Bosque Dorck"
Mapa(40) = "Catacumbas de Ullathorpe"
Mapa(41) = "Catacumbas"
Mapa(42) = "Catacumbas"
Mapa(43) = "Catacumbas"
Mapa(44) = "Catacumbas"
Mapa(45) = "Catacumbas de Nix"
Mapa(46) = "Bosque Dorck Sur"
Mapa(47) = "Isla Veleta"
Mapa(48) = "Deadly Room"
Mapa(49) = "Torneos"
Mapa(50) = "Minas Rapajik"
Mapa(51) = "Minas Rapajik"
Mapa(52) = "Minas Rapajik"
Mapa(53) = "Bosque"
Mapa(54) = "Bosque de Asesinos"
Mapa(55) = "Bosque de Bandidos"
Mapa(56) = "Bosque"
Mapa(57) = "Sur Banderbill"
Mapa(58) = "Entrada a Banderbill"
Mapa(59) = "Banderbill Central"
Mapa(60) = "Banderbill Norte"
Mapa(61) = "Muelles de Banderbill"
Mapa(62) = "Lindos"
Mapa(63) = "Abadía Lindos Sur"
Mapa(64) = "Lindos Este"
Mapa(65) = "Banderbill Oeste"
Mapa(66) = "Prisión de Bandebill"
Mapa(67) = "Bosque de Osos"
Mapa(68) = "Bosque de Jabalíes"
Mapa(69) = "Bosque de Arañas"
Mapa(70) = "Bosque de Goblins"
Mapa(71) = "Bosque"
Mapa(72) = "Bosque"
Mapa(73) = "Bosque"
Mapa(74) = "Bosque"
Mapa(75) = "Bosque de Tortugas"
Mapa(76) = "Entrada Dungeon Marabel"
Mapa(77) = "Dungeon Vespar"
Mapa(78) = "Mar Oeste Nix"
Mapa(79) = "Mar"
Mapa(80) = "Mar"
Mapa(81) = "Mar"
Mapa(82) = "Mar"
Mapa(83) = "Mar"
Mapa(84) = "Mar"
Mapa(85) = "Mar"
Mapa(86) = "Mar"
Mapa(87) = "Mar Sur de Nix"
Mapa(88) = "Mar"
Mapa(89) = "Mar"
Mapa(90) = "Mar"
Mapa(91) = "Mar"
Mapa(92) = "Mar"
Mapa(93) = "Mar"
Mapa(94) = "Mar"
Mapa(95) = "Mar"
Mapa(96) = "Mar"
Mapa(97) = "Mar"
Mapa(98) = "Isla Dungeon Veriil"
Mapa(99) = "Mar"
Mapa(100) = "Mar"
Mapa(101) = "Mar"
Mapa(102) = "Mar"
Mapa(103) = "Mar"
Mapa(104) = "Mar"
Mapa(105) = "Mar"
Mapa(106) = "Mar" '[Wizard 03/09/05] Arreglado el error; en vez de 106 decia 103.
Mapa(107) = "Mar"
Mapa(108) = "Mar de Leviatanes"
Mapa(109) = "Mar"
Mapa(110) = "Isla de Morgolock"
Mapa(111) = "Isla Esperanza"
Mapa(112) = "Isla Esperanza"
Mapa(113) = "Isla Esperanza"
Mapa(114) = "Isla Esperanza"
Mapa(115) = "Dungeon Marabel"
Mapa(116) = "Dungeon Marabel"
Mapa(117) = "Mar"
Mapa(118) = "Mar"
Mapa(119) = "Mar"
Mapa(120) = "Mar"
Mapa(121) = "Mar de Calamares"
Mapa(122) = "Mar"
Mapa(123) = "Mar"
Mapa(124) = "Mar"
Mapa(125) = "Mar"
Mapa(126) = "Mar"
Mapa(127) = "Mar"
Mapa(128) = "Mar"
Mapa(129) = "Mar"
Mapa(130) = "Mar"
Mapa(131) = "Mar"
Mapa(132) = "Mar"
Mapa(133) = "Mar"
Mapa(134) = "Mar"
Mapa(135) = "Mar"
Mapa(136) = "Mar de Calamares"
Mapa(137) = "Mar"
Mapa(138) = "Mar"
Mapa(139) = "Isla de la Muerte"
Mapa(140) = "Entrada Veriil"
Mapa(141) = "Pasillo Principal"
Mapa(142) = "Cercanías del Fuerte"
Mapa(143) = "Pasillos Veriil"
Mapa(144) = "Cercanías a las Minas"
Mapa(145) = "Minas Veriil"
Mapa(146) = "El Fuerte"
Mapa(147) = "Mar"
Mapa(148) = "Mar"
Mapa(149) = "Mar"
Mapa(150) = "Puerto de Arghâl"
Mapa(151) = "Arghâl"
Mapa(152) = "Mar"
Mapa(153) = "Mar"
Mapa(154) = "Mar"
Mapa(155) = "Mar"
Mapa(156) = "Arghâl Oeste"
Mapa(157) = "Bosques de Banderbill"
Mapa(158) = "Bosques de Banderbill"
Mapa(159) = "Bosque"
Mapa(160) = "Bosque de Tortugas"
Mapa(161) = "Bosque de Osos"
Mapa(162) = "Mar Sur de Veleta"
Mapa(163) = "Dungeon Speculum"
Mapa(164) = "Dungeon Speculum"
Mapa(165) = "Dungeon Speculum"
Mapa(166) = "Dungeon Dragon"
Mapa(167) = "Dungeon Newbie"
Mapa(168) = "Dungeon Newbie"
Mapa(169) = "Polo Norte"
Mapa(170) = "Polo Norte"
Mapa(171) = "Polo Norte"
Mapa(172) = "Dungeon Speculum"
Mapa(173) = "Nueva Esperanza"
Mapa(174) = "Dungeon Speculum"
Mapa(175) = "Dungeon Magma"
Mapa(176) = "Retos"
Mapa(177) = "Ciudad Oscura"
Mapa(178) = "Mar"
Mapa(179) = "Mar"
Mapa(180) = "Mar"
Mapa(181) = "Mar"
Mapa(182) = "Mar"
Mapa(183) = "Isla vespar"
Mapa(184) = "Mar"

'*******eLwE 15/05/05********
'[Nacho]Mensajes!!!!!!!
ReDim Mensaje(1 To NUMMENSAJES) As String
'[Misery_Ezequiel 26/06/05] And Ariel
Mensaje(1) = "!No has logrado esconderte!~65~190~156~0~0"
Mensaje(2) = "Ya estás oculto.~65~190~156~0~0"
Mensaje(3) = "¡¡Estás muerto!!.~65~190~156~0~0"
Mensaje(4) = "Primero tienes que seleccionar un personaje, has click izquierdo sobre él.~65~190~156~0~0"
Mensaje(5) = "Estás demasiado lejos.~65~190~156~0~0"
Mensaje(6) = "Estás muy lejos para disparar.~255~0~0~1~0"
Mensaje(7) = "Estás demasiado lejos del vendedor.~65~190~156~0~0"
Mensaje(8) = "El sacerdote no puede curarte debido a que estás demasiado lejos.~65~190~156~0~0"
Mensaje(9) = "El sacerdote no puede resucitarte debido a que estás demasiado lejos.~65~190~156~0~0"
Mensaje(10) = "¡¡Estás muriendo de frío, abrígate o morirás!!.~65~190~156~0~0"
Mensaje(11) = "Estás muy cansado para luchar.~65~190~156~0~0"
Mensaje(12) = "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro escribiendo /SEG.~255~255~255~1~0"
Mensaje(13) = "Aún no es período de elecciones.~255~255~255~1~0"
Mensaje(14) = "Ya has votado!!! Solo se permite un voto por miembro.~255~255~255~1~0"
Mensaje(15) = "No hay ningún miembro con ese nombre.~255~255~255~1~0"
Mensaje(16) = "Tu voto ha sido contabilizado.~255~255~255~1~0"
Mensaje(17) = "¡¡Has sido curado!!.~65~190~156~0~0"
Mensaje(18) = "Las elecciones han finalizado!!.~255~255~255~1~0"
Mensaje(19) = "¡¡¡Has ganado las elecciones, felicitaciones!!!.~255~255~255~1~0"
Mensaje(20) = "La puerta esta cerrada con llave.~65~190~156~0~0"
Mensaje(21) = "No puedes cargar mas objetos.~65~190~156~0~0"
Mensaje(22) = "¡¡Has muerto de frío!!.~65~190~156~0~0"
Mensaje(23) = "Has vuelto a ser visible.~65~190~156~0~0"
Mensaje(24) = "Te sentís menos cansado.~65~190~156~0~0"
Mensaje(25) = "Has matado la criatura!~255~0~0~1~0"
Mensaje(26) = "Estás muerto!! Solo puedes usar ítems cuando estás vivo.~65~190~156~0~0"
Mensaje(27) = "Ya estás comerciando.~65~190~156~0~0"
Mensaje(28) = "Has prendido la fogata.~65~190~156~0~0"
Mensaje(29) = "No has podido hacer fuego.~65~190~156~0~0"
Mensaje(30) = "Servidor> Iniciando WorldSave.~0~185~0~0~0"
Mensaje(31) = "Servidor> WorldSave ha concluido.~0~185~0~0~0"
Mensaje(32) = "Has sido liberado!~65~190~156~0~0"
Mensaje(33) = "Para surcar los mares debes ser nivel 25 o superior.~65~190~156~0~0"
Mensaje(34) = "No comercio objetos para newbies.~65~190~156~0~0"
Mensaje(35) = "El npc no está interesado en comprar ese objeto.~32~51~223~1~1"
Mensaje(36) = "El npc no puede cargar tantos objetos.~65~190~156~0~0"
Mensaje(37) = "No tienes suficiente dinero.~65~190~156~0~0"
Mensaje(38) = "Mapa exclusivo para newbies.~65~190~156~0~0"
Mensaje(39) = "Estás envenenado, si no te curas morirás.~0~255~0~0~0"
Mensaje(40) = "Has sanado.~65~190~156~0~0"
Mensaje(41) = "Gracias por jugar Tierras Del Sur.~65~190~156~0~0"
Mensaje(42) = "Levas 15 segundos bloqueando el ítem, muévete o serás desconectado.~32~51~223~1~1"
Mensaje(43) = "Servidor> Grabando Personajes.~0~185~0~0~0"
Mensaje(44) = "Servidor> Personajes Grabados.~0~185~0~0~0"
Mensaje(45) = "Solo las clases mágicas conocen el arte de la meditación.~65~190~156~0~0"
Mensaje(46) = "¡¡La criatura te ha envenenado!!~255~0~0~1~0"
Mensaje(47) = "¡Has subido de nivel!~65~190~156~0~0"
Mensaje(48) = "¡Has ganado 50 puntos de experiencia!~255~0~0~1~0"
Mensaje(49) = "Pierdes el control de tus mascotas.~65~190~156~0~0"
Mensaje(50) = "Personaje Inexistente.~65~190~156~0~0"
Mensaje(51) = "Comercio cancelado. El otro usuario se ha desconectado.~255~255~255~0~0"
Mensaje(52) = "Servidor> Por favor espera algunos segundos, WorldSave está ejecutándose.~0~185~0~0~0"
Mensaje(53) = "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.~0~185~0~0~0"
Mensaje(54) = "Mensaje del día:~65~190~156~0~0"
Mensaje(55) = "Comentario salvado...~65~190~156~0~0"
Mensaje(56) = "Usuario offline.~65~190~156~0~0"
Mensaje(57) = "Sin prontuario..~65~190~156~0~0"
Mensaje(58) = "Utilice /carcel nick@motivo@tiempo~65~190~156~0~0"
Mensaje(59) = "El usuario no está online.~255~255~255~0~0"
Mensaje(60) = "No puedes encarcelar a administradores.~65~190~156~0~0"
Mensaje(61) = "No puedes encarcelar por mas de 60 minutos.~65~190~156~0~0"
Mensaje(62) = "Usuario offline, Buscando en Charfile.~65~190~156~0~0"
Mensaje(63) = "Usuario offline. Leyendo Charfile...255~255~255~0~0"
Mensaje(64) = "Usuario offline. Leyendo charfile...~255~255~255~0~0"
Mensaje(65) = "No hay GMs Online.~65~190~156~0~0"
Mensaje(66) = "Solo se permite perdonar newbies.~65~190~156~0~0"
Mensaje(67) = "No puedes echar a alguien con jerarquía mayor a la tuya.~65~190~156~0~0"
Mensaje(68) = "Estás loco?? como vas a matar un GM!!!! :@~65~190~156~0~0"
Mensaje(69) = "No está online.~65~190~156~0~0"
Mensaje(70) = "El usuario no está online.~255~255~255~1~0"
Mensaje(71) = "No puedes banear a al alguien de mayor jerarquía..~65~190~156~0~0"
Mensaje(72) = "El personaje ya se encuentra baneado.~65~190~156~0~0"
Mensaje(73) = "Charfile inexistente (no use +)~65~190~156~0~0"
Mensaje(74) = "No tienes los privilegios necesarios.~65~190~156~0~0"
Mensaje(75) = "No hay ningún personaje con ese Niké.~65~190~156~0~0"
Mensaje(76) = "Personaje Offline.~65~190~156~0~0"
Mensaje(77) = "Usuario offline, echando del consejo.~65~190~156~0~0"
Mensaje(78) = "Hay un objeto en el piso en ese lugar.~65~190~156~0~0"
Mensaje(79) = "Servidor habilitado para todos.~0~185~0~0~0"
Mensaje(80) = "Servidor restringido a administradores.~0~185~0~0~0"
Mensaje(81) = "Usar /APASS <pjsinpass>@<pjconpass>~65~190~156~0~0"
Mensaje(82) = "Usar /APASS <pjsinpass> <pjconpass>~65~190~156~0~0"
Mensaje(83) = "Usar /AEMAIL <pj>-<nuevomail>~65~190~156~0~0"
Mensaje(84) = "El usuario esta online, no se puede si está online.~65~190~156~0~0"
Mensaje(85) = "Usar: /ANAME origen@destino~65~190~156~0~0"
Mensaje(86) = "El Personaje está online, debe salir para el cambio.~32~51~223~1~1"
Mensaje(87) = "Transferencia exitosa.~65~190~156~0~0"
Mensaje(88) = "El nick solicitado ya existe.~65~190~156~0~0"
Mensaje(89) = "No está permitido utilizar valores mayores a 95000. Su comando ha quedado en los logs del juego.~65~190~156~0~0"
Mensaje(90) = "No está permitido utilizar valores mayores a mucho. Su comando ha quedado en los logs del juego.~65~190~156~0~0"
Mensaje(91) = "Comando no permitido.~65~190~156~0~0"
Mensaje(92) = "Los datos estan en BYTES.~65~190~156~0~0"
Mensaje(93) = "¡Te has escondido entre las sombras!~65~190~156~0~0"
Mensaje(94) = "No tienes suficientes conocimientos para usar este barco.~65~190~156~0~0"
Mensaje(95) = "No tienes conocimientos de minería suficientes para trabajar este mineral.~65~190~156~0~0"
Mensaje(96) = "No tienes suficientes madera.~65~190~156~0~0"
Mensaje(97) = "No tienes suficientes lingotes de hierro.~65~190~156~0~0"
Mensaje(98) = "No tienes suficientes lingotes de plata.~65~190~156~0~0"
Mensaje(99) = "No tienes suficientes lingotes de oro.~65~190~156~0~0"
Mensaje(100) = "Has construido el arma!.~65~190~156~0~0"
Mensaje(101) = "Has construido el escudo!.~65~190~156~0~0"
Mensaje(102) = "Has construido el casco!.~65~190~156~0~0"
Mensaje(103) = "Has construido la armadura!.~65~190~156~0~0"
Mensaje(104) = "Has construido el objeto!~65~190~156~0~0"
Mensaje(105) = "No tienes suficientes minerales para hacer un lingote.~65~190~156~0~0"
Mensaje(106) = "Has obtenido lingotes!!!~65~190~156~0~0"
Mensaje(107) = "Los minerales no eran de buena calidad, no has logrado hacer un lingote.~65~190~156~0~0"
Mensaje(108) = "La criatura ya te ha aceptado como su amo.~65~190~156~0~0"
Mensaje(109) = "La criatura ya tiene amo.~65~190~156~0~0"
Mensaje(110) = "La criatura te ha aceptado como su amo.~65~190~156~0~0"
Mensaje(111) = "No has logrado domar la criatura.~65~190~156~0~0"
Mensaje(112) = "No puedes controlar mas criaturas.~65~190~156~0~0"
Mensaje(113) = "Necesitas por lo menos tres troncos para hacer una fogata.~65~190~156~0~0"
Mensaje(114) = "Has hecho una fogata.~65~190~156~0~0"
Mensaje(115) = "No has podido hacer la fogata.~65~190~156~0~0"
Mensaje(116) = "¡Has pescado un lindo pez!~65~190~156~0~0"
Mensaje(117) = "¡No has pescado nada!~65~190~156~0~0"
Mensaje(118) = "¡Has pescado algunos peces!~65~190~156~0~0"
Mensaje(119) = "¡No has pescado nada!~65~190~156~0~0"
Mensaje(120) = "Debes quitar el seguro para robar.~255~200~200~1~0"
Mensaje(121) = "¡No has logrado robar nada!~65~190~156~0~0"
Mensaje(122) = "No has logrado robar objetos.~65~190~156~0~0"
Mensaje(123) = "¡No has logrado apuñalar a tu enemigo!~255~0~0~1~0"
Mensaje(124) = "¡Has conseguido algo de leña!~65~190~156~0~0"
Mensaje(125) = "¡No has obtenido leña!~65~190~156~0~0"
Mensaje(126) = "¡Has extraído algunos minerales!~65~190~156~0~0"
Mensaje(127) = "¡No has conseguido nada!~65~190~156~0~0"
Mensaje(128) = "Has terminado de meditar.~65~190~156~0~0"
Mensaje(129) = "Comercio cancelado por el otro usuario.~255~255~255~0~0"
Mensaje(130) = "Has terminado de descansar.~65~190~156~0~0"
Mensaje(131) = "Estas obstruyendo la vía pública, muévete o serás encarcelado!!!~65~190~156~0~0"
Mensaje(132) = "No tienes espacio para mas hechizos.~65~190~156~0~0"
Mensaje(133) = "Ya tienes ese hechizo.~65~190~156~0~0"
Mensaje(134) = "No tienes la suficiente energía para lanzar este hechizo.~65~190~156~0~0"
Mensaje(135) = "No tienes suficientes puntos de magia para lanzar este hechizo.~65~190~156~0~0"
Mensaje(136) = "No tienes suficiente mana.~65~190~156~0~0"
Mensaje(137) = "No puedes lanzar hechizos porque estás muerto.~65~190~156~0~0"
Mensaje(138) = "Este hechizo actúa solo sobre usuarios.~65~190~156~0~0"
Mensaje(139) = "Este hechizo solo afecta a los npcs.~65~190~156~0~0"
Mensaje(140) = "Target invalido.~65~190~156~0~0"
Mensaje(141) = "NADA~65~190~156~0~0"
Mensaje(142) = "NADA~65~190~156~0~0"
Mensaje(143) = "¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!.~65~190~156~0~0"
Mensaje(144) = "No podes atacar a ese npc.~65~190~156~0~0"
Mensaje(145) = "No puedes atacarte a vos mismo.~255~0~0~1~0"
Mensaje(146) = "Éste hechizo solo afecta NPCs que tengan amo.~32~51~223~1~1"
Mensaje(147) = "NADA~65~190~156~0~0"
Mensaje(148) = "No podes atacarte a vos mismo.~255~0~0~1~0"
Mensaje(149) = "No puedes mover el hechizo en esa dirección.~65~190~156~0~0"
Mensaje(150) = "¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!~255~255~255~0~0"
Mensaje(151) = "¡¡ATENCION!! ¡ACABAS DE TIRAR TU ARMADURA FACCIONARIA!~255~255~255~0~0"
Mensaje(152) = "No hay espacio en el piso.~65~190~156~0~0"
Mensaje(153) = "No podes cargar mas objetos.~255~0~0~1~0"
Mensaje(154) = "No puedo cargar mas objetos.~65~190~156~0~0"
Mensaje(155) = "No hay nada aquí.~65~190~156~0~0"
Mensaje(156) = "Solo los newbies pueden usar este objeto.~65~190~156~0~0"
Mensaje(157) = "Tu clase no puede usar este objeto.~65~190~156~0~0"
Mensaje(158) = "Tu clase, genero o raza no puede usar este objeto.~65~190~156~0~0"
Mensaje(159) = "¡¡Debes esperar unos momentos para tomar otra poción!!~65~190~156~0~0"
Mensaje(160) = "Te has curado del envenenamiento.~65~190~156~0~0"
Mensaje(161) = "Sientes un gran mareo y pierdes el conocimiento.~255~0~0~1~0"
Mensaje(162) = "Has abierto la puerta.~65~190~156~0~0"
Mensaje(163) = "La llave no sirve.~65~190~156~0~0"
Mensaje(164) = "Has cerrado la puerta con llave.~65~190~156~0~0"
Mensaje(165) = "No está cerrada.~65~190~156~0~0"
Mensaje(166) = "No hay agua allí.~65~190~156~0~0"
Mensaje(167) = "Estás muy cansado.~65~190~156~0~0"
Mensaje(168) = "Antes de usar la herramienta deberías equipártela.~65~190~156~0~0"
Mensaje(169) = "Estás demasiado hambriento y sediento.~65~190~156~0~0"
Mensaje(170) = "¡Debes aproximarte al agua para usar el barco!~65~190~156~0~0"
Mensaje(171) = "Debes sacar el seguro antes de poder atacar una mascota de un ciudadano.~32~51~223~1~1"
Mensaje(172) = "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos y sus macotas.~32~51~223~1~1"
Mensaje(173) = "No podes atacar mascotas en zonas seguras.~255~0~0~1~0"
Mensaje(174) = "No podes atacar a este NPC.~255~0~0~1~0"
Mensaje(175) = "No podes atacar a un espíritu.~65~190~156~0~0"
Mensaje(176) = "Esta es una zona segura, aquí no podes atacar otros usuarios.~32~51~223~1~1"
Mensaje(177) = "No podes pelear aquí.~32~51~223~1~1"
Mensaje(178) = "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.~32~51~223~1~1"
Mensaje(179) = "NADA.~32~51~223~1~1"
Mensaje(180) = "Los miembros de la Legión oscura tienen prohibido atacarse entre sí.~32~51~223~1~1"
Mensaje(181) = "¡¡No podes atacar a los administradores del juego!!32~51~223~1~1" 'marche. esto lo saque
Mensaje(182) = "Has sido expulsado de las tropas reales!!!.~255~0~0~1~0"
Mensaje(183) = "Has sido expulsado de la Legión oscura!!!.~255~0~0~1~0"
Mensaje(184) = "No estás en guerra con el clan.~255~255~255~1~0"
Mensaje(185) = "Ya estás en paz con el clan.~255~255~255~1~0"
Mensaje(186) = "La propuesta de paz ha sido entregada.~255~255~255~1~0"
Mensaje(187) = "Ya has enviado una propuesta de paz.~255~255~255~1~0"
Mensaje(188) = "¡No pueden expulsar al líder del clan!~255~255~255~1~0"
Mensaje(189) = "Has sido expulsado del clan.~255~255~255~1~0"
Mensaje(190) = "NADA~255~255~255~1~0"
Mensaje(191) = "Tu solicitud ha sido rechazada.~255~255~255~1~0"
Mensaje(192) = "No podes aceptar esta solicitud, el personaje es líder de otro clan.~255~255~255~1~0"
Mensaje(193) = "Felicitaciones, tu solicitud ha sido aceptada.~255~255~255~1~0"
Mensaje(194) = "Solo podes aceptar solicitudes cuando el solicitante esta ONLINE.~255~255~255~1~0"
Mensaje(195) = "Los newbies no pueden conformar clanes.~255~255~255~1~0"
Mensaje(196) = "La solicitud fue recibida por el líder del clan, ahora debes esperar la respuesta.~255~255~255~1~0"
Mensaje(197) = "Tu solicitud ya fue recibida por el líder del clan, ahora debes esperar la respuesta.~255~255~255~1~0"
Mensaje(198) = "La dirección de la Web ha sido actualizada.~255~255~255~1~0"
Mensaje(199) = "Estás en guerra con éste clan, antes debes firmar la paz.~255~255~255~1~0"
Mensaje(200) = "Ya estás aliado con éste clan.~255~255~255~1~0"
Mensaje(201) = "Hoy es la votación para elegir un nuevo líder para el clan!!.~255~255~255~1~0"
Mensaje(202) = "La elección durará 24 horas, se puede votar a cualquier miembro del clan.~255~255~255~1~0"
Mensaje(203) = "Para votar escribe /VOTO NICKNAME.~255~255~255~1~0"
Mensaje(204) = "Solo se computará un voto por miembro.~255~255~255~1~0"
Mensaje(205) = "Para fundar un clan debes de ser nivel 25 o superior.~255~255~255~1~0"
Mensaje(206) = "Para fundar un clan necesitas al menos 90 puntos en liderazgo.~255~255~255~1~0"
Mensaje(207) = "Los datos del clan son inválidos, asegúrate que no contiene caracteres inválidos.~255~255~255~1~0"
Mensaje(208) = "Ya existe un clan con ese nombre.~255~255~255~1~0"
Mensaje(209) = "El otro usuario aún no ha aceptado tu oferta.~255~255~255~0~0"
Mensaje(210) = "No tienes esa cantidad.~255~255~255~0~0"
Mensaje(211) = "No tienes mas espacio en el banco!!~65~190~156~0~0"
Mensaje(212) = "El banco no puede cargar tantos objetos.~65~190~156~0~0"
Mensaje(213) = "No puedes susurrarle a los Gms.~65~190~156~0~0"
Mensaje(214) = "Usuario inexistente.~65~190~156~0~0"
Mensaje(215) = "Has dejado de descansar.~65~190~156~0~0"
Mensaje(216) = "Dejas de meditar.~65~190~156~0~0"
Mensaje(217) = "No podes moverte porque estas paralizado.~65~190~156~0~0"
Mensaje(218) = "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.~65~190~156~0~0"
Mensaje(219) = "No podes usar así este arma.~65~190~156~0~0"
Mensaje(220) = "No puedes tomar ningún objeto.~65~190~156~0~0"
Mensaje(221) = "Has salido del modo de combate.~65~190~156~0~0"
Mensaje(222) = "Has pasado al modo de combate.~65~190~156~0~0"
Mensaje(223) = "No perteneces a ningún clan.~65~190~156~0~0"
Mensaje(224) = "NADA.~255~255~255~1~0"
Mensaje(225) = "NADA.~255~255~255~1~0"
Mensaje(226) = "Has rechazado la oferta del otro usuario.~255~255~255~0~0"
Mensaje(227) = "NADA~65~190~156~0~0"
Mensaje(228) = "NADA.~65~190~156~0~0"
Mensaje(229) = "No podes ocultarte si estás navegando.~65~190~156~0~0"
Mensaje(230) = "No tense municiones.~65~190~156~0~0"
Mensaje(231) = "¡No puedes atacarte a vos mismo!~65~190~156~0~0"
Mensaje(232) = "¡Para atacar ciudadanos desactiva el seguro!~255~200~200~1~0"
Mensaje(233) = "¡Primero selecciona el hechizo que quieres lanzar!~65~190~156~0~0"
Mensaje(234) = "No puedes pescar desde donde te encuentras.~65~190~156~0~0"
Mensaje(235) = "No hay agua donde pescar busca un lago, rió o mar.~65~190~156~0~0"
Mensaje(236) = "No podes robar aquí.~32~51~223~1~1"
Mensaje(237) = "No hay a quién robarle!.~65~190~156~0~0"
Mensaje(238) = "¡No podes robarle en zonas seguras!.~65~190~156~0~0"
Mensaje(239) = "Deberías equiparte el hacha.~65~190~156~0~0"
Mensaje(240) = "No podes talar desde allí.~65~190~156~0~0"
Mensaje(241) = "No hay ningún árbol ahí.~65~190~156~0~0"
Mensaje(242) = "Ahí no hay ningún yacimiento.~65~190~156~0~0"
Mensaje(243) = "No podes domar una criatura que está luchando con un jugador.~65~190~156~0~0"
Mensaje(244) = "No podes domar a esa criatura.~65~190~156~0~0"
Mensaje(245) = "No hay ninguna criatura allí!.~65~190~156~0~0"
Mensaje(246) = "No tienes mas minerales.~65~190~156~0~0"
Mensaje(247) = "Ahí no hay ninguna fragua.~65~190~156~0~0"
Mensaje(248) = "Ahí no hay ningún yunque.~65~190~156~0~0"
Mensaje(249) = "Felicidades has creado el primer clan de Argentum!!!~65~190~156~0~0"
Mensaje(250) = "%%%%%%%%%% INFORMACION DEL HECHIZO %%%%%%%%%%~65~190~156~0~0"
Mensaje(251) = "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%~65~190~156~0~0"
Mensaje(252) = "¡Primero selecciona el hechizo!~65~190~156~0~0"
Mensaje(253) = "No estás comerciando.~65~190~156~0~0"
Mensaje(254) = "No puedes cambiar tu oferta.~255~255~255~0~0"
Mensaje(255) = "No puedes salir estando paralizado.~32~51~223~1~1"
Mensaje(256) = "Comercio cancelado.~255~255~255~0~0"
Mensaje(257) = "Ya has fundado un clan, solo se puede fundar uno por personaje.~65~190~156~0~0"
Mensaje(258) = "Eres líder de un clan, no puedes salir del mismo.~65~190~156~0~0"
Mensaje(259) = "Te acomodas junto a la fogata y comienzas a descansar.~65~190~156~0~0"
Mensaje(260) = "Te levantas.~65~190~156~0~0"
Mensaje(261) = "No hay ninguna fogata junto a la cual descansar.~65~190~156~0~0"
Mensaje(262) = "Comienzas a meditar.~65~190~156~0~0"
Mensaje(263) = "Dejas de meditar.~65~190~156~0~0"
Mensaje(264) = "¡¡No puedes comerciar con los muertos!!~65~190~156~0~0"
Mensaje(265) = "No puedes comerciar con vos mismo...~65~190~156~0~0"
Mensaje(266) = "No puedes comerciar con el usuario en este momento.~65~190~156~0~0"
Mensaje(267) = "No perteneces a ningún clan.~65~190~156~0~0"
Mensaje(268) = "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.~65~190~156~0~0"
Mensaje(269) = "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.~65~190~156~0~0"
Mensaje(270) = "La descripción tiene caracteres inválidos.~65~190~156~0~0"
Mensaje(271) = "La descripción ha cambiado.~65~190~156~0~0"
Mensaje(272) = "El password debe tener al menos 6 caracteres.~65~190~156~0~0"
Mensaje(273) = "El password ha sido cambiado.~65~190~156~0~0"
Mensaje(274) = "Debes indicar el monto de cuánto quieres retirar.~65~190~156~0~0"
Mensaje(275) = "No puedes inmovilizarte a ti mismo.~255~0~0~1~0"
Mensaje(276) = "Este mapa es demasiado peligroso para tu nivel.~65~190~156~0~0"
Mensaje(277) = "No puedes lanzar hechizos curativos a un criminal si tú no lo eres, para hacerlo debes quitar el seguro, y tú también te convertirás en criminal.~65~190~156~0~0"
Mensaje(278) = "Para talar el árbol de tejo, debes equipar el hacha dorada.~65~190~156~0~0"
Mensaje(279) = "Debes equipar el hacha de leñador.~65~190~156~0~0"
Mensaje(280) = "Este npc no está herido.~65~190~156~0~0"
Mensaje(281) = "¡¡El usuario ya está vivo!!~65~190~156~0~0"
Mensaje(282) = "Necesitas un mejor báculo para éste hechizo.~65~190~156~0~0"
Mensaje(283) = "El npc es inmune a este hechizo.~255~0~0~1~0"
Mensaje(284) = "No está herido.~65~190~156~0~0"
Mensaje(285) = "No estás herido.~65~190~156~0~0"
Mensaje(286) = "Tu clase no puede usar éste objeto.~65~190~156~0~0"
Mensaje(287) = "Sólo los newbies pueden usar estos objetos.~65~190~156~0~0"
Mensaje(288) = "¡Has vuelto a ser visible!~65~190~156~0~0"
Mensaje(289) = "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.~65~190~156~0~0"
Mensaje(290) = "Estás muy lejos del usuario.~65~190~156~0~0"
Mensaje(291) = "¡No puedes cerrar en movimiento!~32~51~223~1~1"
Mensaje(292) = "¡¡No podes atacar a nadie porque estas muerto!!~65~190~156~0~0"
Mensaje(293) = "¡¡Estás muerto!! Los muertos no pueden tomar objetos.~65~190~156~0~0"
Mensaje(294) = "Estás demasiado lejos para pescar.~65~190~156~0~0"
Mensaje(295) = "!!El personaje no existe, cree uno nuevo.~65~190~156~0~0"
Mensaje(296) = "¡¡Has sido resucitado!!~65~190~156~0~0"
Mensaje(297) = "¡¡No puedes comerciar si estás navegando!!~65~190~156~0~0"
Mensaje(298) = "Estás demasiado lejos del usuario.~65~190~156~0~0"
Mensaje(299) = "Primero has click izquierdo sobre el personaje.~65~190~156~0~0"
Mensaje(300) = "No puedes cambiarte la descripción si estás muerto.~65~190~156~0~0"
Mensaje(301) = " Tu denuncia fue enviada.~65~190~156~0~0"
Mensaje(302) = "El personaje no está online.~65~190~156~0~0"
Mensaje(303) = "Por el momento no se pueden crear más partís.~255~200~200~1~0"
Mensaje(304) = "La party está llena, no puedes entrar.~255~200~200~1~0"
Mensaje(305) = "¡Has creado una party!~255~200~200~1~0"
Mensaje(306) = "No puedes hacerte líder.~255~200~200~1~0"
Mensaje(307) = "¡Te has convertido en líder de la party!~255~200~200~1~0"
Mensaje(308) = "No tienes suficientes puntos de liderazgo para liderar una party.~255~200~200~1~0"
Mensaje(309) = "Ya pertenecés a una party.~255~200~200~1~0"
Mensaje(310) = "Ya pertenecés a una party, pulsa ''F7'' y luego SALIR PARTY para abandonarla.~255~200~200~1~0"
Mensaje(311) = "El fundador decidirá si te acepta en la party.~255~200~200~1~0"
Mensaje(312) = "Para ingresar a una party debes hacer click sobre el fundador y luego presionar la tecla ''F3''.~255~200~200~1~0"
Mensaje(313) = "No eres miembro de ninguna party.~255~200~200~1~0"
Mensaje(314) = "Los retos están desactivados!.~65~190~156~0~0"
Mensaje(315) = "Debes ser nivel 18 o mas!.~65~190~156~0~0"
Mensaje(316) = "La apuesta mínima es de 5000 monedas de oro!.~65~190~156~0~0"
Mensaje(317) = "Los rings están llenos!.~65~190~156~0~0"
Mensaje(318) = "No puedes crear un reto si estás muerto!!!~65~190~156~0~0"
Mensaje(319) = "¡Para crear un reto debes estar en zona segura!~65~190~156~0~0"
Mensaje(320) = "No puedes retarte a ti mismo.~65~190~156~0~0"
Mensaje(321) = "No puedes retar a un muerto.~65~190~156~0~0"
Mensaje(322) = "Estás demasiado lejos del usuario.~65~190~156~0~0"
Mensaje(323) = "No tienes suficiente oro.~65~190~156~0~0"
Mensaje(324) = "Debes seleccionar un personaje.~65~190~156~0~0"
Mensaje(325) = "Has tranformado al mapa en zona insegura.~65~190~156~0~0"
Mensaje(326) = "Has tranformado al mapa en zona segura.~65~190~156~0~0"
Mensaje(327) = "El jugador no está online.~65~190~156~0~0"
Mensaje(328) = "Usuario no conectado.~65~190~156~0~0"
Mensaje(329) = "Error en el archivo.~65~190~156~0~0"
Mensaje(330) = "Offline.~65~190~156~0~0"
Mensaje(331) = "El usuario no está online.~255~255~255~0~0"
Mensaje(332) = "Usuario offline, Echando de los consejos.~65~190~156~0~0"
Mensaje(333) = "El Personaje está online, debe salir para el cambio.~32~51~223~1~1"
Mensaje(334) = "Transferencia exitosa.~65~190~156~0~0"
Mensaje(335) = "El nick solicitado ya existe.~65~190~156~0~0"
Mensaje(336) = "No está permitido utilizar valores mayores a 950000. Su comando ha quedado en los logs del juego.~65~190~156~0~0"
Mensaje(337) = "No está permitido utilizar valores mayores a mucho. Su comando ha quedado en los logs del juego.~65~190~156~0~0"
Mensaje(338) = "INIs recargados.~65~190~156~0~0"
Mensaje(339) = "OBJData recargado.~65~190~156~0~0"
Mensaje(340) = "El recuperar fue ejecutado correctamente.~65~190~156~0~0"
Mensaje(341) = "¡¡No puedes hacer fogatas en zonas seguras!!~65~190~156~0~0"
Mensaje(342) = "Este npc es inmune a los hechizos.~65~190~156~0~0"
Mensaje(343) = "Este hechizo no afecta a los muertos.~65~190~156~0~0"
Mensaje(344) = "Necesitas un instrumento mágico para devolver la vida.~65~190~156~0~0"
Mensaje(345) = "Escribe /SEG para quitar el seguro.~255~0~0~1~0"
Mensaje(346) = "Has logrado desarmar a tu oponente!~255~0~0~1~0"
Mensaje(347) = "Éste arma solo tiene efecto sobre los dragones.~255~0~0~1~0"
Mensaje(348) = "Tu anillo rechaza los efectos del hechizo.~255~0~0~1~0"
Mensaje(349) = "¡El hechizo no tiene efecto!~255~0~0~1~0"
Mensaje(350) = "Tu clase no puede utilizar este hechizo.~65~190~156~0~0"
Mensaje(351) = "Primero tenés que seleccionar una criatura, hace click izquierdo sobre ella.~65~190~156~0~0"
Mensaje(352) = "Necesitas el anillo mágico para devolver la vida.~65~190~156~0~0"
'marche
Mensaje(353) = "No puede atacar a una mascota en una zona segura.~32~51~223~1~1"
Mensaje(354) = "No tienes la inteligencia y la habilidad necesaria para trabajar con este material.~65~190~156~0~0"
Mensaje(355) = "Has abandonado a tu criatura.~65~190~156~0~0"
Mensaje(356) = "¡El usuario ya está invisible!~65~190~156~0~0"
Mensaje(357) = "¡Está muerto!~65~190~156~0~0"
Mensaje(358) = "No está envenenado.~65~190~156~0~0"
Mensaje(359) = "No estás envenenado.~65~190~156~0~0"
Mensaje(360) = "Solo el fundador puede expulsar miembros de una party.~255~200~200~1~0"
Mensaje(361) = "No eres miembro de ninguna party.~255~200~200~1~0"
Mensaje(362) = "¡Está muerto, no puedes aceptar miembros en ese estado!~255~200~200~1~0"
Mensaje(363) = "No eres líder, no puedes aceptar miembros.~255~200~200~1~0"
Mensaje(364) = "¡No se ha hecho el cambio de mando!~255~200~200~1~0"
Mensaje(365) = "¡No eres el líder!~255~200~200~1~0"
Mensaje(366) = "No podes atacar porque estas muerto.~65~190~156~0~0"
Mensaje(367) = "Para atacar mascotas de ciudadanos debes quitarte el seguro.~255~0~0~1~0"
Mensaje(368) = "¡¡No puedes dejar de navegar en mitad del mar!!~65~190~156~0~0"
Mensaje(369) = "¡¡No puedes hacer fogatas en zonas seguras!!~65~190~156~0~0"
Mensaje(370) = "¡Tu oponente te ha desarmado!~255~0~0~1~0"
Mensaje(371) = "¡No puedes revivir a alguien que esta en modo combate!~255~0~0~1~0"
Mensaje(372) = "¡No estás herido!~255~0~0~1~0"
'[Wizard 03/09/05]
Mensaje(373) = "¡Los consejeros no pueden comerciar con otros usuarios!~255~255~255~0~0"
Mensaje(374) = "El usuario no pertenece a ninguna party.~255~255~255~0~0"
Mensaje(375) = "Empiezas a escuchar los pmsg de la party del usuario.~255~255~255~0~0"
Mensaje(376) = "Ya hay otro Gm escuchando a esta party.~255~255~255~0~0"
Mensaje(377) = "Te has desconcentrado, dejas de meditar.~65~190~156~0~0"
Mensaje(378) = "No puedes arrojar al suelo los objetos newbies.~65~190~156~0~0"
Mensaje(379) = "Para atacar criaturas no hostiles deberas desactivar el seguro.~65~190~156~0~0"
Mensaje(380) = "Para alterar tu faccion, deberas antes salir de tu clan.~65~190~156~0~0"
'[/Wizard]=> soy original y hago la barrita alrevez;)
'[\]Misery_Ezequiel 26/06/05]
'[\Nacho]
Mensaje(381) = "¡No puedes extraer leña de ahi.!~65~190~156~0~0"
'Ya no manda mas de que rango es!
ReDim RangoArmada(0 To NumArm) As String
RangoArmada(0) = " <Ejercito real> " & "<Aprendiz>"
RangoArmada(1) = " <Ejercito real> " & "<Escudero>"
RangoArmada(2) = " <Ejercito real> " & "<Caballero>"
RangoArmada(3) = " <Ejercito real> " & "<Capitan>"
RangoArmada(4) = " <Ejercito real> " & "<Teniente>"
RangoArmada(5) = " <Ejercito real> " & "<Comandante>"
RangoArmada(6) = " <Ejercito real> " & "<Senescal>"
RangoArmada(7) = " <Ejercito real> " & "<Protector>"
RangoArmada(8) = " <Ejercito real> " & "<Guardian del Bien>"
RangoArmada(9) = " <Ejercito real> " & "<Campeón de la Luz>"
'Lo mismo que para la armada
ReDim RangoCaos(0 To NumCaos) As String
RangoCaos(0) = " <Legión oscura> " & "<Esbirro>"
RangoCaos(1) = " <Legión oscura> " & "<Servidor de las Sombras>"
RangoCaos(2) = " <Legión oscura> " & "<Acólito>"
RangoCaos(3) = " <Legión oscura> " & "<Guerrero Sombrío>"
RangoCaos(4) = " <Legión oscura> " & "<Sanguinario>"
RangoCaos(5) = " <Legión oscura> " & "<Caballero de la Oscuridad>"
RangoCaos(6) = " <Legión oscura> " & "<Condenado>"
RangoCaos(7) = " <Legión oscura> " & "<Heraldo Impío>"
RangoCaos(8) = " <Legión oscura> " & "<Corruptor>"
RangoCaos(9) = " <Legión oscura> " & "<Devorador de Almas>"





ReDim ListaClases(1 To NUMCLASES) As String
ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Paladin"
ListaClases(9) = "Cazador"
ListaClases(10) = "Pescador"
ListaClases(11) = "Herrero"
ListaClases(12) = "Leñador"
ListaClases(13) = "Minero"
ListaClases(14) = "Carpintero"
ListaClases(15) = "Pirata"

ReDim SkillsNames(1 To NUMSKILLS) As String
SkillsNames(1) = "Resistencia mágica"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar árboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"

ReDim UserSkills(1 To NUMSKILLS) As Integer
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"

frmOldPersonaje.NameTxt.Text = Config_Inicio.Name
frmOldPersonaje.PasswordTxt.Text = ""


IniciarObjetosDirectX

frmCargando.Label1 = "Cargando Sonidos...."

Dim loopc As Integer
LastTime = GetTickCount

ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)

Call InitTileEngine(frmMain.hWnd, 152, 7, 32, 32, 13, 17, 9)

'Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extras.", 2, 51, 223, 1, 1)
frmCargando.Label1 = "Creando animaciones extra...."
Call CargarAnimsExtra
Call CargarTips
UserMap = 1
Call CargarArrayLluvia
Call CargarAnimArmas
Call CargarAnimEscudos
Call CargarVersiones



Unload frmCargando
LoopMidi = True
If Musica = 0 Then
    Call CargarMIDI(DirMidi & MIdi_Inicio & ".mid")
    Play_Midi
End If

frmPres.Picture = LoadPicture(App.Path & "\Graficos\noland.jpg")
'frmPres.WindowState = vbMaximized
frmPres.Show

Do While Not finpres
    DoEvents
Loop

Unload frmPres
frmConnect.Visible = True
'Loop principal!
'[CODE]:MatuX'
    MainViewRect.left = MainViewLeft + 32 * RenderMod.iImageSize
    MainViewRect.top = MainViewTop + 32 * RenderMod.iImageSize
    MainViewRect.right = (MainViewRect.left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.bottom = (MainViewRect.top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

    MainDestRect.left = ((TilePixelWidth * TileBufferSize) - TilePixelWidth) + 32 * RenderMod.iImageSize
    MainDestRect.top = ((TilePixelHeight * TileBufferSize) - TilePixelHeight) + 32 * RenderMod.iImageSize
    MainDestRect.right = (MainDestRect.left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainDestRect.bottom = (MainDestRect.top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

    Dim OffsetCounterX As Integer
    Dim OffsetCounterY As Integer
'[END]'

Dim mainAntX As Long, mainAntY As Long
'Dim DEBUG_LoopTime As Long 'by Gorlok

PrimeraVez = True
prgRun = True
pausa = False
bInvMod = True
lFrameLimiter = DirectX.TickCount
'[CODE 001]:MatuX'
    lFrameModLimiter = 60
'[END]'
Do While prgRun
    'DEBUG_LoopTime = GetTickCount
    
    'If RequestPosTimer > 0 Then
        'RequestPosTimer = RequestPosTimer - 1
        'If RequestPosTimer = 0 Then
            'Pedimos que nos envie la posicion
         '   Call SendData("RPU")
   '     End If
   ' End If
'    Call RefreshAllChars
    '[CODE 001]:MatuX
    '
    '   EngineRun
    If EngineRun Then
        '[DO]:Dibuja el siguiente frame'
        '[CODE 000]:MatuX'
        'If frmMain.WindowState <> 1 And CurMap > 0 And EngineRun Then
        If frmMain.WindowState <> 1 Then
        '[END]'
            'Call ShowNextFrame(frmMain.Top, frmMain.Left)
            '****** Move screen Left, Right, Up and Down if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            ElseIf AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If
            '****** Update screen ******
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            'Call DoNightFX
            'Call DoLightFogata(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
            '[CODE 000]:MatuX
                'Call MostrarFlags
                If IScombate Then Call Dialogos.DrawText(260, 260, "Modo Combate", vbRed)
                If Istrabajando Then Call Dialogos.DrawText(260, 260, "Trabajando", vbWhite)
                If Dialogos.CantidadDialogos <> 0 Then Call Dialogos.MostrarTexto
                If Cartel Then Call DibujarCartel
                If bInvMod Then DibujarInv
                
                
                If Activado Then

                Call Dialogos.DrawText(260, 606, clantext5, vbGreen)
                Call Dialogos.DrawText(260, 617, clantext4, vbGreen)
                Call Dialogos.DrawText(260, 628, clantext3, vbGreen)
                Call Dialogos.DrawText(260, 639, clantext2, vbGreen)
                Call Dialogos.DrawText(260, 650, clantext1, vbGreen)

                End If
                
                If mainAntX <> frmMain.left Or mainAntY <> frmMain.top Then
                    mainAntX = frmMain.left
                    mainAntY = frmMain.top
                    MainViewRect.left = (frmMain.left / Screen.TwipsPerPixelX) + MainViewLeft + 32 * RenderMod.iImageSize
                    MainViewRect.top = (frmMain.top / Screen.TwipsPerPixelY) + MainViewTop + 32 * RenderMod.iImageSize
                    MainViewRect.right = (MainViewRect.left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
                    MainViewRect.bottom = (MainViewRect.top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)
                End If
                
                Call DrawBackBufferSurface
               
               ' Call RenderSounds
                
                '[DO]:Inventario'
                'Call DibujarInv(frmMain.picInv.hWnd, 0)
                'If bInvMod Then DibujarInv  'lo moví arriba para
                '                             que esté mas ordenadito
                '[END]'
            '[END]'
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    End If
    '[CODE 000]:MatuX'
    'If ControlVelocidad(LastTime) Then
    If (GetTickCount - LastTime > 20) Then
        If Not pausa And frmMain.Visible And Not frmForo.Visible Then
            CheckKeys
            LastTime = GetTickCount
        End If
    End If
    'If Musica = 0 Then
    '    If Not SegState Is Nothing Then
          '  If Not Perf.IsPlaying(Segs, SegState) Then Play_Midi
     '   End If
   ' End If
         'Musica = 0
    'End If
    '[END]'
    '[CODE 001]:MatuX
    ' Frame Limiter
        'FramesPerSec = FramesPerSec + 1
        If DirectX.TickCount - lFrameTimer > 1000 Then
            FramesPerSec = FramesPerSecCounter
            If FPSFLAG Then frmMain.Caption = FramesPerSec
            FramesPerSecCounter = 0
            lFrameTimer = DirectX.TickCount
        End If
        'While DirectX.TickCount - lFrameLimiter < lFrameModLimiter: Wend
        '[Alejo]
            While DirectX.TickCount - lFrameLimiter < 55
                Sleep 5
            Wend
        '[/Alejo]
        lFrameLimiter = DirectX.TickCount
    '[END]'
    'Sistema de timers renovado:
    esttick = GetTickCount
    For loopc = 1 To UBound(timers)
        timers(loopc) = timers(loopc) + (esttick - ulttick)
        'timer de trabajo
        If timers(1) >= tUs Then
            timers(1) = 0
            NoPuedeUsar = False
        End If
        'timer de attaque (77)
        If timers(2) >= tAt Then
            timers(2) = 0
            UserCanAttack = 1
            UserPuedeRefrescar = True
        End If
    Next loopc
    ulttick = GetTickCount
    
'   Debug.Print "[DEBUG] LoopTime: " & (GetTickCount - DEBUG_LoopTime)
    
    DoEvents
Loop

EngineRun = False
frmCargando.Show
frmCargando.Label1 = "Liberando recursos..."
LiberarObjetosDX

If bNoResChange = False Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
    
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
           .dmPelsHeight = oldResHeight
        End With
lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
End If

Call UnloadAllForms

Config_Inicio.tip = tipf
Call EscribirGameIni(Config_Inicio)

End

ManejadorErrores:
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.description & " Fuente:" & Err.Source
    End
    
End Sub

Sub WriteVar(File As String, Main As String, var As String, value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
writeprivateprofilestring Main, var, value, File
End Sub

Function GetVar(File As String, Main As String, var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish

getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = left(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
    Public Function CheckMailString(ByRef sString As String) As Boolean
        On Error GoTo errHnd:
        Dim lPos  As Long, lX    As Long
        Dim iAsc  As Integer
    
        '1er test: Busca un simbolo @
        lPos = InStr(sString, "@")
        If (lPos <> 0) Then
            '2do test: Busca un simbolo . después de @ + 1
            If Not (IIf((InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1)), True, False)) Then _
                Exit Function
    
            '3er test: Valída el ultimo caracter
            If Not (CMSValidateChar_(Asc(right(sString, 1)))) Then _
                Exit Function
    
            '4to test: Recorre todos los caracteres y los valída
            For lX = 0 To Len(sString) - 1 'el ultimo no porque ya lo probamos
                If Not (lX = (lPos - 1)) Then
                    iAsc = Asc(Mid(sString, (lX + 1), 1))
                    If Not (iAsc = 46 And lX > (lPos - 1)) Then _
                        If Not CMSValidateChar_(iAsc) Then _
                            Exit Function
                End If
            Next lX
    
            'Finale
            CheckMailString = True
        End If
    
errHnd:
        'Error Handle
    End Function
    
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean
CMSValidateChar_ = IIf( _
                    (iAsc >= 48 And iAsc <= 57) Or _
                    (iAsc >= 65 And iAsc <= 90) Or _
                    (iAsc >= 97 And iAsc <= 122) Or _
                    (iAsc = 95) Or (iAsc = 45), True, False)
End Function

Function HayAgua(X As Integer, Y As Integer) As Boolean

If MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
   MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
   MapData(X, Y).Graphic(2).GrhIndex = 0 Then
            HayAgua = True
Else
            HayAgua = False
End If
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub
    
Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub

Public Sub LeerLineaComandos()
Dim Tmp As String, T() As String
Dim I As Long

'inicializo los parametros estandar
NoRes = False 'si esta en false, la cambio

Tmp = Command
T = Split(Tmp, " ")

I = LBound(T)
Do While I <= UBound(T)
    Select Case UCase(T(I))
    Case "/NORES" 'no cambiar la resolucion
        NoRes = True
    End Select
    I = I + 1
Loop
End Sub

Public Sub LogDebug(desc As String)
On Error GoTo ErrHandler

Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\debug.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub
ErrHandler:

End Sub

Public Sub CapturarPantalla()
    Dim FreeImage1 As Long
    Dim bOK As Long
    Dim strFName As String
    Dim strFTemp As String
    
    
    strFTemp = App.Path & "\fotos\screen.bmp"
    
    ' hide the form
    ' (as we don't want this in the screen shot)
    DoEvents
    Clipboard.Clear
    
    ' send a print screen button keypress event
    ' and DoEvents to allow windows time to process
    ' the event and capture the image to the clipboard
    keybd_event vbKeySnapshot, 0, 0, 0
    DoEvents
    ' send a print screen button up event
    keybd_event vbKeySnapshot, 0, &H2, 0
    DoEvents
    DoEvents
    ' paste the clipboard contents into the picture box
    '[DEBUGUED BY WIZARD] Totalmente al pedo esto.
    'frmMain.ScreenCapture.Picture = Clipboard.GetData(vbCFBitmap)
    'DoEvents
    'DoEvents
    '[/WIZARD]
    ' change the pointer to an hourglass while the image is processed
    Screen.MousePointer = vbHourglass
    ' save the image to a file using the application path
    '[Wizard; Grabamos la imagen directamente desde el porta papeles]
    SavePicture Clipboard.GetData(vbCFBitmap), strFTemp
    DoEvents
    ' use the FreeImage.dll (http://freeimage.sourceforge.net/)
    ' to load the screen image
    FreeImage1 = FreeImage_Load(FIF_BMP, strFTemp, 0)
    ' save the screen capture as an JPEG image with high quality
    strFName = format(Now, "yyyy_mm_dd_hh_mm_ss")
    'strFName = Replace(Now, "/", "_")
    'strFName = Replace(strFName, ":", "_")
    'strFName = Replace(strFName, " ", "_")
    strFName = App.Path & "\fotos\TDS_foto_" & strFName & ".jpg"
    bOK = FreeImage_Save(FIF_JPEG, FreeImage1, strFName, &H80)
    'unload the images
    FreeImage_Unload (FreeImage1)
    ' restore the mouse pointer
    Screen.MousePointer = vbNormal
    Kill strFTemp
    Call AddtoRichTextBox(frmMain.RecTxt, "Guardado: " & strFName, 0, 200, 200, False, False, False)
End Sub
'********************Misery_Ezequiel 28/05/05********************'
Private Sub DejarDeTrabajar()
Istrabajando = False
frmMain.IntervaloLaburar.enabled = False
frmMain.Macro.enabled = False
SendData ("DEJ")
 Call AddtoRichTextBox(frmMain.RecTxt, "Has terminado de trabajar.", 0, 200, 200, False, False, False)
End Sub
Public Sub Play_Song(song_name As String)
If Fx = 1 Then Exit Sub
On Error Resume Next
Dim balance As Integer
VolumeN = frmOpciones.Slider1
    If VolumeN > 0 Then VolumeN = 0
    IBA.volume = VolumeN
    Set IMPos = IMC
    IMPos.CurrentPosition = 0
    IMC.Run
End Sub
Public Sub DejarDeTrabajars()
Istrabajando = False
frmMain.IntervaloLaburar.enabled = False
frmMain.Macro.enabled = False
SendData ("DEJ")
 Call AddtoRichTextBox(frmMain.RecTxt, "Has terminado de trabajar.", 0, 200, 200, False, False, False)
End Sub
Public Function Music_MP3_Load(ByVal file_path As String, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean '**** Loads a MP3 *****
    Set IMC = New FilgraphManager
    IMC.RenderFile file_path
    Set IBA = IMC
  
End Function
