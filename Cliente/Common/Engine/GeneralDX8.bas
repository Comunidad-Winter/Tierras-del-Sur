Attribute VB_Name = "modGeneral"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit


'Dim Interface_pak As clsFilePaker
Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public lFrameTimer As Long

Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, _
                        ByVal uFlags As Long, ByVal dwItem1 As Long, _
                        ByVal dwItem2 As Long)

' A file type association has changed.
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0



Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean

Public ping_timer As clsPerformanceTimer


Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
End Function

Sub Login()
    Call WriteLoginExistingChar
    DoEvents
    Call FlushBuffer
End Sub

Public Sub Make_Transparent_Richtext(ByVal Hwnd As Long)
'If Win2kXP Then
On Error Resume Next
    Call SetWindowLong(Hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
End Sub

Public Function DirGraficos() As String
    DirGraficos = App.path & "\Datos\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\WAV\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\Datos\armas.dat"
    
    NumWeaponAnims = val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = val(GetVar(App.path & "\Datos\versiones.ini", "Graficos", "Val"))
    Versiones(2) = val(GetVar(App.path & "\Datos\versiones.ini", "Wavs", "Val"))
    Versiones(3) = val(GetVar(App.path & "\Datos\versiones.ini", "Midis", "Val"))
    Versiones(4) = val(GetVar(App.path & "\Datos\versiones.ini", "Init", "Val"))
    Versiones(5) = val(GetVar(App.path & "\Datos\versiones.ini", "Mapas", "Val"))
    Versiones(6) = val(GetVar(App.path & "\Datos\versiones.ini", "E", "Val"))
    Versiones(7) = val(GetVar(App.path & "\Datos\versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\Datos\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).R = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).G = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).R = 255
    ColoresPJ(50).G = 0
    ColoresPJ(50).b = 0
    ColoresPJ(49).R = 0
    ColoresPJ(49).G = 128
    ColoresPJ(49).b = 255
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\Datos\escudos.dat"
    
    NumEscudosAnims = val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

'Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
''******************************************
''Adds text to a Richtext box at the bottom.
''Automatically scrolls to new text.
''Text box MUST be multiline and have a 3D
''apperance!
''Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
''Juan Mart�n Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
''******************************************r
'    With RichTextBox
'        If Len(.Text) > 1000 Then
''            'Get rid of first line
''            .SelStart = InStr(1, .Text, vbCrLf) + 1
''            .SelLength = Len(.Text) - .SelStart + 2
'            .TextRTF = vbNullString
'            .SelStart = 0
'        Else
'            .SelStart = Len(RichTextBox.Text)
'        End If
'
'        .SelLength = 0
'        .SelBold = bold
'        .SelItalic = italic
'
'        If Not red = -1 Then .SelColor = RGB(red, green, blue)
'
'        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
'        .SelStart = Len(RichTextBox.Text) - 1
'        RichTextBox.Refresh
'    End With
'End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    act_charmap
    For loopc = 1 To LastChar
        If CharList(loopc).active = 1 Then
            charmap(CharList(loopc).Pos.x, CharList(loopc).Pos.y) = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("�")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Direcci�n de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
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

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
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
    frmMain.Picture = modZLib.Bin_Resource_Load_Picture(2, rGUI)
    frmMain.InvEqu.Picture = modZLib.Bin_Resource_Load_Picture(3, rGUI)

    frmMain.Refresh
    frmMain.pri = True
    renderasd = True
    
    Call SetMusicInfo("Jugando Arduz AO: " & UserName & " - http://www.arduz.com.ar/", "", "", "Games", , "{0}")
On Error Resume Next
    Make_Transparent_Richtext frmMain.RecTxt.Hwnd


End Sub

'TODO : Si bien nunca estuvo all�, el mapa es algo independiente o a lo sumo dependiente del engine, no va ac�!!!

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (dir$(file, FileType) <> "")
End Function

Public Function IsIP(ByVal IP As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).IP = IP Then
            IsIP = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
'********************************
'Author: Unknown
'Last Modification: 07/26/07
'Last Modified by: Rapsodius
'Added Instruction "CloseClient" before End so the mutex is cleared
'********************************

End Sub

Public Sub InitServersList()

End Sub

Public Function CurServerPasRecPort() As Integer

End Function

Public Function CurServerIp() As String

End Function

Public Function CurServerPort() As Integer

End Function

Public Sub play_intro()
On Error Resume Next
Dim mp3intro As String
If Read_Cfg(Musica_act) Then
    mp3intro = App.path & "\Datos\Intro.mp3"
    If FileExist(mp3intro, vbNormal) Then
        Audio.Music_Stop
        Audio.Music_Load "Intro"
        Audio.Music_Play
    End If
    frmMain.musicc.Enabled = True
    frmConnect.Check1.value = vbChecked
Else
    frmMain.musicc.Enabled = False
    frmConnect.Check1.value = vbUnchecked
End If
End Sub


Sub borra_img(img As String)
If FileExist(Windows_Temp_Dir & img, vbArchive) Then Kill Windows_Temp_Dir & img
End Sub

Public Sub load_cfgs()
    volumenpotas = Read_Cfg(Volumen_potas)
    volumenfx = Read_Cfg(Volumen_fx)
    SoundActivated = Not Read_Cfg(Sonidos_act)
    useRDL = Read_Cfg(RadioDeLuz) <> 0
    useEDS = Read_Cfg(EfectosSol) <> 0
    limitarr = Read_Cfg(Limitar_Fps)
    'Engine.Engine_set_max_fps limitarr, 100
    SuperWater = Read_Cfg(eSuperWater) <> 0
    Force_Software = Read_Cfg(forzar_software) <> 0
    puedo_deslimitar = IsIDE
End Sub

Sub Main()
MsgBox (LenB(MapData(1, 1)) - 16) * 10000& / 1024& / 1024&

'cfnc = fnc.E_Main

If Not IsIDE() Then
    On Error GoTo errr
Else
    DoEvents
End If

Dim llegob As Byte

hechizo_cargado = CByte(109)

DoEvents

1 play_intro

DecimalSeparator

2 Windows_Temp_Dir = modEENESARIO.General_Get_Temp_Dir

3 modZLib.Bin_Load_Headers App.path & "\Datos\grhdata\"

4    Load frmMain
5    frmMain.Visible = False
    macaddr = "soycheatervieja"
6    Init_Hamachi
    
7    Set frmMain.WEBB = New clsWEBA
8    frmMain.WEBB.Initialize frmMain.WEbSOCK

'borra_img "connect.bmp"


    
IniPath = App.path & "\Datos\"


    INT_ATTACK = 1301 - RandomNumber(0, 10)
    INT_ARROWS = 1151 - RandomNumber(0, 10)
    INT_CAST_SPELL = 1151 - RandomNumber(0, 10)
    INT_CAST_ATTACK = 1151 - RandomNumber(0, 10)
    INT_WORK = 701 - RandomNumber(0, 10)
    INT_USEITEMU = 201 - RandomNumber(0, 10)
    INT_USEITEMDCK = 205 - RandomNumber(0, 10)
    INT_SENTRPU = 3001

9    Call LoadClientSetup

load_cfgs
    Load frmCargando
    DoEvents
    'Sleep 0&
'cfnc = fnc.E_Set_Res
10    Call Resolution.SetResolution
    
11    frmCargando.Picture = modZLib.Bin_Resource_Load_Picture(5, rGUI) 'General_Load_Picture_From_Resource("splash.bmp")

If Not IsIDE() Then
    On Local Error Resume Next
Else
    DoEvents
End If

    frmConnect.Timer3.Enabled = True
    frmConnect.CronList.Enabled = True

12  frmCargando.Show
    frmCargando.pb.max = 18
    
    frmCargando.Refresh
    frmCargando.pb.value = 2
    frmCargando.pb.Caption = vbNullString
    
    ChDrive App.path
    ChDir App.path
    
    MD5HushYo = "0123456789abcdef"
    
If Not IsIDE() Then
    On Local Error GoTo errr
Else
    DoEvents
End If

13    AddtoRichTextBox frmCargando.status, "Iniciando Nombres... ", 123, 123, 123, 0, 0, 0

14    Call InicializarNombres
    AddtoRichTextBox frmCargando.status, "Hecho", 123, 123, 123, 0, 0, 0

    AddtoRichTextBox frmCargando.status, "Iniciando Fuentes... ", 123, 123, 123, 0, 0, 0

15  Call Protocol.InitFonts
    
    frmOldPersonaje.NameTxt.Text = GetCfg(App.EXEName, "USER", "act", "Usuario")
    frmOldPersonaje.verpasswD

    AddtoRichTextBox frmCargando.status, "Hecho", 123, 123, 123, 0, 0, 0
    AddtoRichTextBox frmCargando.status, "Iniciando motor gr�fico... ", 123, 123, 123, 0, 0, 0
 
16  Call Engine_Init(99)
    AddtoRichTextBox frmCargando.status, "Hecho", 123, 123, 123, 0, 0, 0
    AddtoRichTextBox frmCargando.status, "Cargando indices... ", 123, 123, 123, 0, 0, 0
17  Call LoadGrhData
    AddtoRichTextBox frmCargando.status, "Cargando cuerpos... ", 123, 123, 123, 0, 0, 0
18  Call CargarCuerpos
19  Call CargarCabezas
20  Call CargarCascos
21  Call CargarFxs

    frmCargando.pb.value = 12
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra... ", 123, 123, 123, 0, 0, 0)
    
22  init_special_slots

    UserMap = 1
    
23  Call CargarAnimArmas
24  Call CargarAnimEscudos
25  Call CargarVersiones
26  Call CargarColores

    Init_weapons

    frmCargando.pb.value = 14
    
    AddtoRichTextBox frmCargando.status, "Hecho", 123, 123, 123, 0, 0, 0
    AddtoRichTextBox frmCargando.status, "Iniciando sonidos... ", 123, 123, 123, 0, 0, 0

    frmCargando.pb.value = 16

27  Inventory_init
    frmCargando.pb.value = 18
    AddtoRichTextBox frmCargando.status, "Hecho", 123, 123, 123, 0, 0, 0
    
#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

28    frmConnect.Picture = modZLib.Bin_Resource_Load_Picture(1, rGUI)
29    Unload frmCargando
    frmConnect.Visible = True
    DoEvents
    'Inicializaci�n de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    magicNumber = 1

    'Set the intervals of timers
30  Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
   'Init timers
31  Call MainTimer.start(TimersIndex.Attack)
    Call MainTimer.start(TimersIndex.Work)
    Call MainTimer.start(TimersIndex.UseItemWithU)
    Call MainTimer.start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.start(TimersIndex.SendRPU)
    Call MainTimer.start(TimersIndex.CastSpell)
    Call MainTimer.start(TimersIndex.Arrows)
    Call MainTimer.start(TimersIndex.CastAttack)
    'Set the dialog's font
    'Dialogos.font = frmMain.font
    'Hits.font = frmMain.font
    'DialogosClanes.font = frmMain.font
    
    AlphaActivadoX = True
    
    ' Load the form for screenshots
    'Call Load(frmScreenshots)
    
    Set ping_timer = New clsPerformanceTimer
    
32    Call SetMusicInfo("Jugando Arduz AO - http://www.arduz.com.ar/", "", "", "Games", , "{0}")

33    setup_ambient

34    Engine.start
    
Exit Sub
On Error GoTo 0
errr:
    send_error "CLIENT_ERR C�digo: " & Err.Number & vbNewLine & "Descripci�n: " & Err.description & vbNewLine & "FNC:" & cfnc & vbNewLine & "DLLE:" & Err.LastDllError & vbNewLine & "Ln:" & Erl & "-" & exerl
    Dim ms As Integer
    ms = MsgBox("Se produjo un error, por favor copia este texto y publicalo en el foro de Arduz asi podremos solucionarlo:" & vbNewLine & "C�digo: " & Err.Number & vbNewLine & "Descripci�n: " & Err.description & vbNewLine & "Funcion: " & cfnc & vbNewLine & "DllError: " & Err.LastDllError & " LINE:" & Erl, vbAbortRetryIgnore Or vbInformation, "Runtime error")
    If ms = vbAbort Then
        End
    ElseIf ms = vbRetry Then
        Err.Clear
    End If
    Audio.Music_Pause
    Resume
End
End Sub

'Sub setfpslabel(STR As String)
'frmMain.FPS.Caption = STR
'frmMain.Label2(0).Caption = STR
'frmMain.Label2(1).Caption = STR
'frmMain.Label2(2).Caption = STR
'frmMain.Label2(3).Caption = STR
'frmMain.Label2(4).Caption = STR
'End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Funci�n para chequear el email
'
'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . despu�s de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los val�da
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer ac�....
Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean
    HayAgua = ((MapData(x, y).Graphic(1).GrhIndex >= 1505 And MapData(x, y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(x, y).Graphic(1).GrhIndex >= 5665 And MapData(x, y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(x, y).Graphic(1).GrhIndex >= 13547 And MapData(x, y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(x, y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
'    If Not frmCantidad.Visible Then
'        frmMain.SendTxt.Visible = True
'        frmMain.SendTxt.SetFocus
'    End If
End Sub

Public Sub ShowSendCMSGTxt()

End Sub
    
Public Sub LeerLineaComandos()

End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 24/06/2006
'
'**************************************************************

   On Error GoTo LoadClientSetup_Error

    ClientSetup.bDinamic = GetCfg(App.EXEName, "CFG", "DYN", "1") = "1"
    ClientSetup.bNoRes = GetCfg(App.EXEName, "CFG", "NORES", "1") = "1"
    ClientSetup.bNoSound = GetCfg(App.EXEName, "CFG", "NOSOUND", "0") = "1"
    ClientSetup.bUseVideo = GetCfg(App.EXEName, "CFG", "VIDEO", "1") = "1"
    ClientSetup.byMemory = val(GetCfg(App.EXEName, "CFG", "VIDEOMEM", "16"))
    NoRes = ClientSetup.bNoRes
If IsIDE = False Then
    NoRes = IIf(MsgBox("Cambiar resolucion?", vbYesNo) = vbYes, True, False)
Else
    NoRes = False
End If
   On Error GoTo 0
   Exit Sub

LoadClientSetup_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure LoadClientSetup of M�dulo modGeneral"
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
   On Error GoTo InicializarNombres_Error

    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Argh�l"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Fisher) = "Pescador"
    ListaClases(eClass.Blacksmith) = "Herrero"
    ListaClases(eClass.Lumberjack) = "Le�ador"
    ListaClases(eClass.Miner) = "Minero"
    ListaClases(eClass.Carpenter) = "Carpintero"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Suerte) = "Suerte"
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apu�alar) = "Apu�alar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar �rboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Wrestling) = "Wrestling"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"

   On Error GoTo 0
   Exit Sub

InicializarNombres_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure InicializarNombres of M�dulo modGeneral"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    'Call Hits.RemoveAllHits
    Call Dialogos.RemoveAllDialogs
End Sub




Public Sub CloseClient()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    
    
    frmCargando.Show
    AddtoRichTextBox frmCargando.status, "Liberando recursos...", 123, 123, 123, 0, 0, 0
    
    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    'Destruimos los objetos p�blicos creados


    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing

    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Set frmMain.WEBB = Nothing
    
    frmMain.Winsock1.Close
    frmMain.WEbSOCK.Close
    
    Call UnloadAllForms
    
    EngineRun = False
    
End Sub


Public Function encode_decode_text(Text As String, ByVal off As Integer, Optional ByVal cript As Byte, Optional ByVal encode As Byte) As String
    Dim i As Integer, l As String
    If encode Then off = 256 - off
    Dim ba() As Byte, bo() As Byte
    Dim lenn%
    ba = StrConv(Text, vbFromUnicode)
    lenn = UBound(ba)
    ReDim bo(0 To lenn)
    For i = 0 To lenn
       bo(i) = ((ba(i) Xor cript) + off) Mod 256 Xor cript
    Next i
    encode_decode_text = StrConv(bo, vbUnicode)
End Function

