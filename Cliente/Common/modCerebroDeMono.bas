Attribute VB_Name = "CDMCerebroDeMono"
Option Explicit

Public Enum CDM_TIPO_UPDATE
    CDM_Upd_Graficos_ind = 1
    CDM_Upd_Tilesets
    CDM_Upd_Presets
    CDM_Upd_Particulas
    CDM_Upd_Grafico
    CDM_Upd_Sonidos
    CDM_Upd_Mapas
End Enum

Public CDM_TMP_PATH         As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type CDM_File
    FileName    As String
    Remote      As String
    MD5         As String * 32
    Parchear    As Boolean
    aparchear   As Byte
    Newer       As Boolean
    nnum        As Long
    Tipo        As CDM_TIPO_UPDATE
    numero      As Integer
    size        As Long
    datos       As String
    id          As Long
    
    C1          As Long
    C2          As Long
    C3          As Long
    C4          As Long
    
    user        As Long
    PrivsPublicos As Long
    
    Data()      As Byte
    data_lenght As Long
    
    EstaListo      As Boolean
    Fallo           As Boolean
    
End Type

Public CDMDownloading       As Boolean
Public CDM_Current_File     As CDM_File

Public Type Script
    title   As String
    date    As Date
    Files() As CDM_File
    Lista   As Byte
End Type

' Constants

Private Const psScriptToken As String = "$ArduzScript$"

' Privates
Public CDM_Script As Script
Public CDM_Remote_File As String

Public actual_version As Long
' <<< END DECLARATIONS

Public B64 As New Base64Class

Private Type CDM_Comit_UDT
    active          As Byte
    FilenameOrData  As String
    num             As Integer
    Tipo            As CDM_TIPO_UPDATE
    localeID        As Integer
    MIME            As String
    FileName        As String
    MD5             As String * 32
End Type

Private ColaComiteos()  As CDM_Comit_UDT
Private CDM_ENCola       As Integer
Private TmpComit        As CDM_Comit_UDT

Public CDM_SinEnviar    As Boolean

Public CDM_Revision     As Long

'Unix timestamp

Private Type SystemTime
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SystemTime
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SystemTime
        DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public CDM_Username     As String
Public CDM_Password     As String
Public CDM_PasswordMD5  As String
Public CDM_UserID       As Long
Public CDM_UserPrivs    As Long
Public CDM_UserSession  As Long

Public Enum CDM_Privs
    EsUsuario = 1
    PuedeComitearGraficos = 2
    PuedeComitearMapas = 4
    PuedeIndexar = 8
    PuedeBorrarPropio = 16
    PuedeBorrarCualquiera = 32
    PuedeEditarCualquiera = 64
    PuedeExtraerArchivos = 128
    PuedeHacerRollBack = 256
End Enum

Public vWindowCDM As vWindow

Public ColaDescargas As Collection

Public Sub IniciarCDM()
    Set vWindowCDM = New vWCDM
End Sub

Public Function FromUnixTime(ByVal sUnixTime As Long) As Date
    Dim NTime As Date, STime As Date
    Dim TZ As TIME_ZONE_INFORMATION
    STime = #1/1/1970#
    NTime = DateAdd("s", sUnixTime, STime)
    GetTimeZoneInformation TZ
    NTime = DateAdd("n", -TZ.Bias, NTime)
    FromUnixTime = NTime
End Function

Public Function ToUnixTime(ByVal STime As Date) As Long
    Dim NTime As Date, sUnix As Date, sUnixTime As Long
    Dim TZ As TIME_ZONE_INFORMATION
    sUnix = #1/1/1970#
    GetTimeZoneInformation TZ
    NTime = DateAdd("n", TZ.Bias, STime)
    sUnixTime = DateDiff("s", sUnix, NTime)
    ToUnixTime = sUnixTime
End Function

Public Function CDM_Time() As Long
    CDM_Time = ToUnixTime(Now)
End Function

Private Function CDM_Buscar(ByVal numero As Integer, ByVal Tipo As CDM_TIPO_UPDATE) As Integer

    CDM_Buscar = 0
    If CDM_ENCola Then
        For CDM_Buscar = 0 To CDM_ENCola
            If ColaComiteos(CDM_Buscar).Tipo = Tipo And ColaComiteos(CDM_Buscar).num = numero Then
                Exit For
            End If
        Next CDM_Buscar
    End If

End Function

Private Function CDM_COMIT_POP() As Boolean

Dim i As Integer

    If CDM_ENCola > 0 Then
        TmpComit = ColaComiteos(1)

        CDM_ENCola = CDM_ENCola - 1
        For i = 0 To CDM_ENCola
            ColaComiteos(i) = ColaComiteos(i + 1)
        Next i

        If CDM_ENCola < 0 Then CDM_ENCola = 0
        ReDim Preserve ColaComiteos(maxl(CDM_ENCola + 1, 1))
        CDM_COMIT_POP = True
    End If

End Function

Public Function CDM_Commit(ByVal DataOrFile As String, ByVal numero As Integer, ByVal Tipo As CDM_TIPO_UPDATE) As Integer

'Envia X archivo
' Open connection {
'  SEND:
'  -USER
'  -PASS
'  -Numero
'  -Tipo

'  RECIVE:
'  -Nuevo numero de revision
' } Close

Dim i As Integer

    i = CDM_Buscar(numero, Tipo)
    If i = 0 Then
        CDM_ENCola = CDM_ENCola + 1
        ReDim Preserve ColaComiteos(0 To CDM_ENCola + 1)
        i = CDM_ENCola
    End If

    If i > CDM_ENCola Then
        CDM_ENCola = i
        ReDim Preserve ColaComiteos(0 To CDM_ENCola + 1)
    End If

    With ColaComiteos(i)
        .active = 1
        .FilenameOrData = DataOrFile
        If FileExist(DataOrFile) Then
            .MD5 = MD5File(DataOrFile)
            .FileName = CDMFileAccess.GetFilenameFromPath(DataOrFile)
            .MIME = GetMimeType(DataOrFile)
        End If
        .num = numero
        .Tipo = Tipo
        .localeID = (GetTimer And &H7FFF)
        CDM_Commit = .localeID
    End With

    CDM_SinEnviar = True

    'CDM_Enviar ColaComiteos(i).LocaleID

    'Guardar en el archivo el nuevo numero de revision

    ' Open connection {
    '   SEND:   [UPDATE_ID][Data]
    '   RECIVE: STATUS.
    ' } Close

End Function

Public Sub CDM_DataSent(ByVal localeID As Integer)

Dim Index As Integer
Dim Value%

    CDM_SinEnviar = False
    For Index = 0 To CDM_ENCola
        If (ColaComiteos(Index).localeID = localeID And ColaComiteos(Index).active > 0) Then
            ColaComiteos(Index).active = 0
            CDMOutput "Datos enviados! #" & ColaComiteos(Index).localeID & ": " & ColaComiteos(Index).FilenameOrData
        End If
        If ColaComiteos(Index).active = 0 Then Value = Value + 1
        If ColaComiteos(Index).active > 0 Then CDM_SinEnviar = True
    Next Index

    If CDM_SinEnviar = False Then
        If frmCDM.visible Then frmCDM.Hide
        GUI_Alert "Se enviaron las actualizaciones seleccionadas", "Cerebro de mono"
        If prgRun = False Then
            DoEvents
            End
        End If
    End If

    frmCDM.pbp.max = CDM_ENCola
    frmCDM.pbp.Value = Value - 1

End Sub

Public Sub CDM_Enviar(ByVal localeID As Integer)

Dim Index As Integer
Dim Value%
Dim objHTTPRequest As New CHTTPRequest

    For Index = 0 To CDM_ENCola
        If ColaComiteos(Index).active = 0 Then Value = Value + 1
    Next Index

    frmCDM.pbp.max = CDM_ENCola
    frmCDM.pbp.Value = Value

    For Index = 0 To CDM_ENCola
        If (ColaComiteos(Index).localeID = localeID And ColaComiteos(Index).active > 0) Then

            CDMOutput "Enviando datos #" & ColaComiteos(Index).localeID & ": " & ColaComiteos(Index).FilenameOrData

            With objHTTPRequest
                .MimeBoundary = "CeReBrOdEmOnO" & Hex(GetTickCount())

                'Form fields
                Call .AddFormData("session", CDM_UserSession)
                Call .AddFormData("pass", CDM_PasswordMD5)

                Call .AddFormData("lid", Trim(CStr(ColaComiteos(Index).localeID)))
                Call .AddFormData("num", Trim(CStr(ColaComiteos(Index).num)))
                Call .AddFormData("tipo", Trim(CStr(ColaComiteos(Index).Tipo)))
                
                Select Case ColaComiteos(Index).Tipo
                Case 5, 6, 7
                    Dim Pak As clsEnpaquetado
                    Dim TmpIH As INFOHEADER
                    
                    If ColaComiteos(Index).Tipo = CDM_Upd_Sonidos Then
                        Set Pak = pakSonidos
                    ElseIf ColaComiteos(Index).Tipo = CDM_Upd_Grafico Then
                        Set Pak = pakGraficos
                    ElseIf ColaComiteos(Index).Tipo = CDM_Upd_Mapas Then
                        Set Pak = pakMapasME
                    End If
                    
                    If Not Pak Is Nothing Then
                        If Pak.Puedo_Editar(ColaComiteos(Index).num, CDM_UserPrivs, CDM_UserID) Then
                            Call .AddFormData("MD5", ColaComiteos(Index).MD5)
                            If Pak.IH_Get(ColaComiteos(Index).num, TmpIH) Then
                                Call .AddFile("file", Pak.Cabezal_GetFilenameName(ColaComiteos(Index).num), Pak.LeerRAW(ColaComiteos(Index).num), ColaComiteos(Index).MIME)
                                Call .AddFormData("data", Pak.Cabezal_GetFilenameName(ColaComiteos(Index).num))
                                .AddFormData "c1", TmpIH.complemento_1
                                .AddFormData "c2", TmpIH.complemento_2
                                .AddFormData "c3", TmpIH.complemento_3
                                .AddFormData "c4", TmpIH.complemento_4
                            End If
                        Else
                            CDMOutput "NO tenés permiso para comitear el archivo " & ColaComiteos(Index).FileName
                            GoTo skip
                        End If
                    Else
                        If FileExist(ColaComiteos(Index).FilenameOrData) Then
                            Call .AddFormData("MD5", MD5File(ColaComiteos(Index).FilenameOrData))
                            If FileExist(ColaComiteos(Index).FilenameOrData, vbNormal) = False Then
                                MsgBox "No se encontró el archivo: " & ColaComiteos(Index).FilenameOrData
                                CDMOutput "No se encontro el archivo: " & ColaComiteos(Index).FilenameOrData
                                GoTo skip
                            End If
                            Call .AddFile("file", GetFilenameFromPath(ColaComiteos(Index).FilenameOrData), GetFileQuick(ColaComiteos(Index).FilenameOrData), GetMimeType(ColaComiteos(Index).FilenameOrData))
                            Call .AddFormData("data", GetFilenameFromPath(ColaComiteos(Index).FilenameOrData))
                        Else
                            Call .AddFormData("data", ColaComiteos(Index).FilenameOrData)
                        End If
                    End If
                Case Else
                    If FileExist(ColaComiteos(Index).FilenameOrData) Then
                        Call .AddFormData("MD5", MD5File(ColaComiteos(Index).FilenameOrData))
                        If FileExist(ColaComiteos(Index).FilenameOrData, vbNormal) = False Then
                            MsgBox "No se encontró el archivo: " & ColaComiteos(Index).FilenameOrData
                            CDMOutput "No se encontro el archivo: " & ColaComiteos(Index).FilenameOrData
                            GoTo skip
                        End If
                        Call .AddFile("file", GetFilenameFromPath(ColaComiteos(Index).FilenameOrData), GetFileQuick(ColaComiteos(Index).FilenameOrData), GetMimeType(ColaComiteos(Index).FilenameOrData))
                        Call .AddFormData("data", GetFilenameFromPath(ColaComiteos(Index).FilenameOrData))
                    Else
                        Call .AddFormData("data", ColaComiteos(Index).FilenameOrData)
                    End If
                End Select
                
                

            End With

            frmCDM.webb.SendEXT "mono_commit", objHTTPRequest, ""
        End If
skip:
    Next Index

End Sub

Public Sub CDM_Enviar_Updates()

'Agregar update:

' Open connection {
'  SEND:
'  -USER
'  -PASS

'  RECIVE:
'  -UPDATE_ID
' } Close

'For each update {
' Open connection {
'   SEND:   [UPDATE_ID][Tipo][Numero][Data]
'   RECIVE: STATUS.
' } Close
'}

End Sub

Public Sub CDM_EnviarTodo()

Dim i As Integer
Dim TmpStr As String

    If frmCDM.visible = False Then frmCDM.Show
    frmCDM.picLista.visible = True
    frmCDM.list_cdm.Clear
    CDM_SinEnviar = False

    If CDM_ENCola Then
        For i = 1 To CDM_ENCola
            With ColaComiteos(i)
                Select Case .Tipo
                Case CDM_TIPO_UPDATE.CDM_Upd_Grafico
                    TmpStr = "Gráfico: " & .FilenameOrData
                Case CDM_TIPO_UPDATE.CDM_Upd_Sonidos
                    TmpStr = "Sonido: " & .FilenameOrData
                Case CDM_TIPO_UPDATE.CDM_Upd_Graficos_ind
                    TmpStr = "Indexación: " & .num
                Case CDM_TIPO_UPDATE.CDM_Upd_Tilesets
                    TmpStr = "Tileset: " & .num
                Case Else
                    TmpStr = "Tipo=" & .Tipo & "; Num=" & .num & "; Dato=" & .FilenameOrData
                End Select
                frmCDM.list_cdm.AddItem i & " - " & TmpStr
                frmCDM.list_cdm.Selected(frmCDM.list_cdm.ListCount - 1) = .active
                If .active = 1 Then CDM_SinEnviar = True
            End With
        Next i
    End If

    If CDM_SinEnviar = False Then
        If frmCDM.visible Then frmCDM.Hide
        frmCDM.picLista.visible = False
        MsgBox "No hay cosas para enviar..."
    End If

End Sub

Public Sub CDM_EnviarTodo_real()

Dim i As Integer

    If frmCDM.visible = False Then frmCDM.Show
    CDM_SinEnviar = False
    If CDM_ENCola Then
        For i = 1 To CDM_ENCola
            If ColaComiteos(i).active > 0 Then
                CDM_SinEnviar = True
                CDM_Enviar ColaComiteos(i).localeID
            End If
        Next i
    End If

    If CDM_SinEnviar = False Then
        If frmCDM.visible Then frmCDM.Hide
    End If

End Sub

Public Sub CDM_LeerListbox(list_cdm As ListBox)

Dim i As Integer
Dim TmpInt As Integer

    If CDM_ENCola Then
        For i = 0 To list_cdm.ListCount - 1
            TmpInt = val(list_cdm.list(i))
            If TmpInt <= CDM_ENCola Then
                ColaComiteos(TmpInt).active = IIf(list_cdm.Selected(i), 1, 0)
            End If
        Next i
    End If

End Sub

Public Sub CDM_Login(user As String, Password As String)
Dim POST As New CHTTPRequest
    
    CDM_Username = user
    CDM_Password = Password
    CDM_PasswordMD5 = MD5String(CDM_Password)
    
    POST.AddFormData "nick", CDM_Username
    POST.AddFormData "pass", CDM_PasswordMD5

    frmCDM.webb.SendEXT "mono_login", POST, ""
    frmCDM.webb.TryRequest
End Sub

Public Sub CDM_Update()

    CDMOutput "Solicitando datos..."
    frmCDM.webb.Send "mono_lista", , Trim$(CStr(CDM_Revision))
    frmCDM.webb.TryRequest
    
    'Manda la ultima revision
    'y recibe la lista de cosas para descargar

    'Bajar un update.

    ' Open connection {
    '  SEND:
    '  -Last Revision Local
    '  -Tipo
    '  -Numero

    '  RECIVE:
    '  -Status
    '  --Update
    ' } Close

End Sub

Public Sub CDMOutput(sText As String, Optional bIndent As Boolean = False)

    With frmCDM.txtOutput
        .SelStart = Len(.text)
        .SelLength = 0
        If sText = "_" Then
            .SelText = "_________________________________________________" & vbCrLf & vbCrLf
        Else
            If bIndent = True Then
                .SelText = "    • " & sText & vbCrLf
            Else
                .SelText = "> " & sText & vbCrLf
            End If
        End If
    End With
    LogCustomCDM sText

End Sub

' Download_File: opens up the socket connection to download the current file (CDM_Current_File)

Private Sub Download_File()

    CDMOutput "Descargando " & CDM_Current_File.FileName & "..."

    'Marce On local error resume next
        'If FileExist(CDM_TMP_PATH & CDM_Current_File.FileName, vbNormal) Then ' does the file already exists?
        Kill CDM_TMP_PATH & CDM_Current_File.FileName
        DoEvents
        'End If
    'Marce 'Marce On local error goto 0

    frmCDM.Inet.URL = CDM_Current_File.Remote
    frmCDM.Inet.Execute frmCDM.Inet.URL, "GET"

    CDMDownloading = True ' set the downloading status to true

End Sub

' Download_Start: starts the downloading process

Public Sub Download_Start(sfFiles() As CDM_File)

Dim iCount As Integer
Dim Descargar As Boolean
Dim TmpInt As Integer
Dim Reload_GRHINI As Boolean
Dim Modificado_Tilesets As Boolean
Dim ts() As String
Dim tc As Integer

Dim FileStack As New clsStack



    On Error GoTo Download_Start_Error

    frmCDM.pbp.max = maxl(UBound(sfFiles), 1)
    Dim j As New vWCDM
    GUI_SetFocus j

    For iCount = 0 To UBound(sfFiles) ' loop through our array of files
        CDMDownloading = False ' set the downloading status to false
        CDM_Current_File = sfFiles(iCount) ' set the current file
        Descargar = Len(CDM_Current_File.Remote) > 0 And Len(CDM_Current_File.FileName) > 0

        If Descargar Then
            FileStack.Push CDM_Current_File
        Else
            Select Case CDM_Current_File.Tipo
            Case CDM_TIPO_UPDATE.CDM_Upd_Graficos_ind
            
                CDMOutput "Indexando gráfico: " & CDM_Current_File.numero
                WriteVar DBPath & "Graficos.ini", "Graphics", "Grh" & CStr(CDM_Current_File.numero), CDM_Current_File.datos
                If GrhVersion < CDM_Current_File.id Then
                    GrhVersion = CDM_Current_File.id
                    WriteVar DBPath & "Graficos.ini", "INIT", "Version", GrhVersion
                End If
    
                If CDM_Current_File.numero > GrhCount Then
                    GrhCount = CDM_Current_File.numero
                    WriteVar DBPath & "Graficos.ini", "INIT", "NumGrh", GrhCount
    
                    ReDim Preserve GrhData(0 To GrhCount)
                End If
    
                indexar_from_string CDM_Current_File.numero, CDM_Current_File.datos
    
                Reload_GRHINI = True
            Case CDM_TIPO_UPDATE.CDM_Upd_Tilesets
                CDMOutput "Indexando Tileset: " & CDM_Current_File.numero
        
                If TilesetVersion < CDM_Current_File.id Then
                    TilesetVersion = CDM_Current_File.id
                    WriteVar DBPath & "Tilesets.ini", "tilesets", "ver", TilesetVersion
                End If
    
                If CDM_Current_File.numero > Tilesets_count Then
                    Tilesets_count = CDM_Current_File.numero
                    WriteVar DBPath & "Tilesets.ini", "tilesets", "num", Tilesets_count
    
                    ReDim Preserve Tilesets(0 To Tilesets_count)
                End If
                
                Procesar_tileset CDM_Current_File.numero, CDM_Current_File.datos
                Modificado_Tilesets = True
            End Select
    
            frmCDM.pbp.Value = maxl(iCount, 1)
            If CDM_Current_File.id > CDM_Revision Then
                Call WriteVar(app.Path & "\ME.ini", "CEREBRO_DE_MONO", "Revision", CDM_Current_File.id)
                CDM_Revision = CDM_Current_File.id
            End If
        End If


    Next iCount
    
    j.SetStack FileStack

    If Reload_GRHINI Then
        If CargarGraficosIni() Then
            If IndexarGraficosMemoria() Then
                CDMOutput "Graficos.ind creado..."
            Else
                CDMOutput "Error al crear Graficos.ind..."
            End If
        Else
            CDMOutput "Error al cargar Graficos.ini..."
        End If
    End If
    If Modificado_Tilesets = True Then
        GuardarTilesetsMemoria
        CargarTilesetsIni DBPath & "tilesets.ini"
        GuardarTilesetsMemoria
        CDMOutput "Tilesets guardados..."
    End If
    GUI_Alert "Se instalaron las actualizaciones!", "Cerebro de mono"

    'frmCDM.Hide
    
    CDM_Script.Lista = 1

    'Marce 'Marce 'Marce On error goto 0

Exit Sub

Download_Start_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") Download_Start of " & Erl()

End Sub

Private Function GetFileExtension(strFileName As String) As String

'Error check

    If Len(strFileName) < 3 Or InStr(1, strFileName, ".") = 0 Then
        GetFileExtension = ""
        Exit Function
    End If

    'Return
    GetFileExtension = mid$(strFileName, InStrRev(strFileName, ".") + 1)

End Function

Private Function GetMimeType(strFileName As String) As String

Dim strExtension As String

    strExtension = LCase$(GetFileExtension(strFileName))

    'Error check
    If strExtension = "" Then
        GetMimeType = "text/plain"
        Exit Function
    End If

    Select Case strExtension
    Case "bmp"
        GetMimeType = "image/bmp"

    Case "gif"
        GetMimeType = "image/gif"

    Case "jpg", "jpeg"
        GetMimeType = "image/jpeg"

    Case "swf"
        GetMimeType = "application/x-shockwave-flash"

    Case "mpg", "mpeg"
        GetMimeType = "video/mpeg"

    Case "wmv"
        GetMimeType = "video/x-ms-wmv"

    Case "avi"
        GetMimeType = "video/avi"

    Case Else
        GetMimeType = "text/plain"

    End Select

End Function

Private Function GetNextLine(ByRef sText As String, Optional ByVal reset As Boolean = False, Optional ByRef final As Boolean = False) As String

Static lLineStart As Long
Dim lLineEnd As Long
Dim lLength As Long

    If Right$(sText, 2) <> vbCrLf Then
        sText = sText & vbCrLf
    End If
    If lLineStart = 0 Then lLineStart = 1
    If reset = True Then lLineStart = 1
    lLineStart = InStr(lLineStart, sText, vbCrLf)
    lLineStart = lLineStart + 2

    If lLineStart < Len(sText) Then
        lLineEnd = InStr(lLineStart, sText, vbCrLf)
        lLength = lLineEnd - lLineStart
        GetNextLine = mid$(sText, lLineStart, lLength)
    Else
        GetNextLine = vbNullString
        lLineStart = 1
        final = True
    End If

End Function

' Script_Analyse: reads the script and gathers all our information

Public Sub Script_Analyse(texto As String)

Dim sData() As String
Dim sVars() As String
Dim sLine As String
Dim iCount As Integer
Dim iFiles As Integer
Dim bOpen As Boolean
Dim sFilename As String
Dim sRemote As String
Dim sLocal As String
Dim sDate As String
Dim MD5 As String * 32
Dim nnum        As Long
Dim nTipo       As Integer
Dim nDatos      As String
Dim nNumero     As Integer
Dim nSize       As Long
Dim sID         As Long
Dim Parchear As Boolean
Dim parcheara As Byte

Dim C1%, C2%, C3%, C4%, owner%, privs&

    CDM_Script.title = "" 'reset the scripts title
    CDM_Script.date = format(Now, "mm/dd/yyyy") ' if no date is specified, just use todays
    ReDim CDM_Script.Files(0 To 0) ' clear our the files array
    CDM_Script.Lista = 0
    CDMOutput "Analizando versiones..."
    sData = Split(texto, vbCrLf)    ' read the contents of the file then split each

    iFiles = 0 ' reset the file count
    bOpen = False ' reset structure closed
    sFilename = "" ' clear the current filename
    sRemote = "" ' clear the current remote file
    sLocal = "" ' clear the local remote path
    sDate = "" ' clear the date
    Debug.Print texto
    For iCount = 0 To UBound(sData) ' loop through each line in the array
        sLine = Trim(Replace(sData(iCount), vbTab, "")) ' remove any formatting tabs

        If sLine = "" Or sLine = vbCrLf Then GoTo skipline ' skip empty lines

        If bOpen = False Then ' if we dont have a structure open
            If Replace(sLine, " ", "") = "file{" Then ' opens a new file structure
                bOpen = True ' set the open structure variable
                GoTo skipline ' head to the next line
            Else
                ' see if we have any script specifications (ex: title or name)
                If InStr(1, sLine, ":") > 0 Then ' check if it has a colon which is used to set the variable
                    sVars = Split(sLine, ":", 2) ' split up are vars into an array
                    Select Case LCase(sVars(0))
                    Case "files"
                        If val("1" & Trim(sVars(1))) = 10 Then
                            'MsgBox "No hay actualizaciones"
                            vTextoAlerta = "Cerebro de mono:" & Chr(255) & " No hay actualizaciones!"
                            GUI_Load New vWAlert
                            Exit Sub
                        End If
                    End Select
                End If
            End If
        Else
            ' if we DO have a structure open
            If sLine = "}" Then ' this closes the structure, lets validate the information
                bOpen = False ' first close the structure variable
                If nNumero Then
                    iFiles = iFiles + 1 ' increment our file count
                    ReDim Preserve CDM_Script.Files(0 To iFiles - 1) ' rebuild the array to include our new
                    ' file structure

                    With CDM_Script.Files(iFiles - 1)
                        .FileName = sFilename ' set the filename
                        .MD5 = MD5
                        .nnum = nnum
                        .Parchear = Parchear
                        .aparchear = parcheara
                        .datos = nDatos
                        .Tipo = nTipo
                        .numero = nNumero
                        .size = nSize
                        .id = sID
                        
                        .C1 = C1
                        .C2 = C2
                        .C3 = C3
                        .C4 = C4
                        
                        'If .FileName = "" Then .FileName = Round(Rnd * GetTickCount) & ".tmp"
                        
                        .user = owner
                        .PrivsPublicos = privs
                        
                        If nTipo = CDM_TIPO_UPDATE.CDM_Upd_Grafico Or nTipo = CDM_TIPO_UPDATE.CDM_Upd_Sonidos Or nTipo = CDM_TIPO_UPDATE.CDM_Upd_Mapas Then
                            .Remote = WEBSERVER & WEBPATH & "uploads/" & sID & ".tmp"
                            Debug.Print .Remote
                        Else
                            .Remote = ""
                        End If
                    End With

                    nSize = 0
                    nDatos = ""
                    nTipo = 0
                    nNumero = 0
                    MD5 = ""
                    Parchear = False
                    parcheara = 0
                    sFilename = "" ' blank our all our vars again
                    sRemote = ""
                    sLocal = ""
                    sDate = ""
                    sID = 0
                End If
                GoTo skipline ' head to the next line

            ElseIf InStr(1, sLine, "=") >= 1 Then ' structure still open so check for variables
                sVars = Split(sLine, "=", 2) ' split the key and the variable
                Select Case LCase(sVars(0))
                Case "filename" ' filename variable
                    sFilename = Trim(sVars(1))
                    GoTo skipline ' head to the next line
                Case "md5" ' remote file variable
                    MD5 = Trim(CStr(sVars(1)))
                    GoTo skipline ' head to the next line
                Case "tipo"
                    nTipo = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "id"
                    sID = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "datos"
                    nDatos = B64.DecodeToString(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "numero"
                    nNumero = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "size"
                    nSize = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "c1"
                    C1 = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "c2"
                    C2 = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "c3"
                    C3 = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "c4"
                    C4 = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "user"
                    owner = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                Case "permisos_publicos"
                    privs = val(Trim(CStr(sVars(1))))
                    GoTo skipline
                End Select
            End If
        End If

skipline:
    Next iCount

    If iFiles >= 1 Then ' did we have any valid file structures?
        DoEvents
        Script_Summary
        Exit Sub
    End If

End Sub

Public Sub Script_Downloaded(texto As String)

Dim sData As String

    CDMOutput "Leyendo lista..."
    sData = texto
    If Left(sData, Len(psScriptToken)) <> psScriptToken Then
        Script_Invalid
        Debug.Print texto
        Exit Sub
    Else

        Script_Valid sData
        Exit Sub
    End If

End Sub

Private Sub Script_Invalid()

    CDMOutput "Error en la lista de actualizaciones"

End Sub

Private Sub Script_Summary()

    CDMOutput "Comenzando la descarga..."

    'Aplicar cambios....

    Download_Start CDM_Script.Files

End Sub

Private Sub Script_Valid(sData As String)

    CDMOutput "_"
    Script_Analyse sData

End Sub

Function URLEncode(ByVal text As String) As String

Dim i As Integer
Dim acode As Integer
Dim Char As String

    URLEncode = text

    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(mid$(URLEncode, i, 1))
        Select Case acode
        Case 48 To 57, 65 To 90, 97 To 122
            ' don't touch alphanumeric chars
        Case 32
            ' replace space with "+"
            Mid$(URLEncode, i, 1) = "+"
        Case Else
            ' replace punctuation chars with "%hex"
            URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & mid$(URLEncode, i + 1)
        End Select
    Next i

End Function

':) Ulli's VB Code Formatter V2.24.17 (2010-Oct-22 00:50)  Decl: 66  Code: 701  Total: 767 Lines
':) CommentOnly: 67 (8,7%)  Commented: 45 (5,9%)  Filled: 596 (77,7%)  Empty: 171 (22,3%)  Max Logic Depth: 7
