Attribute VB_Name = "modZLib"
'ARCHIVO COMPARTIDO.

'                  ____________________________________________
'                 /_____/  http://www.arduz.com.ar/ao/   \_____\
'                //            ____   ____   _    _ _____      \\
'               //       /\   |  __ \|  __ \| |  | |___  /      \\
'              //       /  \  | |__) | |  | | |  | |  / /        \\
'             //       / /\ \ |  _  /| |  | | |  | | / /   II     \\
'            //       / ____ \| | \ \| |__| | |__| |/ /__          \\
'           / \_____ /_/    \_\_|  \_\_____/ \____//_____|_________/ \
'           \________________________________________________________/
'           MZEngine DX8             Manejador de archivos de recursos
'           Hecho por Menduz <3
'           TODO:   Pasarlo a C++ para agilizarlo,
'                   ya que vb es una mierda lenta

Option Explicit
'C:\PC VIEJA\aonuevo\ClienteDX8\Datos\mapas\
Public Type INFOHEADER
    CRC                     As Long
    cript                   As Byte
    lngFileSizeUncompressed As Long

    originalname            As String * 32

    file_type               As Integer

    compress                As Byte

    size_compressed         As Long
    flags                   As Long

    EmpiezaByte             As Long

    future_expansion3       As Long
    future_expansion4       As Long
    future_expansion5       As Long

    futurei_e1              As Integer
    futurei_e2              As Integer
    futurei_e3              As Integer
    complemento_1           As Integer 'ID DE EL 1er ITEM COMPLEMENTARIO A LA TEXTURA
    complemento_2           As Integer 'ID DE EL 2do ITEM COMPLEMENTARIO A LA TEXTURA
End Type

Public Enum eTiposRecursos
    rDesconocido = 0
    rPng = 1
    rBmp = 2
    rJpg = 3
    rInit = 4
    rMapData = 5
End Enum

Public Enum e_resource_file
    rMapas = 0
    rGUI = 1
    rGrh = 2
End Enum

#If False Then
Private rDesconocido, rPng, rBmp, rJpg, rInit, rMapData, rMapas, rGUI, rGrh
#End If

Public Const header_s As String * 16 = "MZEngineSyngler§"
Private Const header_b As String * 16 = "MZEngineBinarir§"

Private Declare Function compress Lib "zlib.dll" _
        (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" _
        (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef dest As Any, ByRef Source As Any, ByVal ByteCount As Long)

Private Declare Function CRC32 Lib "MZEngine.dll" Alias "CRC_BA" _
        (ByRef bArray As Byte, ByVal lLen As Long, ByVal lCrc As Long) As Long

Private Declare Sub Xor_Bytes Lib "MZEngine.dll" Alias "Xor_Bytes_BA" _
        (ByRef FirstByte As Byte, ByVal lenght As Long, ByVal code As Byte, ByVal CryptKey As Byte)

Private Declare Sub MDFile Lib "aamd532.dll" _
        (ByVal f As String, ByVal r As String)

Private Declare Sub MDStringFix Lib "aamd532.dll" _
        (ByVal f As String, ByVal t As Long, ByVal r As String)
        
Private Declare Function CreateStreamOnHGlobal Lib "ole32" _
    (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
    
Private Declare Function OleLoadPicture Lib "olepro32" _
    (pstream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
    
Private Declare Function CLSIDFromString Lib "ole32" _
    (ByVal lpsz As Any, pclsid As Any) As Long
    
Private Declare Function GlobalAlloc Lib "kernel32" _
    (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    
Private Declare Function GlobalLock Lib "kernel32" _
    (ByVal hMem As Long) As Long
    
Private Declare Function GlobalUnlock Lib "kernel32" _
    (ByVal hMem As Long) As Long
    
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDest As Any, pSource As Any, ByVal dwLength As Long)


Private Const CryptKey      As Byte = 108
Private Const CryptKeyL     As Long = 984362498

Private Path_res            As String
Private Const FN_Mapas      As String = "Mapas.TDS"
Private Const FN_Grh        As String = "Graficos.TDS"
Private Const FN_GUI        As String = "Interface.TDS"

Public last_file_ext        As INFOHEADER
Public Extraidox            As Boolean

Private CabezalInterface()  As INFOHEADER
Private CabezalGraficos()   As INFOHEADER
Private CabezalMapas()      As INFOHEADER

Private UltimoBInterface    As Long
Private UltimoBGraficos     As Long
Private UltimoBMapas        As Long

Private CantidadInterface   As Integer
Private CantidadGraficos    As Integer
Private CantidadMapas       As Integer

Private Const bTRUE         As Byte = 255
Private Const bFALSE        As Byte = 0

Private Const Min_Offset    As Integer = 500 ' El "cacho" de slots libres del array para agregar archivos
Private Const Max_Int_Val   As Integer = 32767 ' (2 ^ 16) / 2 - 1


Public Function PictureFromByteStream(b() As Byte) As IPicture
'código roñosooo!!!!
    Dim LowerBound  As Long
    Dim ByteCount   As Long
    Dim hMem        As Long
    Dim lpMem       As Long
    Dim IID_IPicture(15)
    Dim istm        As stdole.IUnknown

    On Error GoTo Err_Init
    If UBound(b, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                  Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:
    If Err.number = 9 Then
        'Uninitialized array
        LogError "PictureFromByteStream->BA empty"
    Else
        LogError "PictureFromByteStream->(" & Err.number & ") " & Err.Description
    End If
End Function


Private Sub AddItem2Array1D(ByRef VarArray As Variant, ByVal VarValue As Variant)

Dim i  As Long
Dim iVarType As Integer

    iVarType = VarType(VarArray) - 8192
    i = UBound(VarArray)

    Select Case iVarType

    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte

        If VarArray(0) = 0 Then
            i = 0
        Else
            i = i + 1
        End If

    Case vbDate

        If VarArray(0) = "00:00:00" Then
            i = 0
        Else
            i = i + 1
        End If

    Case vbString

        If VarArray(0) = vbNullString Then
            i = 0
        Else
            i = i + 1
        End If

    Case vbBoolean

        If VarArray(0) = False Then
            i = 0
        Else
            i = i + 1
        End If

    Case Else

    End Select

    ReDim Preserve VarArray(i)
    VarArray(i) = VarValue

End Sub

Public Function AllFilesInFolders(ByRef sFolderPath As String, Optional ByRef pattern As String = "*.*") As String()

Dim sTemp As String
Dim sDirIn As String
Dim i As Integer, j As Integer
Dim sFilelist() As String

    ReDim sFilelist(0) As String
Dim slist() As String

    sDirIn = sFolderPath
    'If Not (Right$(sDirIn, 1) = "\") Then sDirIn = sDirIn & "\"
    If Not (Right$(sDirIn, 1) = "\") Then
        sDirIn = sDirIn & "\"
    End If

    On Error Resume Next
        slist = Split(pattern, ";")
        For i = 0 To UBound(slist)
            sTemp = dir$(sDirIn & slist(i))
            Do While LenB(sTemp) <> 0
                'If (Len(sTemp)) Then _
                     AddItem2Array1D sFilelist(), sTemp
                If (Len(sTemp)) Then
                    AddItem2Array1D sFilelist(), sTemp
                End If
                sTemp = dir
            Loop
        Next i
        AllFilesInFolders = sFilelist

    On Error GoTo 0

End Function

Public Function Bin_Create_From_Folder(ByVal rFile_type As e_resource_file, ByRef folder As String, ByRef output_folder As String) As Boolean

'On Error GoTo errh

Dim handle          As Integer

Dim abierto         As Byte

Dim Nueva_Cantidad  As Integer
Dim Ultimo_Byte     As Long
Dim Cabezal()       As INFOHEADER
Dim InfoHead        As INFOHEADER
Dim cabezal_ptr     As Integer

Dim SourceData()    As Byte
Dim File_List()     As String

Dim i               As Integer
Dim cantidad_array  As Integer

Dim max_cantidad    As Integer
Dim TmpInt          As Integer

Dim tmplng          As Long
Dim int_list()      As Integer

Dim handleB         As Integer
Dim FileName        As String

Dim new_file        As String
Dim asd             As String * 16
Dim necesita_hacer  As Byte
Dim tmpbn As Double
    File_List = AllFilesInFolders(folder, Bin_Rs_Get_File_Pattern(rFile_type))

    cantidad_array = UBound(File_List)
    ReDim int_list(cantidad_array)

    For i = 0 To cantidad_array
    tmpbn = val(Split(File_List(i), ".", 2)(0))
        If tmpbn <= Max_Int_Val Then
        TmpInt = tmpbn

        'If max_cantidad < tmpint Then max_cantidad = tmpint
        If max_cantidad < TmpInt Then
            max_cantidad = TmpInt
        End If
        int_list(i) = TmpInt
        End If
    Next i

    tmplng = max_cantidad + Min_Offset
    'If tmplng > Max_Int_Val Then tmplng = Max_Int_Val
    If tmplng > Max_Int_Val Then
        tmplng = Max_Int_Val
    End If
    Nueva_Cantidad = tmplng

    ReDim Cabezal(Nueva_Cantidad)

    'If Len(output_folder) = 0 Then output_folder = App.path
    If Len(output_folder) = 0 Then
        output_folder = app.Path
    End If
    'If Right$(output_folder, 1) <> "\" Then output_folder = output_folder & "\"
    If Right$(output_folder, 1) <> "\" Then
        output_folder = output_folder & "\"
    End If
    
    Select Case rFile_type
    Case e_resource_file.rGUI
        FileName = output_folder & FN_GUI
    Case e_resource_file.rGrh
        FileName = output_folder & FN_Grh
    Case e_resource_file.rMapas
        FileName = output_folder & FN_Mapas
    End Select

    'If (Dir$(filename, vbNormal) <> "") Then Kill filename
    If (dir$(FileName, vbNormal) <> "") Then
        Kill FileName
    End If
    handleB = FreeFile()
    Open FileName For Binary Access Read Write As handleB
    Seek handleB, 1

    Ultimo_Byte = 0
    Put handleB, , header_b
    Put handleB, , Nueva_Cantidad
    Put handleB, , Ultimo_Byte
    Put handleB, , Cabezal

    For i = 0 To cantidad_array
        If int_list(i) > 0 Then
            new_file = folder & File_List(i)
            handle = FreeFile()
            Open new_file For Binary Access Read Lock Write As handle: abierto = bTRUE
            Get handle, , asd
            If StrComp(asd, header_s, vbTextCompare) Then
                necesita_hacer = bTRUE
            Else
                Get handle, , InfoHead

                ReDim SourceData(InfoHead.size_compressed) As Byte

                Get handle, , SourceData()

'                If InfoHead.size_compressed > 1024 Then
'                    InfoHead.crc = CRC32(SourceData(0), 1024, 0)
'                Else
'                    InfoHead.crc = CRC32(SourceData(0), InfoHead.size_compressed - 1, 0)
'                End If
            End If
            Close handle: abierto = bFALSE

            If necesita_hacer Then
                Resource_Generate_IH new_file, InfoHead, SourceData
            End If

            InfoHead.EmpiezaByte = seek(handleB)

            Cabezal(int_list(i)) = InfoHead

            Put handleB, , SourceData
            Ultimo_Byte = Ultimo_Byte + InfoHead.size_compressed
            '#If CRIPTER Then
            Debug.Print "Push ("; int_list(i); ") - "; File_List(i); " - CRC:"; Hex(InfoHead.CRC); InfoHead.size_compressed; "- ptr:"; Hex$(Ultimo_Byte)
            '#End If
        End If
    Next i

    Seek handleB, 1

    Put handleB, , header_b
    Put handleB, , Nueva_Cantidad
    Put handleB, , Ultimo_Byte
    Put handleB, , Cabezal
    Close handleB

    Bin_Create_From_Folder = True

Exit Function

errh:
    LogError "Error en el archivo de recursos """ & FileName & """ Err:" & Err.number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

'///////////////////////////////////////////////////////////////////////////
'///////////////////////PASAR A C++ PARA GANAR VELOCIDAD!///////////////////
'///////////////////////////////////////////////////////////////////////////

'Private Sub Xor_Bytes(ByRef ByteArray() As Byte, ByVal code As Byte)
'    Dim i As Integer
'    For i = 0 To UBound(ByteArray)
'        ByteArray(i) = code Xor (ByteArray(i) Xor CryptKey)
'    Next
'End Sub
'//Public Declare sub Xor_Bytes Lib "MZEngine.dll" Alias "Xor_Bytes_BA" (ByRef FirstByte As Byte, ByVal Lenght As Long, ByVal code As byte, ByVal CryptKey As byte)

'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////

Public Sub Bin_Load_Headers(ByVal Path As String)

Dim cantidad    As Integer
Dim handle      As Integer
Dim t_str       As String * 16
Dim abierto     As Byte
Dim FileName    As String

    Path_res = Path
    handle = FreeFile()
    FileName = Path_res & FN_GUI
    Open FileName For Binary As handle: abierto = bTRUE
    Get handle, 1, t_str
    'If StrComp(t_str, header_b, vbTextCompare) Then GoTo erra
    If StrComp(t_str, header_b, vbTextCompare) Then
        GoTo erra
    End If

    Get handle, , CantidadInterface
    Get handle, , UltimoBInterface

    ReDim CabezalInterface(CantidadInterface)

    Get handle, , CabezalInterface 'Get handle, UltimoBInterface, CabezalInterface
    Close handle: abierto = bFALSE

    handle = FreeFile()
    FileName = Path_res & FN_Mapas
    Open FileName For Binary As handle: abierto = bTRUE
    Get handle, 1, t_str
    'If StrComp(t_str, header_b, vbTextCompare) Then GoTo erra
    If StrComp(t_str, header_b, vbTextCompare) Then
        GoTo erra
    End If

    Get handle, , CantidadMapas
    Get handle, , UltimoBMapas

    ReDim CabezalMapas(CantidadMapas)

    Get handle, , CabezalMapas 'Get handle, UltimoBMapas, CabezalMapas
    Close handle: abierto = bFALSE

    handle = FreeFile()
    FileName = Path_res & FN_Grh
    Open FileName For Binary As handle: abierto = bTRUE
    Get handle, 1, t_str
    'If StrComp(t_str, header_b, vbTextCompare) Then GoTo erra
    If StrComp(t_str, header_b, vbTextCompare) Then
        GoTo erra
    End If

    Get handle, , CantidadGraficos
    Get handle, , UltimoBGraficos

    ReDim CabezalGraficos(CantidadGraficos)

    Get handle, , CabezalGraficos 'Get handle, UltimoBGraficos, CabezalGraficos
    Close handle: abierto = bFALSE

Exit Sub

erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If
    LogError "El archivo : """ & FileName & """ no es un archivo de recursos valido."
    'End

End Sub

Public Function Bin_Resource_Add_To_Listbox(ByVal rFile_type As e_resource_file, ByRef List As ListBox) As Boolean
On Error Resume Next
Dim InfoHead    As INFOHEADER
Dim abierto     As Byte
Dim i           As Integer

    List.Clear
    Select Case rFile_type
    Case e_resource_file.rGUI
        For i = 0 To CantidadInterface
            InfoHead = CabezalInterface(i)
            'If InfoHead.size_compressed Then _
                List.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript)))
            If InfoHead.size_compressed Then
                List.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript))) & " - CRC:" & Hex$(InfoHead.CRC) & " - " & Round(InfoHead.size_compressed / 1024, 1) & "KB"
            End If
        Next i
    Case e_resource_file.rGrh
        For i = 0 To CantidadGraficos
            InfoHead = CabezalGraficos(i)
            'If InfoHead.size_compressed Then _
                List.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript)))
            If InfoHead.size_compressed Then
                List.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript))) & " - CRC:" & Hex$(InfoHead.CRC) & " - " & Round(InfoHead.size_compressed / 1024, 1) & "KB"
            End If
        Next i
    Case e_resource_file.rMapas
        For i = 0 To CantidadMapas
            InfoHead = CabezalMapas(i)
            'If InfoHead.size_compressed Then _
                List.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript)))
            If InfoHead.size_compressed Then
                List.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript))) & " - CRC:" & Hex$(InfoHead.CRC) & " - " & Round(InfoHead.size_compressed / 1024, 1) & "KB"
            End If
        Next i
    End Select

End Function

Public Function Bin_Resource_Extract(ByRef nro As Integer, ByVal rFile_type As e_resource_file, ByRef dest As String) As Boolean

'On Error GoTo errh

Dim SourceData() As Byte
Dim handle%

    handle = FreeFile()

    'If (Dir$(dest, vbNormal) <> "") Then Kill dest
    If (dir$(dest, vbNormal) <> "") Then
        Kill dest
    End If

    If Bin_Resource_Get(nro, SourceData, rFile_type) Then
        Open dest For Binary Access Read Write As handle
        Put handle, , SourceData()
        Close handle
        Bin_Resource_Extract = True
    End If

errh:

End Function

Public Function Bin_Resource_Load_Picture(ByVal nro As Integer, ByVal rFile_type As e_resource_file) As IPicture

On Error GoTo errh

Dim SourceData()    As Byte
Dim LowerBound      As Long
Dim ByteCount       As Long
Dim hMem            As Long
Dim lpMem           As Long
Dim istm            As stdole.IUnknown
Dim IID_IPicture(15) ' no sabe no contesta

    If Bin_Resource_Get(nro, SourceData, rFile_type) Then
        LowerBound = LBound(SourceData)
        ByteCount = (UBound(SourceData) - LowerBound) + 1
        hMem = GlobalAlloc(&H2, ByteCount)
        If hMem <> 0 Then
            lpMem = GlobalLock(hMem)
            If lpMem <> 0 Then
                MoveMemory ByVal lpMem, SourceData(LowerBound), ByteCount
                Call GlobalUnlock(hMem)
                If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                    If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                      Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), Bin_Resource_Load_Picture)
                    End If
                End If
            End If
        End If
    End If
    Exit Function

errh:
    If Err.number = 9 Then
        LogError "Bin_Resource_Load_Picture->BA empty"
    Else
        LogError "Bin_Resource_Load_Picture->(" & Err.number & ") " & Err.Description
    End If

End Function

Public Function Bin_Resource_Get(ByRef nFile As Integer, ByRef Data() As Byte, ByVal rFile_type As e_resource_file) As Boolean
On Error GoTo errh

Dim handle As Integer
Dim SourceData() As Byte
Dim InfoHead As INFOHEADER
Dim abierto As Byte
Dim FileName As String

    Select Case rFile_type
    Case e_resource_file.rGUI
        FileName = Path_res & FN_GUI
        If CantidadInterface >= nFile Then
            InfoHead = CabezalInterface(nFile)
        Else
            Bin_Resource_Get = False
            Exit Function
        End If
    Case e_resource_file.rGrh
        FileName = Path_res & FN_Grh
        If CantidadGraficos >= nFile Then
            InfoHead = CabezalGraficos(nFile)
        Else
            Bin_Resource_Get = False
            Exit Function
        End If
    Case e_resource_file.rMapas
        FileName = Path_res & FN_Mapas
        If CantidadMapas >= nFile Then
            InfoHead = CabezalMapas(nFile)
        Else
            Bin_Resource_Get = False
            Exit Function
        End If
    End Select
    
    If nFile = 0 Then
        Bin_Resource_Get = False
        Exit Function
    End If

    Extraidox = False

    handle = FreeFile()
    
#If esCLIENTE = 1 Or esME = 1 Then
    If InfoHead.EmpiezaByte = 0 Then
        LogError nFile & " NO ESTA EN ENPAQEUTADO " & rFile_type
    Else
        If rFile_type = e_resource_file.rMapas Then
            Call map_load_from(FileName, InfoHead.EmpiezaByte, nFile) 'Cargamos el mapa desde el archivo binario, nos ahorramos un par de accesos al puto disco ;)
        Else
        
            Open FileName For Binary Access Read Lock Write As handle: abierto = bTRUE
                Seek handle, InfoHead.EmpiezaByte ' movemos el puntero de handle a EmpiezaByte
                ReDim SourceData(InfoHead.size_compressed) As Byte
                Get handle, , SourceData()

                If InfoHead.compress = 1 Then
                    Decompress_Data SourceData(), InfoHead.lngFileSizeUncompressed Xor CryptKeyL Xor InfoHead.cript
                End If
                
                Data = SourceData
            Close handle: abierto = bFALSE
        End If
    End If
#Else
    Open FileName For Binary Access Read Lock Write As handle: abierto = bTRUE
            Debug.Print "[EXTRAYENDO]"


        Debug.Print " Offset file:"; InfoHead.EmpiezaByte
        Debug.Print " Tamaño:"; InfoHead.size_compressed
        Debug.Print " Comprimido: "; CBool(InfoHead.compress)
        Debug.Print ""
        
        Seek handle, InfoHead.EmpiezaByte ' movemos el puntero de handle a EmpiezaByte
        ReDim SourceData(InfoHead.size_compressed) As Byte
        Get handle, , SourceData()
        
        If InfoHead.compress = 1 Then
            Decompress_Data SourceData(), InfoHead.lngFileSizeUncompressed Xor CryptKeyL Xor InfoHead.cript
        End If
        
        Data = SourceData
    Close handle: abierto = bFALSE
#End If

    last_file_ext = InfoHead
    Bin_Resource_Get = True

Exit Function

errh:
    LogError "Error en el archivo de recursos """ & FileName & """ Err:" & Err.number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

Public Function Bin_Resource_Get_crc(ByRef nFile As Integer, ByVal rFile_type As e_resource_file) As Long

' esta func se puede usar para el parcheo
' return 0 cuando es invalido o error

    On Error Resume Next
        Bin_Resource_Get_crc = &H0
        Select Case rFile_type
        Case e_resource_file.rGUI
            'If UBound(CabezalInterface) >= nFile Then _
                       Bin_Resource_Get_crc = CabezalInterface(nFile).crc
            If UBound(CabezalInterface) >= nFile Then
                Bin_Resource_Get_crc = CabezalInterface(nFile).CRC
            End If
        Case e_resource_file.rGrh
            'If UBound(CabezalGraficos) >= nFile Then _
                       Bin_Resource_Get_crc = CabezalGraficos(nFile).crc
            If UBound(CabezalGraficos) >= nFile Then
                Bin_Resource_Get_crc = CabezalGraficos(nFile).CRC
            End If
        Case e_resource_file.rMapas
            'If UBound(CabezalMapas) >= nFile Then _
                       Bin_Resource_Get_crc = CabezalMapas(nFile).crc
            If UBound(CabezalMapas) >= nFile Then
                Bin_Resource_Get_crc = CabezalMapas(nFile).CRC
            End If
        End Select

End Function

Public Function Bin_Resource_Get_Raw(ByRef nFile As Integer, ByVal rFile_type As e_resource_file) As String

Dim SourceData() As Byte

    If Bin_Resource_Get(nFile, SourceData, rFile_type) Then
        Bin_Resource_Get_Raw = StrConv(SourceData, vbUnicode)
    Else
        Bin_Resource_Get_Raw = vbNullString
    End If

End Function

Public Function Bin_Resource_GET_IH(ByRef nFile As Integer, ByRef InfoHead As INFOHEADER, ByVal rFile_type As e_resource_file) As Boolean
    Select Case rFile_type
        Case e_resource_file.rGUI
            If CantidadInterface < nFile Then
                Exit Function
            Else
                InfoHead = CabezalInterface(nFile)
            End If
        Case e_resource_file.rGrh
            If CantidadGraficos < nFile Then
                Exit Function
            Else
                InfoHead = CabezalGraficos(nFile)
            End If
        Case e_resource_file.rMapas
            If CantidadMapas < nFile Then
                Exit Function
            Else
                InfoHead = CabezalMapas(nFile)
            End If
    End Select
    Bin_Resource_GET_IH = True
End Function


Public Function Bin_Resource_MOD_IH(ByRef nFile As Integer, ByRef InfoHead As INFOHEADER, ByVal rFile_type As e_resource_file) As Boolean

'On Error GoTo errh

Dim handle          As Integer

Dim abierto         As Byte

Dim FileName        As String
Dim fhp             As Long
Dim TIH As INFOHEADER

    
    
    If Bin_Resource_GET_IH(nFile, TIH, rFile_type) = False Then Exit Function
    
    If TIH.lngFileSizeUncompressed = 0 Then Exit Function
    
    InfoHead.compress = TIH.compress
    InfoHead.cript = TIH.cript
    InfoHead.EmpiezaByte = TIH.EmpiezaByte
    InfoHead.size_compressed = TIH.size_compressed
    InfoHead.lngFileSizeUncompressed = TIH.lngFileSizeUncompressed
    InfoHead.originalname = TIH.originalname
    'InfoHead.originalname = Xor_String(TIH.originalname, InfoHead.cript)
    InfoHead.file_type = TIH.file_type
    InfoHead.CRC = (TIH.CRC + 1) Mod &HFFFFFFF

    Select Case rFile_type
    Case e_resource_file.rGUI:
        FileName = Path_res & FN_GUI
        CabezalInterface(nFile) = InfoHead
    Case e_resource_file.rGrh:
        FileName = Path_res & FN_Grh
        CabezalGraficos(nFile) = InfoHead
    Case e_resource_file.rMapas:
        FileName = Path_res & FN_Mapas
         CabezalMapas(nFile) = InfoHead
    End Select
    
    handle = FreeFile()
    Debug.Print "[MODIFICANDO:" & nFile & " de " & rFile_type & "]"
    
    fhp = 23 + CLng(nFile) * Len(InfoHead) ' muejeje
    Open FileName For Binary Access Read Write As handle: abierto = bTRUE
        Put handle, fhp, InfoHead
    Close handle: abierto = bFALSE
    
    Debug.Print " Offset head:"; fhp
    Debug.Print " MODIFICADO OK."
    Debug.Print ""
    Bin_Resource_MOD_IH = True
Exit Function

errh:
    LogError "Error en mod el archivo de recursos """ & FileName & """ Err:" & Err.number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function



Public Function Bin_Resource_Patch(ByRef nFile As Integer, ByRef new_file As String, ByVal rFile_type As e_resource_file, Optional ByVal CRC As Long = -1) As Boolean

'On Error GoTo errh

Dim handle          As Integer

Dim InfoHead        As INFOHEADER
Dim abierto         As Byte
Dim file_len        As Long

Dim Resize_Header   As Byte
Dim tmp_s           As String * 16
Dim tmpcrc          As Long
Dim necesita_hacer  As Byte
Dim Ultimo_Byte     As Long

Dim es_igual_viejo  As Byte

Dim SourceData()    As Byte

Dim Nueva_Cantidad  As Integer

Dim FileName        As String
Dim fhp             As Long

'    Resize_Header = bFALSE
'    necesita_hacer = bFALSE
'    es_igual_viejo = bFALSE

    file_len = FileLen(new_file)

    'If Not LenB(Dir$(path & filename, vbNormal)) Then GoTo errh
    'If LenB(Dir$(new_file, vbNormal)) Then
    '    GoTo errh
    'End If

    handle = FreeFile

    Open new_file For Binary Access Read Lock Write As handle: abierto = bTRUE
    Get handle, , tmp_s
    If StrComp(tmp_s, header_s, vbTextCompare) Then ' StrComp es MUCHO más rápido que If Str1 = Str2 Then
        necesita_hacer = bTRUE
    Else
        Get handle, , InfoHead

        ReDim SourceData(InfoHead.size_compressed) As Byte

        Get handle, , SourceData()

'        If InfoHead.size_compressed > 1024 Then
'            InfoHead.crc = CRC32(SourceData(0), 1024, 0)
'        Else
'            InfoHead.crc = CRC32(SourceData(0), InfoHead.size_compressed - 1, 0)
'        End If
    End If
    Close handle: abierto = bFALSE

    If necesita_hacer Then
        Resource_Generate_IH new_file, InfoHead, SourceData
    End If
    
If InfoHead.size_compressed = 0 Then
MsgBox "Error al parchear."
End If

    Select Case rFile_type
    Case e_resource_file.rGUI
        FileName = Path_res & FN_GUI
        If CantidadInterface < nFile Then
            ReDim Preserve CabezalInterface(nFile)
            Resize_Header = bTRUE
            Nueva_Cantidad = CantidadInterface + 1
        Else
            'es_igual_viejo = CabezalInterface(nFile).crc = InfoHead.crc
            Nueva_Cantidad = CantidadInterface
        End If
        InfoHead.EmpiezaByte = UltimoBInterface + 2
        CabezalInterface(nFile) = InfoHead
        
        Ultimo_Byte = InfoHead.EmpiezaByte + InfoHead.size_compressed
        UltimoBInterface = Ultimo_Byte
        
        
    Case e_resource_file.rGrh
        FileName = Path_res & FN_Grh
        If CantidadGraficos < nFile Then
            ReDim Preserve CabezalGraficos(nFile)
            Resize_Header = bTRUE
            Nueva_Cantidad = CantidadGraficos + 1
        Else
            'es_igual_viejo = CabezalGraficos(nFile).crc = InfoHead.crc
            Nueva_Cantidad = CantidadGraficos
        End If
        
        InfoHead.EmpiezaByte = UltimoBGraficos + 2
        CabezalGraficos(nFile) = InfoHead
        
        Ultimo_Byte = InfoHead.EmpiezaByte + InfoHead.size_compressed
        UltimoBGraficos = Ultimo_Byte
        
        
    Case e_resource_file.rMapas
        FileName = Path_res & FN_Mapas
        If CantidadMapas < nFile Then
            ReDim Preserve CabezalMapas(nFile)
            Resize_Header = bTRUE
            Nueva_Cantidad = CantidadMapas + 1
        Else
            'es_igual_viejo = CabezalMapas(nFile).crc = InfoHead.crc
            Nueva_Cantidad = CantidadMapas
        End If
        
        InfoHead.EmpiezaByte = UltimoBMapas + 2
        CabezalMapas(nFile) = InfoHead

        Ultimo_Byte = InfoHead.EmpiezaByte + InfoHead.size_compressed
        UltimoBMapas = Ultimo_Byte
    End Select
    
    If es_igual_viejo Then
        Bin_Resource_Patch = True
        Exit Function
    End If
    
    If CRC <> -1 Then InfoHead.CRC = CRC
    
    handle = FreeFile()
    Open FileName For Binary Access Read Write As handle: abierto = bTRUE
'        Seek handle, 1              ' movemos el puntero de handle a UltimoBMapas
'
'        Put handle, , header_b
'        Put handle, , Nueva_Cantidad
'        Put handle, , Ultimo_Byte
'        fhp = Seek(handle) + CLng(nFile) * Len(InfoHead) ' muejeje
'        Put handle, fhp, InfoHead
'        Put handle, InfoHead.EmpiezaByte, SourceData

        Debug.Print "[PARCHEANDO:" & FileName & "]"

        Seek handle, LOF(handle) + 1
        InfoHead.EmpiezaByte = seek(handle)
        Debug.Print " Offset file:"; InfoHead.EmpiezaByte
        Put handle, , SourceData
        Ultimo_Byte = seek(handle)
        
        Seek handle, 1
        Put handle, , header_b
        Put handle, , Nueva_Cantidad
        Put handle, , Ultimo_Byte
        fhp = seek(handle) + CLng(nFile) * Len(InfoHead) ' muejeje
        Put handle, fhp, InfoHead
        
        Debug.Print " Offset head:"; fhp
        
        Debug.Print " Tamaño:"; InfoHead.size_compressed
        Debug.Print " Comprimido: "; CBool(InfoHead.compress)
        Debug.Print " Total de archivos:"; Nueva_Cantidad
        Debug.Print " Ultimo byte:"; Ultimo_Byte
        Debug.Print " PARCHEADO OK."
        Debug.Print ""
        
        'Seek handle,  CLng(Ultimo_Byte - InfoHead.size_compressed)' movemos el puntero de handle a UltimoBMapas
        'Put handle, InfoHead.EmpiezaByte, SourceData
    
        Bin_Resource_Patch = True

    Close handle: abierto = bFALSE

    Select Case rFile_type
        Case e_resource_file.rGUI
            CabezalInterface(nFile) = InfoHead
            UltimoBInterface = Ultimo_Byte
        Case e_resource_file.rGrh
            CabezalGraficos(nFile) = InfoHead
            UltimoBGraficos = Ultimo_Byte
        Case e_resource_file.rMapas
            CabezalMapas(nFile) = InfoHead
            UltimoBMapas = Ultimo_Byte
    End Select

Exit Function

errh:
    LogError "Error en el archivo de recursos """ & FileName & """ Err:" & Err.number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

Public Function Bin_Rs_Get_File_Pattern(ByVal rFile_type As e_resource_file) As String

    Select Case rFile_type
    Case e_resource_file.rGrh
        Bin_Rs_Get_File_Pattern = "*.bmp;*.png;*.dds;*.tga;*.mzg"
    Case e_resource_file.rGUI
        Bin_Rs_Get_File_Pattern = "*.jpg;*.jpeg"
    Case e_resource_file.rMapas
        Bin_Rs_Get_File_Pattern = "*.am"
    End Select

End Function

Public Sub Compress_Data(ByRef Data() As Byte)

Dim Dimensions As Long
Dim DimBuffer As Long
Dim BufTemp() As Byte
Dim BufTemp2() As Byte
Dim loopc As Long

    Dimensions = UBound(Data)

    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)

    compress BufTemp(0), DimBuffer, Data(0), Dimensions

    Erase Data

    ReDim Preserve BufTemp(DimBuffer - 1)

    Data = BufTemp

    Erase BufTemp

    Data(0) = Data(0) Xor CryptKey Xor Data(1)

End Sub

Public Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)

Dim BufTemp() As Byte

    ReDim BufTemp(OrigSize - 1)

    Data(0) = Data(0) Xor CryptKey Xor Data(1)

    UnCompress BufTemp(0), OrigSize, Data(0), UBound(Data) + 1

    ReDim Data(OrigSize - 1)

    Data = BufTemp

    Erase BufTemp

End Sub

Public Function MD5File(f As String) As String

' compute MD5 digest on o given file, returning the result

Dim r As String * 32

    r = Space(32)
    MDFile f, r
    MD5File = r

End Function

Public Function MD5String(p As String) As String

' compute MD5 digest on a given string, returning the result

Dim r As String * 32, t As Long

    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r

End Function

Public Sub Resource_Convert(ByRef sourcepath As String, ByRef Path As String, ByRef FileName As String, Optional ByVal arg1 As Integer = 0)

'On Error GoTo errh

Dim handle As Integer
Dim SourceData() As Byte
Dim InfoHead As INFOHEADER
Dim abierto As Byte
Dim tmpcrc As Long
Dim ts As String * 3
Dim freem%
Dim tmpl&

    'If Right$(path, 1) <> "\" Then path = path & "\"
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    If (dir$(Path & FileName, vbNormal) <> "") Then
        Kill Path & FileName
    End If
    If (dir$(sourcepath, vbNormal) <> "") Then
        Resource_Generate_IH sourcepath, InfoHead, SourceData
        handle = FreeFile
        Open Path & FileName For Binary Access Read Write As handle
        Put handle, , header_s
        Put handle, , InfoHead
        Put handle, , SourceData()
        Close handle
        Debug.Print Path & FileName & " PACKED_OK - C:" & Hex$(InfoHead.CRC) & " - COMP:" & CStr(CBool(InfoHead.compress))
        Erase SourceData()
    Else
        LogError "Error en el archivo de a comprimir """ & FileName & """ - El archivo No existe."
    End If

Exit Sub

errh:
    LogError "Error en el archivo de recursos """ & FileName & """"

End Sub

Public Function Resource_Extract(ByRef Path As String, ByRef FileName As String, ByRef dest As String) As Boolean

'On Error GoTo errh

Dim SourceData() As Byte
Dim handle%

    handle = FreeFile()

    Resource_Get Path, FileName, SourceData

    'If (Dir$(dest, vbNormal) <> "") Then Kill dest
    If (dir$(dest, vbNormal) <> "") Then
        Kill dest
    End If

    If Extraidox = True Then
        Open dest For Binary Access Read Write As handle
        Put handle, , SourceData()
        Close handle
    End If

    Resource_Extract = Extraidox
errh:

End Function

Private Sub Resource_Generate_IH(ByRef FileName As String, ByRef InfoHead As INFOHEADER, ByRef Data() As Byte)

'On Error GoTo errh

Dim handle          As Integer
Dim SourceData()    As Byte
Dim abierto         As Byte
Dim tmpcrc          As Long
Dim ts              As String * 3
Dim freem%
Dim tmpl&
Dim filename1()     As String
Dim name_temp As String

    filename1 = Split(FileName, "\")

    freem = FreeFile()

    If (dir$(FileName, vbNormal) <> "") Then

        Open FileName For Binary Lock Read As freem
        InfoHead.lngFileSizeUncompressed = LOF(freem)
        ReDim SourceData(InfoHead.lngFileSizeUncompressed - 1) As Byte
        Get freem, , SourceData()
        Close freem

        If InfoHead.lngFileSizeUncompressed > 0 Then
            With InfoHead
                .cript = CByte(CInt(Rnd * 125)) + 1
                .originalname = LCase$(filename1(UBound(filename1)))
                name_temp = .originalname
                .originalname = Xor_String(.originalname, .cript)

                ts = LCase$(Right$(FileName, 3))
                Select Case ts
                Case "int", "dat", "ini", "ind", "xml"
                    .file_type = eTiposRecursos.rInit
                    .compress = 1
                Case "inf", "map"
                    .file_type = eTiposRecursos.rMapData
                Case "jpg", "jpeg"
                    .file_type = eTiposRecursos.rJpg
                Case "png", "tga", "dds"
                    .file_type = eTiposRecursos.rPng
                Case "bmp"
                    .file_type = eTiposRecursos.rBmp
                Case Else
                    .file_type = eTiposRecursos.rDesconocido
                End Select

'                If name_temp Like "#.#.*" Then
'                    filename1 = Split(name_temp, ".")
'                    .complemento_1 = val(filename1(LBound(filename1) + 1))
'                    If filename Like "#.#.#.*" Then
'                        .complemento_2 = val(filename1(LBound(filename1) + 2))
'                    End If
'                End If
                filename1 = Split(name_temp, ".")
                
                If UBound(filename1) > 1 Then
                    .complemento_1 = Abs(val(filename1(1))) And &H7FFF
                End If
                If UBound(filename1) > 2 Then
                    .complemento_2 = Abs(val(filename1(2))) And &H7FFF
                End If
                
                'If (.lngFileSizeUncompressed > 1500000) Then .compress = 1
                .lngFileSizeUncompressed = (.lngFileSizeUncompressed Xor CryptKeyL Xor .cript)

                If .compress Then
                    Compress_Data SourceData()
                End If

                .size_compressed = UBound(SourceData)
'                If .size_compressed > 1024 Then
'                    .crc = CRC32(SourceData(0), 1024, 0)
'                Else
'                    .crc = CRC32(SourceData(0), .size_compressed - 1, 0)
'                End If
                Data = SourceData
            End With
        Else
            Debug.Print "ERROR, FILELEN 0"; FileName
        End If
    Else
        LogError "Error en el archivo de a comprimir """ & FileName & """ - El archivo No existe."
    End If

Exit Sub

errh:
    LogError "Error en el archivo de recursos """ & FileName & """"

End Sub

Public Function Resource_Get(ByRef Path As String, ByRef FileName As String, ByRef Data() As Byte) As Boolean

    On Error GoTo errh
Dim handle As Integer
Dim SourceData() As Byte
Dim InfoHead As INFOHEADER
Dim abierto As Byte
Dim tmpcrc As Long
Dim asd As String * 16
Dim tmpl As Long

    handle = FreeFile
    'If Right$(path, 1) <> "\" Then path = path & "\"
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If '

    If LenB(dir$(Path & FileName, vbNormal)) Then
        Open Path & FileName For Binary Access Read Lock Write As handle: abierto = bTRUE
        Get handle, , asd

        If StrComp(asd, header_s, vbTextCompare) Then
            'LogError "El archivo : """ & filename & """ no es un archivo de recursos valido."
            GoTo erra
        End If

        Get handle, , InfoHead

        With InfoHead

            Extraidox = False
'            If Left$(LCase$(Xor_String(CStr(.originalname), .cript)), Len(filename)) <> LCase$(filename) Then
'                Debug.Print "Invalid Filename"
'
'#If Debuging = 0 Then
'                LogError "Error en el archivo de recursos Invalid Checksum : """ & filename & """"
'                'If abierto Then Close handle
'                If abierto Then
'                    Close handle
'                End If
'#Else
'                LogError "Error en el archivo de recursos Invalid Checksum : """ & filename & """ [" & Left$(Xor_String(CStr(.originalname), .cript), Len(filename)) & "]-[" & filename & "]"
'#End If
'                GoTo erra
'            End If

            'FINAL, leer datos, descomprimir si esta comprimido
            .lngFileSizeUncompressed = (.lngFileSizeUncompressed Xor CryptKeyL Xor .cript)
            ReDim SourceData(.size_compressed) As Byte

            Get handle, , SourceData()

'            If .size_compressed > 1024 Then
'            '    tmpcrc = CRC32(SourceData(0), 1024, 0)
'            Else
'            '    tmpcrc = CRC32(SourceData(0), .size_compressed - 1, 0)
'            End If

            'If .compress Then Decompress_Data SourceData(), .lngFileSizeUncompressed
            If .compress Then
                Decompress_Data SourceData(), .lngFileSizeUncompressed
            End If

            Data = SourceData
            last_file_ext = InfoHead
'
'                            If tmpcrc <> .crc Then
'                                Debug.Print "Invalid CRC"
'                                LogError "Error en el archivo de recursos Invalid Checksum2 : """ & filename & """ O:" & Hex(tmpcrc) & " E:" & Hex(CLng(.cript)) & " C:" & Hex(.crc)
'
'                                #If Debuging = 0 Then
'                                    If abierto = 1 Then Close handle
'                                    End
'                                #End If
'                                GoTo erra
'                            End If

            Extraidox = True
        End With
        Close handle: abierto = bFALSE
        Resource_Get = True
    Else
        LogError "Error en el archivo de recursos """ & FileName & """ - El archivo no existe."
    End If

Exit Function

errh:
    LogError "Error en el archivo de recursos """ & FileName & """ Err:" & Err.number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

Public Function Resource_Get_CRC(ByRef Path As String, ByRef FileName As String) As Long

    On Error GoTo errh
Dim handle As Integer
Dim SourceData() As Byte
Dim InfoHead As INFOHEADER
Dim abierto As Byte
Dim tmpcrc As Long
Dim asd As String * 16
Dim tmpl As Long

    handle = FreeFile
    'If Right$(path, 1) <> "\" Then path = path & "\"
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If

    If LenB(dir$(Path & FileName, vbNormal)) Then
        Open Path & FileName For Binary Access Read Lock Write As handle
        abierto = 1
        Get handle, , asd

        If StrComp(asd, header_s, vbTextCompare) Then
            LogError "El archivo : """ & FileName & """ no es un archivo de recursos valido."
            GoTo erra
        End If

        Get handle, , InfoHead

        With InfoHead

            Extraidox = False
            If Left$(UCase$(Xor_String(CStr(.originalname), .cript)), Len(FileName)) <> UCase$(FileName) Then
                Debug.Print "Invalid Filename"
                LogError "Error en el archivo de recursos Invalid Checksum : """ & FileName & """ [" & Left$(Xor_String(CStr(.originalname), .cript), Len(FileName)) & "]-[" & FileName & "]"
                GoTo erra
            End If

            'FINAL, leer datos, descomprimir si esta comprimido
            .lngFileSizeUncompressed = (.lngFileSizeUncompressed Xor CryptKeyL Xor .cript)
            If .size_compressed > 1024 Then
                ReDim SourceData(1024) As Byte
            Else
                ReDim SourceData(.size_compressed) As Byte
            End If

            Get handle, , SourceData()

            If .size_compressed > 1024 Then
                tmpcrc = CRC32(SourceData(0), 1024, 0)
            Else
                tmpcrc = CRC32(SourceData(0), .size_compressed - 1, 0)
            End If

            Resource_Get_CRC = tmpcrc
        End With
        Close handle
        Resource_Get_CRC = 0
    Else
        LogError "Error en el archivo de recursos """ & FileName & """ - El archivo no existe."
    End If

Exit Function

errh:
    LogError "Error en el archivo de recursos """ & FileName & """ Err:" & Err.number & " - Desc : " & Err.Description
erra:
    'If abierto = 1 Then Close handle
    If abierto = 1 Then
        Close handle
    End If
    Resource_Get_CRC = 0

End Function

Public Function Resource_Get_Raw(ByRef Path As String, ByRef FileName As String) As String

Dim SourceData() As Byte

    Resource_Get Path, FileName, SourceData

    If Extraidox = True Then
        Resource_Get_Raw = StrConv(SourceData, vbUnicode)
    Else
        Resource_Get_Raw = vbNullString
    End If
errh:

End Function

Public Function Resource_Read_sdf(ByRef Path As String, ByRef FileName As String) As String
'On Error GoTo errh
    Dim handle As Integer
    Dim Jo As String
    Dim abierto As Byte
    Dim tmpcrc As Byte
    Dim asd As String * 16
    Dim tmpl As Long
    Dim tmpla As Long
    Dim Bytes() As Byte
    Dim i As Integer
    Dim tr As String

    handle = FreeFile

    If Right$(Path, 1) <> "\" Then Path = Path & "\"

    If LenB(dir$(Path & FileName, vbNormal)) Then
        Open Path & FileName For Binary Access Read Lock Write As handle
            Get handle, , asd
            Get handle, , tmpcrc
            Get handle, , tmpl
            Get handle, , tmpla
            ReDim Bytes(tmpl)
            Get handle, , Bytes
        Close handle

        If StrComp(asd, header_s, vbTextCompare) Then
            #If IsServer = 0 Then
            LogError "El archivo : """ & FileName & """ no es un archivo valido."
            #End If
            GoTo errh
        Else
            tr = StrConv(Bytes, vbUnicode)
            tr = Xor_String(tr, tmpcrc)
            If CRC16(CLng(tmpcrc), tr) = tmpla / CLng(tmpcrc) Then
                Resource_Read_sdf = tr
            Else
                LogError "Se borró el archivo de recursos " & FileName
                Kill Path & FileName
                Resource_Read_sdf = vbNullString
            End If
        End If
    End If

Exit Function
errh:
LogError "Error en el archivo de recursos """ & FileName & """ Err:" & Err.number & " - Desc : " & Err.Description
End Function

Public Sub Resource_Create_sdf(ByRef datos As String, ByRef Path As String, ByRef FileName As String)
    Dim handle As Integer
    Dim tmpcrc As Byte

    Dim Jo As String
    Dim tmpl As Long
    Dim tmpla As Long
    Dim Bytes() As Byte
    
    Dim Data As String
    
    Data = datos
    
    Dim i As Long
    
    Bytes = StrConv(Data, vbFromUnicode)
    tmpcrc = CByte(CInt(Rnd * 200)) + 50
    tmpl = UBound(Bytes)
    tmpla = CRC16(CLng(tmpcrc), Data) * CLng(tmpcrc)
    
    Data = Xor_String(Data, tmpcrc)
    Bytes = StrConv(Data, vbFromUnicode)

    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    If FileExist(Path & FileName, vbNormal) Then Kill Path & FileName
    DoEvents
    
    handle = FreeFile
    
    Open Path & FileName For Binary Access Write As handle
        Put handle, , header_s
        Put handle, , tmpcrc
        Put handle, , tmpl
        Put handle, , tmpla
        Put handle, , Bytes
    Close handle
End Sub


Private Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

    If dir(file, FileType) = "" Then
        FileExist = False
      Else
        FileExist = True
    End If

End Function

Public Function Xor_String(ByRef t As String, ByVal code As Byte) As String

Dim Bytes() As Byte
Bytes = StrConv(t, vbFromUnicode)
    Call Xor_Bytes(Bytes(0), Len(t), code, CryptKey)
    Xor_String = StrConv(Bytes, vbUnicode)

End Function


Public Function Resource_Read_CFG_LNG(ByRef FileName As String, ByVal cual_cfg As Long) As Long
    Dim handle As Integer
    Dim asd As String * 16
    Dim tmpl As Long
    Dim reade As Long
    
    If LenB(dir$(FileName, vbNormal)) Then
        reade = 17 + (4 * cual_cfg)
        handle = FreeFile
        Open FileName For Binary Access Read Lock Write As handle
            Get handle, , asd
            Get handle, reade, tmpl
            If tmpl <> 0 Then
                Resource_Read_CFG_LNG = (tmpl Xor &HCD6B5CBD)
            Else
                Resource_Read_CFG_LNG = 0
            End If
        Close handle
    End If
End Function

Public Sub Resource_WRITE_CFG_LNG(ByRef FileName As String, ByVal cual_cfg As Long, ByVal value As Long)
    Dim handle As Integer
    Dim tmpl As Long
    Dim reade As Long
    
    reade = 17 + (4 * cual_cfg)
    handle = FreeFile
    
    If value = 0 Then
        tmpl = 0
    Else
        tmpl = value Xor &HCD6B5CBD
    End If
    
    If LenB(dir$(FileName, vbNormal)) Then
        Open FileName For Binary Access Read Write As handle
            Put handle, reade, tmpl
        Close handle
    Else
        Open FileName For Binary Access Read Write As handle
            Put handle, , header_s
            Put handle, reade, tmpl
        Close handle
    End If
End Sub



Private Function CRC16(ByVal key As Long, ByVal Data As String) As Integer
'**************************************************************
'Author: Salvito
'Last Modify Date: 2/07/2007
'Computes a custom CRC16 designed by Alejandro Salvo
'**************************************************************
    Dim i As Long
    Dim vstr() As Byte
    Dim SumaEspecialDeCaracteres As Long
    
    vstr = StrConv(Data, vbFromUnicode)
    
    For i = 0 To Len(Data) - 1
        SumaEspecialDeCaracteres = SumaEspecialDeCaracteres + vstr(i) * (1 + key - i)
    Next i
    
    CRC16 = CInt(Abs(SumaEspecialDeCaracteres) Mod 32000)
End Function



Public Function Get_Obj_Last_Infoheader(ByRef obj As clsEnpaquetado, ByRef ih As INFOHEADER) As Boolean
    If Not obj Is Nothing Then
        DXCopyMemory ih, ByVal obj.LastIHPtr, Len(ih)
        Get_Obj_Last_Infoheader = True
    End If
End Function

Public Function Get_Obj_Infoheader(ByRef obj As clsEnpaquetado, nro As Integer, ByRef ih As INFOHEADER) As Boolean
    If Not obj Is Nothing Then
        Dim Ptr As Long
        Ptr = obj.GetIHPtr(nro)
        If Ptr Then
            DXCopyMemory ih, ByVal Ptr, Len(ih)
            Get_Obj_Infoheader = True
        End If
    End If
End Function

Public Function Set_Obj_Infoheader(ByRef obj As clsEnpaquetado, nro As Integer, ByRef ih As INFOHEADER) As Boolean
    If Not obj Is Nothing Then
        Dim Ptr As Long
        If Ptr Then
            Set_Obj_Infoheader = obj.IH_Mod(nro, VarPtr(ih))
        End If
    End If
End Function

Public Function clsEnpaquetado_LeerIPicture(ByRef obj As clsEnpaquetado, ByVal nro As Integer) As IPicture
If Not obj Is Nothing Then
    On Error GoTo errh
    
    Dim SourceData()    As Byte
    Dim LowerBound      As Long
    Dim ByteCount       As Long
    Dim hMem            As Long
    Dim lpMem           As Long
    Dim istm            As stdole.IUnknown
    Dim IID_IPicture(15) ' no sabe no contesta
    
        If obj.Leer(nro, SourceData) Then
            LowerBound = LBound(SourceData)
            ByteCount = (UBound(SourceData) - LowerBound) + 1
            hMem = GlobalAlloc(&H2, ByteCount)
            If hMem <> 0 Then
                lpMem = GlobalLock(hMem)
                If lpMem <> 0 Then
                    MoveMemory ByVal lpMem, SourceData(LowerBound), ByteCount
                    Call GlobalUnlock(hMem)
                    If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                        If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                          Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), clsEnpaquetado_LeerIPicture)
                        End If
                    End If
                End If
                
            End If
        End If
        Exit Function
    
errh:
        If Err.number = 9 Then
            LogError "LeerIPicture->BA empty"
        Else
            LogError "LeerIPicture->(" & Err.number & ") " & Err.Description
        End If
End If
End Function

