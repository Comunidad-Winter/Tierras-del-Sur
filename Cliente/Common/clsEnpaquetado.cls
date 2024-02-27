VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnpaquetado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: ErrReportAndTrace
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

Option Explicit


Private Const header_b As String * 16 = "MZEngineBinarir§"

Private Declare Function compress Lib "zlib.dll" _
        (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" _
        (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef dest As Any, ByRef Source As Any, ByVal ByteCount As Long)

Private Declare Sub MDFile Lib "aamd532.dll" _
        (ByVal f As String, ByVal r As String)

Private Declare Sub MDStringFix Lib "aamd532.dll" _
        (ByVal f As String, ByVal T As Long, ByVal r As String)
        
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



Private Const CryptKeyL     As Long = 984362498

Public Path_res            As String

Private last_file_ext       As INFOHEADER

Private Cabezal()  As INFOHEADER
Private Cabezal_Ultimo_Byte    As Long
Private Cabezal_Cantidad   As Integer

Private Const Min_Offset    As Integer = 5000 ' El "cacho" de slots libres del array para agregar archivos

Private Iterador_AnteriorIndex As Integer

Private Enum File_Privs
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

Public Property Get LastIHPtr() As Long
    LastIHPtr = VarPtr(last_file_ext)
End Property

Public Function GetIHPtr(ByVal nro As Integer) As Long
    'Marce On error resume next
    If nro <= Cabezal_Cantidad Then
        GetIHPtr = VarPtr(Cabezal(nro))
    End If
End Function


Public Function Cargar(ByVal Path As String) As Boolean 'Carga un archivo enpaquetado en la memoriaz(solo headers)

    Dim cantidad    As Integer
    Dim handle      As Integer
    Dim t_str       As String * 16
    Dim abierto     As Byte

    Path_res = Path
    handle = FreeFile()

    Open Path_res For Binary As handle: abierto = bTRUE
        Get handle, 1, t_str
        
        If StrComp(t_str, header_b, vbTextCompare) Then GoTo erra
    
        Get handle, , Cabezal_Cantidad
        Get handle, , Cabezal_Ultimo_Byte
    
        ReDim Cabezal(Cabezal_Cantidad)
    
        Get handle, , Cabezal
        
    Close handle: abierto = bFALSE

    Cargar = True
Exit Function

erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If
    LogError "El archivo : """ & Path_res & """ no es un archivo de recursos valido. Error: " & Err.Number & " " & Err.Description
    'End

End Function

Public Function CrearVacio(ByVal FileName As String, ByVal Tamaño As Long) As Boolean
'Crea un nuevo enpaquetADO EN LA memoria
'False si el archivo ya existe,
    Dim handle      As Integer

    Path_res = FileName
    
    If (Dir$(Path_res, vbNormal) <> "") Then
        CrearVacio = False
    Else
        '-Reiniciamos...------------------------
            Cabezal_Cantidad = Tamaño + Min_Offset
            If Cabezal_Cantidad > Max_Int_Val Then Cabezal_Cantidad = Max_Int_Val
            ReDim Cabezal(Cabezal_Cantidad)
            Cabezal_Ultimo_Byte = 0
        '---------------------------------------
        
        CrearVacio = True
        handle = FreeFile
        Open Path_res For Binary Access Read Write As handle
            Seek handle, 1
            
            Put handle, , header_b
            Put handle, , Cabezal_Cantidad
            Put handle, , Cabezal_Ultimo_Byte
            Put handle, , Cabezal
            
            Cabezal_Ultimo_Byte = seek(handle) + 100 ' Marcamos el header + 100 bytes para que empiece a escribir los archivos después.
            
            Seek handle, 1
            Put handle, , header_b
            Put handle, , Cabezal_Cantidad
            Put handle, , Cabezal_Ultimo_Byte
            
        Close handle
    End If
End Function

Public Function CrearDesdeCarpeta(ByVal FileName As String, ByRef archivos() As String) As Boolean

'On Error GoTo errh

Dim handle          As Integer

Dim abierto         As Byte

Dim Nueva_Cantidad  As Integer
Dim Ultimo_Byte     As Long
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

Dim new_file        As String
Dim asd             As String * 16
Dim necesita_hacer  As Byte
Dim tmpbn As Double

    

    File_List = archivos
    cantidad_array = UBound(File_List) + 1                'Cuento los archivos
    
    If cantidad_array = 0 Then                          'Si no hay archivos, cancelo la operación
        CrearDesdeCarpeta = False
        Exit Function
    End If
    
    ReDim int_list(cantidad_array - 1)

    For i = 0 To cantidad_array - 1
        tmpbn = val(mid$(File_List(i), InStrRev(File_List(i), "\") + 1))  'Obtengo el número de archivo
        If tmpbn <= Max_Int_Val Then
            TmpInt = tmpbn
    
            If max_cantidad < TmpInt Then max_cantidad = TmpInt
            int_list(i) = TmpInt                        'Defino que archivo del array corresponde a este número de cabezal
        End If
    Next i

    tmplng = max_cantidad + Min_Offset                  'Agrando el tamaño del cabezal
    If tmplng > Max_Int_Val Then tmplng = Max_Int_Val
    
    Cabezal_Cantidad = tmplng

    ReDim Cabezal(Cabezal_Cantidad)

    Path_res = FileName

    If (Dir$(Path_res, vbNormal) <> "") Then Kill Path_res
    
    handleB = FreeFile()
    Open Path_res For Binary Access Read Write As handleB
    Seek handleB, 1

    Ultimo_Byte = 0
    Put handleB, , header_b
    Put handleB, , Cabezal_Cantidad
    Put handleB, , Cabezal_Ultimo_Byte
    Put handleB, , Cabezal

    For i = 0 To cantidad_array - 1
        If int_list(i) > 0 Then
            new_file = File_List(i)
            handle = FreeFile()
            Open new_file For Binary Access Read Lock Write As handle: abierto = bTRUE
            Get handle, , asd
            
            If StrComp(asd, header_s, vbTextCompare) Then
                necesita_hacer = bTRUE
            Else
                Get handle, , InfoHead
                ReDim SourceData(InfoHead.size_compressed) As Byte
                Get handle, , SourceData()
            End If
            
            Close handle: abierto = bFALSE

            If necesita_hacer Then
                Resource_Generate_IH new_file, InfoHead, SourceData
            End If

            InfoHead.EmpiezaByte = seek(handleB)

            Cabezal(int_list(i)) = InfoHead

            Put handleB, , SourceData
            Cabezal_Ultimo_Byte = seek(handleB) + 1
            
            Debug.Print "Push ("; int_list(i); ") - "; File_List(i); " - CRC:"; Hex(InfoHead.CRC); InfoHead.size_compressed; "- ptr:"; Hex$(Ultimo_Byte)
        End If
    Next i

    Seek handleB, 1

    Put handleB, , header_b
    Put handleB, , Cabezal_Cantidad
    Put handleB, , Cabezal_Ultimo_Byte
    Put handleB, , Cabezal
    
    Close handleB

    CrearDesdeCarpeta = True

Exit Function

errh:
    LogError "Error en el archivo de recursos """ & FileName & """ Err:" & Err.Number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

' Obtiene un numero de recurso poniendo la informacion de los bytes del recurso en Data()
'TODO rFile_type Parametro muerto.
Public Function Leer(ByRef nFile As Integer, ByRef Data() As Byte, Optional ByVal rFile_type As e_resource_file = -1) As Boolean
On Error GoTo errh

Dim handle As Integer
Dim SourceData() As Byte
Dim InfoHead As INFOHEADER
Dim abierto As Byte

    If nFile = 0 Then
        Leer = False
        Exit Function
    End If

    'Obtenemos la info del archivo
    If Cabezal_Cantidad >= nFile Then
        InfoHead = Cabezal(nFile)
    Else
        Leer = False
        Exit Function
    End If

    handle = FreeFile()
    
    If InfoHead.EmpiezaByte Then
        'Abrimos el archivo
        Open Path_res For Binary Access Read Lock Write As handle: abierto = bTRUE
            Seek handle, InfoHead.EmpiezaByte   ' movemos el puntero de handle a EmpiezaByte
            
            ' Generamos un array en donde vamos a poner la info
            ReDim SourceData(InfoHead.size_compressed) As Byte
            Get handle, , SourceData()          ' Leo el contenido.
        
            If InfoHead.compress = 1 Then       ' Si esta comprimido, lo descomprimimos.
                Decompress_Data SourceData(), InfoHead.lngFileSizeUncompressed Xor CryptKeyL Xor InfoHead.cript
            End If
            
            Data = SourceData                   'REturneo la data
        Close handle: abierto = bFALSE
    
        last_file_ext = InfoHead
        Leer = True
    End If
    
Exit Function

errh:
    LogError "Error en el archivo de recursos """ & Path_res & """ Err:" & Err.Number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

'Lee en base a un infoheader
Friend Function LeerIH(ByRef Data() As Byte, ByRef InfoHead As INFOHEADER) As Boolean
On Error GoTo errh

Dim handle As Integer
Dim SourceData() As Byte
Dim abierto As Byte

    handle = FreeFile()
    
    If InfoHead.EmpiezaByte Then
        Open Path_res For Binary Access Read Lock Write As handle: abierto = bTRUE
            Seek handle, InfoHead.EmpiezaByte   ' movemos el puntero de handle a EmpiezaByte
            
            ReDim SourceData(InfoHead.size_compressed) As Byte
            Get handle, , SourceData()          ' Leo el contenido.
        
            If InfoHead.compress = 1 Then       ' Si esta comprimido, lo descomprimimos.
                Decompress_Data SourceData(), InfoHead.lngFileSizeUncompressed Xor CryptKeyL Xor InfoHead.cript
            End If
            
            Data = SourceData                   'REturneo la data
        Close handle: abierto = bFALSE
    
        last_file_ext = InfoHead
        LeerIH = True
    End If
    
Exit Function

errh:
    LogError "Error en el archivo de recursos """ & Path_res & """ Err:" & Err.Number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

Public Function Extraer(ByRef nro As Integer, ByRef dest As String) As Boolean

'On Error GoTo errh

Dim SourceData() As Byte
Dim handle%

    
Debug.Print dest
    If (Dir$(dest, vbNormal) <> "") Then
        Kill dest
    End If

    If Leer(nro, SourceData) Then
        handle = FreeFile()
        Open dest For Binary Access Read Write As handle
        Put handle, , SourceData()
        Close handle
        Extraer = True
    End If

errh:

End Function

'Igual que extraer pero recibe como parametro el INFOHEADER
Friend Function ExtraerIH(ByRef InfoHead As INFOHEADER, ByRef dest As String) As Boolean

'On Error GoTo errh

    Dim SourceData() As Byte
    Dim handle%

    If (Dir$(dest, vbNormal) <> "") Then
        Kill dest
    End If

    If LeerIH(SourceData, InfoHead) Then
        handle = FreeFile()
        Open dest For Binary Access Read Write As handle
        Put handle, , SourceData()
        Close handle
        ExtraerIH = True
    End If

errh:
'If handle Then Close handle
End Function

Public Function Cabezal_GetFileSize(ByVal nFile As Integer) As Long
    If nFile = 0 Then
        Exit Function
    End If

    If Cabezal_Cantidad >= nFile Then
        Cabezal_GetFileSize = Cabezal(nFile).size_compressed
    Else
        Exit Function
    End If
End Function

Public Function Cabezal_GetFilePtr(ByVal nFile As Integer) As Long
    If nFile = 0 Then
        Exit Function
    End If

    If Cabezal_Cantidad >= nFile Then
        Cabezal_GetFilePtr = Cabezal(nFile).EmpiezaByte
    Else
        Exit Function
    End If
End Function

Public Function Cabezal_GetFilenameName(ByVal nFile As Integer) As String
    Cabezal_GetFilenameName = vbNullString
    If nFile = 0 Then
        Exit Function
    End If

    Dim Complementos As String
    Dim nombreSolo As String
    
    If Cabezal_Cantidad >= nFile Then
        Complementos = "." & Cabezal(nFile).complemento_1 & "." & Cabezal(nFile).complemento_2 & "." & Cabezal(nFile).complemento_3 & "." & Cabezal(nFile).complemento_4 & "."
         
        nombreSolo = Replace(Trim$(LCase$(Xor_String(Cabezal(nFile).originalname, Cabezal(nFile).cript))), nFile & Complementos, "")
        
        Cabezal_GetFilenameName = nFile & Complementos & nombreSolo
    Else
        Exit Function
    End If
End Function

'Retorna solamente el nombre del archivo, sin complementos
Public Function Cabezal_GetFileNameSinComplementos(ByVal nFile As Integer) As String
    Cabezal_GetFileNameSinComplementos = vbNullString
    If nFile = 0 Then
        Exit Function
    End If

    Dim Complementos As String

    If Cabezal_Cantidad >= nFile Then
        Cabezal_GetFileNameSinComplementos = Replace(Trim$(LCase$(Xor_String(Cabezal(nFile).originalname, Cabezal(nFile).cript))), Complementos, "")
    Else
        Exit Function
    End If
End Function

Public Function Cabezal_GetCRC(ByRef nFile As Integer) As Long

    Cabezal_GetCRC = &H0

    If UBound(Cabezal) >= nFile Then _
        Cabezal_GetCRC = Cabezal(nFile).CRC


End Function

Public Function LeerRAW(ByRef nFile As Integer) As String
Dim SourceData() As Byte

    If Leer(nFile, SourceData) Then
        LeerRAW = StrConv(SourceData, vbUnicode)
    Else
        LeerRAW = vbNullString
    End If
End Function

'Carga en InfoHead la informacion de un Recurso. Devuelve FALSE sino existe.
Friend Function IH_Get(ByRef nFile As Integer, ByRef InfoHead As INFOHEADER) As Boolean
    If Cabezal_Cantidad < nFile Then
        Exit Function
    Else
        InfoHead = Cabezal(nFile)
        IH_Get = True
    End If
End Function


Private Sub IH_Get_Ptr(ByVal Ptr As Long, ByRef IH As INFOHEADER)
Dim handle As Integer
handle = FreeFile

If Ptr = 0 Then Exit Sub

Open Path_res For Binary Access Read As handle
    Seek handle, Ptr
    Get handle, , IH
Close handle


End Sub

Friend Function IH_Mod(ByRef nFile As Integer, ByRef InfoHead As INFOHEADER) As Boolean

'On Error GoTo errh

Dim handle          As Integer

Dim abierto         As Byte

Dim fhp             As Long
Dim TIH As INFOHEADER

    
    ' Obtenemos el info Header, sino existe, salgo
    If IH_Get(nFile, TIH) = False Then Exit Function
    
    ' Si no existe, salgo
    If TIH.lngFileSizeUncompressed = 0 Then Exit Function
    
    'Copio los datos
    InfoHead.compress = TIH.compress
    InfoHead.cript = TIH.cript
    InfoHead.EmpiezaByte = TIH.EmpiezaByte
    InfoHead.size_compressed = TIH.size_compressed
    InfoHead.lngFileSizeUncompressed = TIH.lngFileSizeUncompressed
    InfoHead.originalname = TIH.originalname
    InfoHead.PreviousHeader = TIH.PreviousHeader
    InfoHead.CantidadVersionesEnPak = TIH.CantidadVersionesEnPak
    InfoHead.file_type = TIH.file_type

    'TODO Redundante?
    Cabezal(nFile) = InfoHead
    
    'Obtenemos un archivo
    handle = FreeFile()
    Debug.Print "[MODIFICANDO:" & nFile & "]"
    
    'TODO WTF?
    fhp = 23 + CLng(nFile) * Len(InfoHead) ' muejeje
    
    Open Path_res For Binary Access Read Write As handle: abierto = bTRUE
        Put handle, fhp, InfoHead
    Close handle: abierto = bFALSE
    
    Debug.Print " Offset head:"; fhp
    Debug.Print " MODIFICADO OK."
    Debug.Print ""
    IH_Mod = True
Exit Function

errh:
    LogError "Error en mod el archivo de recursos """ & Path_res & """ Err:" & Err.Number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

Friend Function Iterar(ba() As Byte, ByRef nFile As Integer, ByRef InfoHeadPtr As INFOHEADER) As Boolean
    Dim i As Integer
    Dim handle As Integer
    
    If Iterador_AnteriorIndex >= Cabezal_Cantidad Then
        Iterar = False
        Exit Function
    End If
    
    For i = Iterador_AnteriorIndex + 1 To Cabezal_Cantidad
        If Cabezal(i).size_compressed Then
            InfoHeadPtr = Cabezal(i)
            Debug.Print "Parcheando " & Xor_String(Cabezal(i).originalname, Cabezal(i).cript)
            
            Iterador_AnteriorIndex = i
            
            nFile = i
            
            handle = FreeFile
        
            Open Path_res For Binary Access Read As handle
                Seek handle, Cabezal(i).EmpiezaByte           ' movemos el puntero de handle a EmpiezaByte
                ReDim ba(Cabezal(i).size_compressed) As Byte
                Get handle, , ba()                          ' Leo el contenido.
            Close handle
            
            Iterar = True
            
            Exit For
        End If
    Next i
End Function

Public Function ExtraerVersiones(ByRef nFile As Integer, ByVal Path As String) As Boolean
    Dim IH As INFOHEADER
    Dim tmpPath As String
    
    If IH_Get(nFile, IH) Then
        ExtraerVersiones = True
        ExtraerVersiones = ExtraerVersiones And ExtraerIH(IH, Path & Trim(Xor_String(IH.originalname, IH.cript)))
        Do While IH.PreviousHeader <> 0
            IH_Get_Ptr CLng(IH.PreviousHeader), IH
            
            tmpPath = Trim(Path & Xor_String(IH.originalname, IH.cript))
            If FileExist(tmpPath, vbNormal) Then
                ExtraerVersiones = ExtraerVersiones And ExtraerIH(IH, tmpPath & ".r" & IH.CantidadVersionesEnPak)
            Else
                ExtraerVersiones = ExtraerVersiones And ExtraerIH(IH, tmpPath)
            End If
        Loop
    End If
End Function


Public Function ExtraerVersion(ByRef nFile As Integer, ByRef Version As Integer, ByVal Path As String) As Boolean
    Dim IH As INFOHEADER
    Dim tmpPath As String
    
    If IH_Get(nFile, IH) Then
    
        If IH.CantidadVersionesEnPak = Version Then
            tmpPath = Trim(Path & Xor_String(IH.originalname, IH.cript))
            If FileExist(tmpPath, vbNormal) Then
                If MsgBox("Desea sobreescribir el archivo ¿""" & tmpPath & """?", vbYesNo) = vbNo Then
                    Exit Function
                End If
            End If
            ExtraerIH IH, tmpPath
            ExtraerVersion = True
            Exit Function
        End If
        
        Do While IH.PreviousHeader <> 0
            IH_Get_Ptr CLng(IH.PreviousHeader), IH
            
            If IH.CantidadVersionesEnPak = Version Then
                tmpPath = Trim(Path & Xor_String(IH.originalname, IH.cript))
                If FileExist(tmpPath, vbNormal) Then
                    If MsgBox("Desea sobreescribir el archivo ¿""" & tmpPath & """?", vbYesNo) = vbNo Then
                        Exit Function
                    End If
                End If
                ExtraerIH IH, tmpPath
                ExtraerVersion = True
                Exit Function
            End If
        Loop
        
    End If
End Function


Friend Function ParchearByteArray(ba() As Byte, ByVal nFile As Integer, ByRef InfoHead As INFOHEADER) As Boolean
'On Error GoTo errh

Dim handle          As Integer

Dim abierto         As Byte

Dim fhp As Long

    'Si no hay lugar, salgo
    If nFile > Cabezal_Cantidad Then Exit Function

    'Si no tiene nada para parchear, salgo
    If InfoHead.size_compressed = 0 Then
        MsgBox "Error al parchear. nFile=" & nFile & " - Está vacío."
        Exit Function
    End If
    
    'Tomo un archivo
    handle = FreeFile()
    
    ' Abro el empaquetado
    Open Path_res For Binary Access Read Write As handle: abierto = bTRUE
    
        'Me posiciono al final
        Seek handle, LOF(handle) + 1
        
        If Cabezal(nFile).EmpiezaByte > 0 Then
            'Muevo la cabecera actual al final del archivo para poder hacer rollback
            InfoHead.PreviousHeader = seek(handle)
            InfoHead.CantidadVersionesEnPak = Cabezal(nFile).CantidadVersionesEnPak + 1
            Put handle, , Cabezal(nFile) 'Escribimos la informacion vieja del cabezal
        Else
            'Es la primera vez que se agrega el archivo
            InfoHead.PreviousHeader = 0
            InfoHead.CantidadVersionesEnPak = 0
        End If
        
        ' Me posiciono al final nuevamente
        Seek handle, LOF(handle) + 1
        
        'El nuevo elemento va al final del archivo
        InfoHead.EmpiezaByte = seek(handle) 'obtengo el inicio del archivo
        
        ' Se escribe el elemento
        Put handle, , ba                    'Escribo el archivo
        
        'Guardamos el ultimo byte del archivo
        Cabezal_Ultimo_Byte = seek(handle)
        
        ' Vamos al principio y escribimos
        Seek handle, 1
        Put handle, , header_b
        Put handle, , Cabezal_Cantidad
        Put handle, , Cabezal_Ultimo_Byte
        
        ' Escribe el cabezal modificado
        ' Datos + La cantidad de cada cabezal * el numero de archivo.
        ' Se esta suponiendo que el cabezal siempre va aestar en la misma posicion.
        fhp = seek(handle) + CLng(nFile) * Len(InfoHead) ' muejeje, escribo el cabezal
        Put handle, fhp, InfoHead
        
        ' Info de debug
        #If ModoDebug = 1 Then
            LogDebug "[PARCHEANDO:" & Path_res & "]"
            LogDebug " File:" & nFile & " Nombre: " & Xor_String(InfoHead.originalname, InfoHead.cript)
            LogDebug " Offset head:" & fhp
            LogDebug " Offset file:" & InfoHead.EmpiezaByte
            LogDebug " Tamaño:" & InfoHead.size_compressed
            LogDebug " Comprimido: " & CBool(InfoHead.compress)
            LogDebug " Ultimo byte:" & Cabezal_Ultimo_Byte
            LogDebug " Cantidad rollback:" & InfoHead.CantidadVersionesEnPak
            LogDebug " Complemento 1 " & InfoHead.complemento_1 & ". 2: " & InfoHead.complemento_2 & ". 3:" & InfoHead.complemento_3 & ". 4:" & InfoHead.complemento_4
            LogDebug " PARCHEADO OK."
        #End If
        
        ParchearByteArray = True
    Close handle: abierto = bFALSE

    Cabezal(nFile) = InfoHead

Exit Function

errh:
    LogError "Error en ParchearByteArray """ & Path_res & """ Err:" & Err.Number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If




End Function

Public Function Parchear(ByRef nFile As Integer, ByRef new_file As String, Optional ByVal CRC As Long = -1) As Boolean

'On Error GoTo errh

Dim handle          As Integer

Dim InfoHead        As INFOHEADER

Dim abierto         As Byte
Dim file_len        As Long


Dim tmp_s           As String * 16






Dim SourceData()    As Byte




Dim resultado             As Boolean
Dim ArchivoPack     As clsEnpaquetado

    file_len = FileLen(new_file)

    If new_file = Path_res Then
        Err.Raise 2371465, , "Queres parchar con el mismo parche?"
        GoTo errh
    End If

   

    'If Not LenB(Dir$(path & filename, vbNormal)) Then GoTo errh
    
    If file_len = 0 Then
        'Err.Raise 2371465, , "Archivo " & Parche & " invalido"
        MsgBox "El archivo no tiene contenido"
        GoTo errh                           'Si no existe el archivo q vamos a parchear, chau.
    End If

    handle = FreeFile

    Open new_file For Binary Access Read Lock Write As handle: abierto = bTRUE
    Get handle, , tmp_s
    
    Select Case tmp_s
    Case header_s
        Get handle, , InfoHead
        ReDim SourceData(InfoHead.size_compressed) As Byte
        Get handle, , SourceData()
        Close handle: abierto = bFALSE
        Parchear = ParchearByteArray(SourceData(), nFile, InfoHead)
    Case header_b
        Set ArchivoPack = New clsEnpaquetado
        Close handle: abierto = bFALSE              ' Cierro el archivo para evitar runtimes
        DoEvents
        
        If ArchivoPack.Cargar(new_file) Then        'Cargo el archivo denuevo
            LogDebug "Parcheador->" & "Cargando archivo enpaquetado: " & new_file & " para " & Path_res
            Parchear = True
            While ArchivoPack.Iterar(SourceData(), nFile, InfoHead)
                'CopyMemory InfoHead, ByVal InfoHeadPtr, Len(InfoHead)
                resultado = ParchearByteArray(SourceData(), nFile, InfoHead)
                LogDebug "____ Parcheador--->" & "Parcheando: " & nFile & " Resultado: " & resultado
                Parchear = Parchear And resultado
            Wend
            LogDebug "____ [Parcheado = " & Parchear & "]"
            
            Exit Function
        Else
            LogError "Error al cargar el parche enpaquetado"
        End If
    Case Else
        Close handle: abierto = bFALSE
        Resource_Generate_IH new_file, InfoHead, SourceData
        Parchear = ParchearByteArray(SourceData(), nFile, InfoHead)
    End Select

    If abierto Then Close handle: abierto = bFALSE
    
    

Exit Function

errh:
    LogError "Error queriendo parchar en el archivo de recursos """ & Path_res & """ el archivo numero " & nFile & " proveniente del archivo " & new_file & " Err:" & Err.Number & " - Desc : " & Err.Description
erra:
    'If abierto Then Close handle
    If abierto Then
        Close handle
    End If

End Function

Private Function FileExist(file As String, FileType As VbFileAttribute) As Boolean
    If Dir(file, FileType) = "" Then
        FileExist = False
      Else
        FileExist = True
    End If
End Function

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


'******************************************************************************
' PERMISOS de los elementos del empaquetado
'******************************************************************************

Public Function Add_To_Listbox(ByRef list As ListBox) As Boolean
''Marce On error resume next
Dim InfoHead    As INFOHEADER
Dim abierto     As Byte
Dim i           As Integer

    list.Clear

    For i = 0 To Cabezal_Cantidad
        InfoHead = Cabezal(i)
        'If InfoHead.size_compressed Then _
            List.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript)))
        If InfoHead.size_compressed Then
            list.AddItem i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript))) & " - CRC:" & Hex$(InfoHead.CRC) & " - " & Round(InfoHead.size_compressed / 1024, 1) & "KB"
        End If
    Next i

End Function

Public Function Add_To_Listbox_Permisos(ByRef list As ListBox, ByVal MisPermisos As Long, ByVal MiUser As Integer) As Boolean
''Marce On error resume next
Dim InfoHead    As INFOHEADER
Dim abierto     As Byte
Dim i           As Integer
Dim Strin       As String
Dim ST          As String


    list.Clear

    If Cabezal_Cantidad = 0 Then
        list.AddItem "No se cargó el enpaquetado o está vacío."
        Exit Function
    End If

    For i = 0 To Cabezal_Cantidad
        InfoHead = Cabezal(i)

        If InfoHead.size_compressed Then
            Strin = i & " - " & Trim$(LCase$(Xor_String(InfoHead.originalname, InfoHead.cript)))
            ST = ""

            'If (InfoHead.owner = MiUser And MiUser > 0) Or MisPermisos = -1 Then
            '    List.AddItem Strin & " - [Editar Extraer Rollback]"
            'Else
                If (InfoHead.privs And File_Privs.PuedeEditarCualquiera) And (MisPermisos And File_Privs.PuedeEditarCualquiera) Then ST = ST & "Editar"
                If (InfoHead.privs And File_Privs.PuedeExtraerArchivos) And (MisPermisos And File_Privs.PuedeExtraerArchivos) Then ST = ST & "Extraer"
                If (InfoHead.privs And File_Privs.PuedeHacerRollBack) And (MisPermisos And File_Privs.PuedeHacerRollBack) Then ST = ST & "Rollback {" & InfoHead.CantidadVersionesEnPak & "}"
                
                If ST = "" Then
                    list.AddItem "*" & Strin
                Else
                    list.AddItem Strin & " - [" & ST & "]"
                End If
            'End If
        End If
    Next i

End Function

Public Function Add_To_Listbox_Versiones(ByRef list As ListBox, ByVal nFile As Integer) As Boolean
    list.Clear

    Dim IH As INFOHEADER
    If IH_Get(nFile, IH) Then
    

        list.AddItem IH.CantidadVersionesEnPak & vbTab & "[" & Trim$(LCase$(Xor_String(IH.originalname, IH.cript))) & "] <- EN USO"
        
        Do While IH.PreviousHeader <> 0
            IH_Get_Ptr CLng(IH.PreviousHeader), IH
            list.AddItem IH.CantidadVersionesEnPak & vbTab & Trim$(LCase$(Xor_String(IH.originalname, IH.cript)))
        Loop
        
    End If


End Function


Public Function Puedo_Extraer(ByVal nFile As Integer, Optional ByVal MisPermisos As Long = -1, Optional ByVal MiUser As Integer = 0) As Boolean
''Marce On error resume next
Dim InfoHead    As INFOHEADER
    If nFile <= Cabezal_Cantidad Then
        InfoHead = Cabezal(nFile)

        If InfoHead.size_compressed Then
            If (InfoHead.owner = MiUser And MiUser > 0) Or MisPermisos = -1 Then
                Puedo_Extraer = True
            Else
                If (InfoHead.privs And File_Privs.PuedeExtraerArchivos) And (MisPermisos And File_Privs.PuedeExtraerArchivos) Then Puedo_Extraer = True
            End If
        End If
    End If
End Function

Public Function Puedo_RollBack(ByVal nFile As Integer, Optional ByVal MisPermisos As Long = -1, Optional ByVal MiUser As Integer = 0) As Boolean
''Marce On error resume next
Dim InfoHead    As INFOHEADER
    If nFile <= Cabezal_Cantidad Then
        InfoHead = Cabezal(nFile)

        If InfoHead.size_compressed > 0 And InfoHead.CantidadVersionesEnPak > 0 Then
            If (InfoHead.owner = MiUser And MiUser > 0) Or MisPermisos = -1 Then
                Puedo_RollBack = True
            Else
                If (InfoHead.privs And File_Privs.PuedeHacerRollBack) And (MisPermisos And File_Privs.PuedeHacerRollBack) Then Puedo_RollBack = True
            End If
        End If
    End If
End Function

Public Function Puedo_Editar(ByVal nFile As Integer, Optional ByVal MisPermisos As Long = -1, Optional ByVal MiUser As Integer = 0) As Boolean
''Marce On error resume next
Dim InfoHead    As INFOHEADER
    If nFile <= Cabezal_Cantidad Then
        InfoHead = Cabezal(nFile)

        If InfoHead.size_compressed Then
            If (InfoHead.owner = MiUser And MiUser > 0) Or MisPermisos = -1 Then
                Puedo_Editar = True
            Else
                If (InfoHead.privs And File_Privs.PuedeEditarCualquiera) And (MisPermisos And File_Privs.PuedeEditarCualquiera) Then Puedo_Editar = True
            End If
        Else
            Puedo_Editar = True
        End If
    End If
End Function

Public Function getCantidadElementos() As Long
    getCantidadElementos = Cabezal_Cantidad
End Function

Public Function obtenerIndiceArchivo(nombreArchivo As String) As Integer
    Dim loopRecurso As Integer
    Dim Nombre As String
    
    ' Para la comparación no tengo en cuenta a los complementos
    For loopRecurso = LBound(Cabezal) To UBound(Cabezal)
    
        Nombre = Cabezal_GetFileNameSinComplementos(loopRecurso)
        
        If LCase$(nombreArchivo) = LCase$(Nombre) Then
            obtenerIndiceArchivo = loopRecurso
            Exit Function
        End If
    Next
    
    obtenerIndiceArchivo = 0
End Function

