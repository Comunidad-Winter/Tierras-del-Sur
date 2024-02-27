Attribute VB_Name = "CDMFileDialogs"
'Author: Luis Cantero
'© 2002-2005 L.C. Enterprises
'http://LCen.com

Option Explicit

Private Type OPENFILENAME

    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000 ' new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
'Folder
Private Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Const WM_USER = &H400
Public Const LPTR = (&H0 Or &H40)
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

'Open/Save
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

    If uMsg = 1 Then
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If

End Function

Public Function BrowseForFolder(strTitle As String, lngHwnd As Long, Optional strInitialDirectory As String) As String

  Dim Browse_for_folder As BrowseInfo
  Dim lngItemID As Long
  Dim lngInitDirPointer As Long
  Dim strTempPath As String * 256

    If strInitialDirectory = "" Then strInitialDirectory = app.Path

    With Browse_for_folder
        .hOwner = lngHwnd 'Window Handle
        .lpszTitle = strTitle 'Dialog Title
        .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) 'Dialog callback function that preselectes the folder specified
        lngInitDirPointer = LocalAlloc(LPTR, Len(strInitialDirectory) + 1) 'Allocate a string
        Call CopyMemory(ByVal lngInitDirPointer, ByVal strInitialDirectory, Len(strInitialDirectory) + 1) 'Copy the path to the string
        .lParam = lngInitDirPointer  'The folder to preselect
    End With

    lngItemID = SHBrowseForFolder(Browse_for_folder) 'Execute the BrowseForFolder API

    If lngItemID Then
        If SHGetPathFromIDList(lngItemID, strTempPath) Then ' Get the path for the selected folder in the dialog
            BrowseForFolder = Left$(strTempPath, InStr(strTempPath, vbNullChar) - 1) ' Take only the path without the nulls
        End If

        Call CoTaskMemFree(lngItemID) 'Free the lngItemID
    End If

    Call LocalFree(lngInitDirPointer) 'Free the string from the memory

End Function

Private Function FunctionPointer(FunctionAddress As Long) As Long

    FunctionPointer = FunctionAddress

End Function

'"JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|CompuServe GIF (*.gif)|*.gif"
Public Function OpenDialog(strFilter As String, strTitle As String, strDefaultExtension As String, strInitialDirectory As String, lngHwnd As Long) As String

    On Error GoTo Problems
  Dim OpenFile As OPENFILENAME
  Dim strTemp As String
  Dim intNull As Integer

    If Right$(strFilter, 1) <> Chr$(0) Then strFilter = strFilter & Chr$(0)
    strFilter = Replace(strFilter, "|", Chr$(0))

    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hWndOwner = lngHwnd
    OpenFile.lpstrInitialDir = strInitialDirectory
    OpenFile.hInstance = app.hInstance
    OpenFile.lpstrFilter = strFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String$(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrTitle = strTitle
    OpenFile.lpstrDefExt = strDefaultExtension
    OpenFile.flags = OFN_HIDEREADONLY

    If GetOpenFileName(OpenFile) = 0 Then
        OpenDialog = ""
      Else
        strTemp = OpenFile.lpstrFile
        intNull = InStr(1, strTemp, Chr$(0))
        OpenDialog = mid$(strTemp, 1, intNull - 1)
    End If

Exit Function

Problems:
    MsgBox Err.Description, 16, "Error " & Err.number

End Function

Public Function SaveDialog(strFilter As String, strTitle As String, strInitialDirectory As String, lngHwnd As Long, Optional strFileName As String) As String

  Dim OpenFile As OPENFILENAME
  Dim strExtension As String

    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hWndOwner = lngHwnd
    OpenFile.hInstance = app.hInstance
    If Right$(strFilter, 1) <> "|" Then strFilter = strFilter + "|"

    strFilter = Replace(strFilter, "|", Chr$(0))
    If strFileName = "" Then strFileName = Space$(254) Else strFileName = strFileName & Space$(254 - Len(strFileName))

    OpenFile.lpstrFilter = strFilter
    OpenFile.lpstrFile = strFileName
    OpenFile.nMaxFile = 255
    OpenFile.lpstrFileTitle = Space$(254)
    OpenFile.nMaxFileTitle = 255
    OpenFile.lpstrInitialDir = strInitialDirectory
    OpenFile.lpstrTitle = strTitle
    OpenFile.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT

    If GetSaveFileName(OpenFile) Then
        SaveDialog = Trim$(OpenFile.lpstrFile)
        strExtension = mid$(Right$(strFilter, 5), 1, 4)
        strFileName = Left$(SaveDialog, Len(SaveDialog) - 1)
        If Right$(strFileName, 4) = strExtension Then strExtension = ""
        SaveDialog = strFileName & strExtension
        If strFilter = "*.*" & Chr$(0) Then SaveDialog = strFileName
      Else
        SaveDialog = ""
    End If

End Function

':) Ulli's VB Code Formatter V2.13.6 (18.08.2005 22:27:20) 82 + 123 = 205 Lines

