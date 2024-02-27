VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'//-- Name:    ExToolTip Class 1.0
'//-- Author : Mario Flores Gonzalez
'//-- E:mail : sistec_de_juarez@hotmail.com

'//-- Using this code in commercial software, it's not prohibited,
'     however, selling or asking money for parcial or complete part of the code it's prohibited,
'     please help keep this code Open Source.

'     You are welcome to upload this code on any site, as long as you keep the original credits.


'//-- Updates at http://www.geocities.com/sistec_de_juarez/ExTooltip/


'//--  Special Thanks to:  Paul Caton (Paul_Caton@hotmail.com)
'      for providing Subclass routine



Option Explicit

'==================================================================================================
' Apis Declarations
'==================================================================================================

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long


'==================================================================================================
'Subclasser declarations
'==================================================================================================

Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const CODE_LEN               As Long = 197                                      'Length of the machine code in bytes
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aBuf(1 To CODE_LEN)       As Byte                                            'Code buffer byte array
Private sc_pCWP                      As Long                                            'Address of the CallWindowsProc
Private sc_pEbMode                   As Long                                            'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                      As Long                                            'Address of the SetWindowsLong function


'==================================================================================================
'Class Public Enumerations
'==================================================================================================

Public Enum ttIconSize
    TTIcon16 = 16            '//-- 16x16 Icon Size
    TTIcon24 = 24            '//-- 24x24 Icon Size
    TTIcon32 = 32            '//-- 32x32 Icon Size
    TTIcon48 = 48            '//-- 48x48 Icon Size
    TTIcon72 = 72            '//-- 72x72 Icon Size
End Enum

Public Enum ttIcon
    TTI_ERROR = 1           '//--Predefined Error Icon
    TTI_INFO = 2            '//--Predefined Info Icon
    TTI_WARNING = 3         '//--Predefined Warning Icon
End Enum

Public Enum ttStyleEnum
    TTStandard               '//-- Classic Rectanlge Tooltip
    TTBalloon                '//-- Ballon Style Tooltip
End Enum

Public Enum ttFillStyle
    TTNrml = 1               '//-- Normal (Solid) Fill.
    TTGdtH = 2               '//-- Gradient Fill Vertical.
    TTGdtV = 3               '//-- Gradient Fill Horizontal.
    TTPict = 4               '//-- Picture BackGround.
End Enum

'==================================================================================================
'THE POINTAPI STRUCTURE
'--------------------------------------------------------------------------------------------------
'The POINTAPI structure defines the x- and y-coordinates of a point.
'==================================================================================================

Private Type POINTAPI
    X As Long
    Y As Long
End Type

'==================================================================================================
'THE RECT STRUCTURE
'--------------------------------------------------------------------------------------------------
'The RECT structure defines the coordinates of the upper-left and lower-right corners of a rectangle.
'==================================================================================================

Private Type RECT
    Left   As Long
    TOP    As Long
    Right  As Long
    Bottom As Long
End Type

'==================================================================================================
'THE RGBQUAD STRUCTURE
'--------------------------------------------------------------------------------------------------
'The RGBQUAD structure describes a color consisting of relative intensities of red, green, and blue.
'==================================================================================================

Private Type RGBQUAD
    rgbBlue     As Byte
    rgbgreen    As Byte
    rgbred      As Byte
    rgbReserved As Byte
End Type

'==================================================================================================
'THE BITMAPINFOHEADER STRUCTURE
'--------------------------------------------------------------------------------------------------
'The BITMAPINFOHEADER structure contains information about the dimensions and color format
'of a device-independent bitmap (DIB).
'==================================================================================================

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

'==================================================================================================
'THE BITMAPINFO STRUCTURE
'--------------------------------------------------------------------------------------------------
'The BITMAPINFO structure defines the dimensions and color information for a Windows
'device-independent bitmap (DIB).
'==================================================================================================

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'==================================================================================================
'THE EXWINDOW STRUCTURE
'--------------------------------------------------------------------------------------------------
'The EXWINDOW structure Handles the API Created Window ToolTip information
'==================================================================================================

Private Type EXWINDOW
    hwnd    As Long
    X       As Long
    Y       As Long
    aExt    As Long
    bExt    As Long
    Width   As Long
    Height  As Long
    Hover   As Boolean
    Flag    As Boolean
    Visible As Boolean
    Alpha   As Byte
    hRect   As RECT
    Title   As String
    Text    As String
    TextH   As Long
    TextW   As Long
    Icon    As Long
    Pic     As StdPicture
End Type


'==================================================================================================
'ExTooltip Class Private Constants
'==================================================================================================

'=====================================================
'WINDOW STYLES
Private Const GWL_STYLE         As Long = -16
Private Const GWL_EXSTYLE       As Long = -20
'=====================================================

'=====================================================
'WINDOWS SUBCLASS MESSAGES
Private Const WM_DESTROY        As Long = &H2
Private Const WM_TIMER          As Long = &H113
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSEACTIVATE  As Long = &H21

'=====================================================


'=====================================================
'EXTENDED WINDOW STYLES
Private Const WS_EX_TOOLWINDOW  As Long = &H80&
Private Const WS_EX_CLIENTEDGE  As Long = &H200
'=====================================================

'=====================================================
'WINDOWS STYLES
Private Const WS_BORDER         As Long = &H800000
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOZORDER      As Long = &H4
'=====================================================

'=====================================================
'TEXT FORMAT CONST
Private Const DT_CALCRECT       As Long = &H400
'=====================================================

'=====================================================
'SHADOW CONSTS
Private Const CS_DROPSHADOW     As Long = &H20000
Private Const GCL_STYLE         As Long = (-26)
'=====================================================

'=====================================================
'AREA REGION CONSTS
Private Const RGN_DIFF          As Long = 4
'=====================================================

'=====================================================
'DRAWICONEX CONSTS
Private Const DI_NORMAL         As Long = &H3
'=====================================================

'=====================================================
'VARIABLES USED IN PROCESS
'=====================================================
Private iFnt          As IFont        '//-- Font Value
Private m_fnt         As IFont        '//-- Font Value
Private hFntOld       As Long         '//-- Font Value
Private CPOS          As POINTAPI     '//-- Coordinate Values of the Position of the Mouse
Private CMOS          As POINTAPI     '//-- Memory Coordinate Values of the Position of the Mouse
Private bTime         As Long         '//-- API Timer Variable Counter
Private LWnd          As Long         '//-- Control Handle Window
Private sc_aSubData() As tSubData     '//-- Array of Subclass Values
Private M_ExWindow    As EXWINDOW     '//-- Ex Window Tooltip Values
Private mToolTipStyle As ttStyleEnum
Private mIconSize     As ttIconSize
Private mTextColor    As Long
Private mBackColor    As Long
Private mGCEnd        As Long
Private mGCStart      As Long
Private mBackStyle    As Long
Private mDelayTime    As Long
Private mShadow       As Boolean
Private mKillTime     As Long
Private mAlive        As Boolean
Private mCP(100, 255) As Byte         '//-- Holds the percent calculation
Private mCC(255, 255) As Byte         '//-- Holds the Add Colors math

'=====================================================
'VARIABLES USED TO DRAW TEMP DC
'=====================================================
Private m_Hdc      As Long
Private m_ThDC     As Long
Private m_hBmp     As Long             '//-- Temporary DC (Where All The Drawings are made)
Private m_hBmpOld  As Long             '//-- It Also Provides an Anti-Flicker Draw
Private m_lWidth   As Long
Private m_lHeight  As Long
'=====================================================

'=====================================================
'VARIABLES USED TO DRAW REGION
Private rgnNorm    As Long
Private rgn2       As Long
Private rgn1       As Long
'=====================================================


'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
Select Case uMsg

   
'==============================================================================================
'==============================================================================================
'The WM_MOUSEMOVE message is posted to a window when the cursor moves.
'If the mouse is not captured, the message is posted to the window that contains the cursor.
'Otherwise, the message is posted to the window that has captured the mouse.
'==============================================================================================
'==============================================================================================
   
   Case WM_MOUSEMOVE


       '//-- If the User is Over the Calling Control Turn 'ON' The Hover Flag.
            If InsideArea(LWnd) Then
                 If M_ExWindow.Hover = False Then
                    M_ExWindow.Hover = True        '/-- If The Mouse Enter The Control Area
                    SetTimer LWnd, 1, 1, 0         '/-- Start API Timer
                 End If
            Else
                    M_ExWindow.Hover = False       '/-- If The Mouse Left The Control Area
            End If
      
'==============================================================================================
'==============================================================================================
'The WM_TIMER message is posted to the installing thread's message queue when a timer expires.
'You can process the message by providing a WM_TIMER case in the window procedure.
'Otherwise, the default window procedure will call the TimerProc callback function specified
'in the call to the SetTimer function used to install the timer.
'==============================================================================================
'==============================================================================================
    
   Case WM_TIMER
       
        
        If M_ExWindow.Hover And Not M_ExWindow.Flag Then
            
            CMOS.X = CPOS.X     '//-- Assign Memory X POS.
            CMOS.Y = CPOS.Y     '//-- Assign Memory Y POS.
            GetCursorPos CPOS   '//-- Capture New X-Y Mouse Coordinates.

            '//-- ToolTips Don't Popup Just When the mouse is moving,
            '     they do so until the mouse is "still" for a few moments.
            '     Lets do it until that's 1 second.
            If (CMOS.X = CPOS.X) And (CMOS.Y = CPOS.Y) Then
                bTime = bTime + 1
                If bTime > mDelayTime Then    '//-- At This Point The Mouse Cursor Has been still for 1 second in the same position
                   bTime = 0                  '//-- Lets Reset the Timer, To count the time that the Tooltip is going to be Alive.
                   If mAlive = True Then      '//-- Ensure no child controls are on same Mouse Coords(in case Lwnd has Child Controls)
                      CreateToolTip           '//-- Begin The ToolTip Drawing.
                      M_ExWindow.Flag = True  '//-- We Now Know That We are ready to draw the ToolTip.
                   End If
                   Exit Sub
                End If
            Else
                bTime = 0                  '//-- At This Point, the Mouse Cursor has not been still for 1 second in the same position, let's reset the timer.
            End If

        End If

                
        If InsideArea(LWnd) = False Then
             KillTimer LWnd, 1                  '//-- Destroy the API Timer
             DestroyWindow M_ExWindow.hwnd      '//-- Destroy the ToolTip Window when the Mouse Leaves the Control
             M_ExWindow.hwnd = 0                '//-- Reset the HWnd    Value
             M_ExWindow.Hover = False           '//-- Reset the Hover   Value
             M_ExWindow.Flag = False            '//-- Reset the Flag    Value
             Class_Initialize                   '//-- Reset Default Values for next Instance.
        ElseIf bTime >= mKillTime Then
             DestroyWindow M_ExWindow.hwnd      '//-- AT This Point, the Tool tip as past out the visible time, and must be Destroyed.
             Class_Initialize                   '//-- Reset Default Values for next Instance.
        ElseIf M_ExWindow.Hover And M_ExWindow.hwnd <> 0 Then
              bTime = bTime + 1
        End If
                

'==============================================================================================
'==============================================================================================
'The WM_MOUSEACTIVATE message is sent when the cursor is in an inactive window and the user
'presses a mouse button.
'==============================================================================================
'==============================================================================================
   
   Case WM_MOUSEACTIVATE
        
        '//-- At This Point, the Tool tip has to be destroyed,
        '     because the user click on the calling control (natural Tooltip behavior).
                              
             KillTimer LWnd, 1                  '//-- Destroy the API Timer
             DestroyWindow M_ExWindow.hwnd      '//-- Send WM_DESTROY Message.
             M_ExWindow.Hover = False           '//-- Reset the Hover Value
             Class_Initialize                   '//-- Reset Default Values for next Instance.
             mAlive = True                      '//-- Force User To Move Mouse Outside Control Before another Tooltip is Displayed.
                

'==============================================================================================
'==============================================================================================
'The WM_DESTROY message is sent when a window is being destroyed. It is sent to the window procedure
'of the window being destroyed after the window is removed from the screen.
'==============================================================================================
'==============================================================================================
   
   Case WM_DESTROY
            
        mAlive = False
            
End Select
   
   
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function InIDE() As Boolean
  Debug.Assert zSetTrue(InIDE)
End Function

'Start subclassing the passed window handle
Private Function Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sSubCode                As String                                                 'Subclass code string
Const GMEM_FIXED              As Long = 0                                               'Fixed memory GlobalAlloc flag
Const PAGE_EXECUTE_READWRITE  As Long = &H40&                                           'Allow memory to execute without violating XP SP2 Data Execution Prevention
Const PATCH_01                As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02                As Long = 68                                              'Address of the previous WndProc
Const PATCH_03                As Long = 78                                              'Relative address of SetWindowsLong
Const PATCH_06                As Long = 116                                             'Address of the previous WndProc
Const PATCH_07                As Long = 121                                             'Relative address of CallWindowProc
Const PATCH_0A                As Long = 186                                             'Address of the owner object
Const FUNC_CWP                As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
Const FUNC_EBM                As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL                As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER                As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5                As String = "vba5"                                        'Location of the EbMode function if running VB5
Const MOD_VBA6                As String = "vba6"                                        'Location of the EbMode function if running VB6

'If it's the first time through here..
  If sc_aBuf(1) = 0 Then

'Build the hex pair subclass string
    sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3"
    
'Convert the string from hex pairs to bytes and store in the machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, I, 2))                                   'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If InIDE Then                                                                       'If we're running in the VB IDE
      sc_aBuf(16) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_aBuf(17) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                        'Get the address of EbMode in vba6.dll
      If sc_pEbMode = 0 Then                                                            'Found?
        sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                      'VB5 perhaps
      End If
    End If
    
    Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                            'Patch the address of this object instance into the static machine code buffer
    
    sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                             'Get the address of the CallWindowsProc function
    sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                             'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, I)           'Mark memory as executable
    Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                           'Copy the machine code from the static byte array to the code array in sc_aSubData
    
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    
    Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                                     'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                                        'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                                        'Patch the relative address of the CallWindowProc api function
  End With
End Function

'Stop all subclassing
Private Sub StopAll()
On Error Resume Next

  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call StopWnd(.hwnd)                                                             'Subclass_Stop
      End If
    End With
    
    I = I - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub StopWnd(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function



Public Property Let Alive(ByVal vData As Boolean)
   mAlive = vData
End Property

Public Property Get Alive() As Boolean
   Alive = mAlive
End Property

Public Property Let Shadow(ByVal vData As Boolean)
   mShadow = vData
End Property

Public Property Get Shadow() As Boolean
   Shadow = mShadow
End Property

Public Property Get Font() As IFont
   Set Font = m_fnt
End Property

Public Property Set Font(ByRef fnt As IFont)
   Set m_fnt = fnt    'Defined By System but can change by user choice.
End Property

Public Property Let Font(ByRef fnt As IFont)
   Set m_fnt = fnt
End Property

Public Property Let IconSize(ByVal vData As ttIconSize)
   mIconSize = vData
End Property

Public Property Get IconSize() As ttIconSize
   IconSize = mIconSize
End Property

Public Property Let ToolTipStyle(ByVal Style As ttStyleEnum)
   mToolTipStyle = Style
End Property

Public Property Get ToolTipStyle() As ttStyleEnum
   ToolTipStyle = mToolTipStyle
End Property

Public Property Let TextColor(ByVal vcolor As Long)
   mTextColor = vcolor
End Property

Public Property Get TextColor() As Long
   TextColor = mTextColor
End Property

Public Property Let GradientColorStart(ByVal vcolor As Long)
   mGCStart = vcolor
End Property

Public Property Get GradientColorStart() As Long
   GradientColorStart = mGCStart
End Property

Public Property Let GradientColorEnd(ByVal vcolor As Long)
   mGCEnd = vcolor
End Property

Public Property Get GradientColorEnd() As Long
   GradientColorEnd = mGCEnd
End Property

Public Property Let BackColor(ByVal vcolor As Long)
   mBackColor = vcolor
End Property

Public Property Get BackColor() As Long
   BackColor = mBackColor
End Property

Public Property Let BackStyle(ByVal bStyle As Long)
   mBackStyle = bStyle
End Property

Public Property Get BackStyle() As Long
   BackStyle = mBackStyle
End Property

Public Property Let DelayTime(ByVal Time As Long)
   mDelayTime = Time
End Property

Public Property Get DelayTime() As Long
   DelayTime = mDelayTime
End Property

Public Property Let KillTime(ByVal Time As Long)
   mKillTime = Time
End Property

Public Property Get KillTime() As Long
   KillTime = mKillTime
End Property

Public Property Set Picture(xPic As StdPicture)
    Set M_ExWindow.Pic = xPic
End Property

Public Property Get Picture() As StdPicture
Set Picture = M_ExWindow.Pic
End Property


'======================================================================
'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long

    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If

End Function
'======================================================================

'======================================================================
'DRAWS A LINE WITH A DEFINED COLOR
Private Sub DrawLine( _
           ByVal X As Long, _
           ByVal Y As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal cHdc As Long, _
           ByVal Color As Long)

    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim Outline As Long
    Dim POS     As POINTAPI

    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    
        MoveToEx cHdc, X, Y, POS
        LineTo cHdc, Width, Height
          
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1

End Sub
'======================================================================

'======================================================================
'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawFillRectangle(ByRef dRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)

Dim hBrush    As Long
   
   hBrush = CreateSolidBrush(GetLngColor(Color))
   FillRect MyHdc, dRect, hBrush
   DeleteObject hBrush

End Sub
'======================================================================

'======================================================================
'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawRgnRectangle(ByRef hRect As Long, ByVal Color As Long, ByVal MyHdc As Long)

Dim hBrush As Long

   hBrush = CreateSolidBrush(GetLngColor(Color))
   FrameRgn MyHdc, hRect, hBrush, 1, 1
   DeleteObject hBrush

End Sub
'======================================================================

'======================================================================
'DRAWS A IMAGE AREA OF AN SPECIFIED DC
Private Sub DrawBkgImgDC(ByRef dRect As RECT, ByVal m_Picture As Long)

Dim Brush      As Long
Dim origBrush  As Long

       If m_Picture = 0 Then Exit Sub                        '//-- In Case No Picture is Choosen
              
       Brush = CreatePatternBrush(m_Picture)                 '//-- Use Pattern Picture Draw
       origBrush = SelectObject(m_Hdc, Brush)
       
       PatBlt m_Hdc, 0, 0, dRect.Right - dRect.Left, dRect.Bottom - dRect.TOP, vbPatCopy
         
       SelectObject m_Hdc, origBrush
       DeleteObject Brush
       
End Sub
'======================================================================


Private Function CreateNewWindow() As Long

    
    If M_ExWindow.hwnd <> 0 Then
        KillTimer LWnd, 1                       '//-- Destroy the API Timer
        DestroyWindow M_ExWindow.hwnd           '//-- Destroy the ToolTip Window when the Mouse Leaves the Control
        M_ExWindow.hwnd = 0                     '//-- Reset the HWnd    Value
        CreateNewWindow = -1                    '//-- If The ToolTip is Visible Exit Function
        Exit Function
    End If
    
    M_ExWindow.Hover = False                    '//-- Reset the Hover Value
    M_ExWindow.Flag = False                     '//-- Reset the Flag  Value
   
    Call Start(LWnd)                            '//-- SubClass The Calling Control To Initialize a Timer.
    Call AddMsg(LWnd, WM_MOUSEMOVE, MSG_AFTER)
    Call AddMsg(LWnd, WM_TIMER, MSG_AFTER)
    Call AddMsg(LWnd, WM_MOUSEACTIVATE, MSG_AFTER)

End Function


Private Sub DrawToolTipBody()
     
Dim ClientRect As RECT
Dim hIcon      As Long
        
    '================================================================
    'DRAW THE FILL OF THE TOOLTIP (BACKGROUND,TEXT,ICON)
    '================================================================
        
    With M_ExWindow
       
           '============================================================================================
           '/-- Get The ToolTip Area Cords.
           GetClientRect .hwnd, ClientRect
           '============================================================================================
           '/-- Solid BackGround Fill
           If mBackStyle = TTNrml Then
              DrawFillRectangle ClientRect, mBackColor, m_Hdc
           '============================================================================================
           '/-- Picture BackGround Fill
           ElseIf mBackStyle = TTPict Then
              If Not Nothing Is M_ExWindow.Pic Then DrawBkgImgDC ClientRect, M_ExWindow.Pic.handle
           '============================================================================================
           '/-- Gradient BackGround Fill
           Else
              DrawGradient mGCEnd, mGCStart, ClientRect, m_Hdc, IIf(mBackStyle = TTGdtH, True, False)
           End If
           '============================================================================================
           '/-- Draw the Tooltip Body Border.
           DrawRgnRectangle rgn2, mTextColor, m_Hdc
               
     End With
        

     DeleteObject rgnNorm  '//-- Delete Rgn
     DeleteObject rgn2     '//-- Delete Rgn
     SetBkMode m_Hdc, 1    '//-- Sets the background mix mode (TRANSPARENT Background remains untouched)
     
     '============================================================================================
     '/-- Draw The Icon
     With M_ExWindow
        If .Icon <> 0 Then
            Select Case .Icon
                Case TTI_ERROR, TTI_INFO, TTI_WARNING
                    If mIconSize > TTIcon24 Then
                       ExtractIconEx GetSystemDir & "\comctl32.dll", .Icon, hIcon, 0, 1
                    Else
                       ExtractIconEx GetSystemDir & "\comctl32.dll", .Icon, 0, hIcon, 1
                    End If
                    DrawIconEx m_Hdc, 12, 8 + IIf(.aExt > 0, .aExt, 0), hIcon, mIconSize, mIconSize, 0, 0, DI_NORMAL
                    DestroyIcon hIcon
                Case Else
                    DrawIconEx m_Hdc, 12, 8 + IIf(.aExt > 0, .aExt, 0), .Icon, mIconSize, mIconSize, 0, 0, DI_NORMAL
            End Select
        End If
        '============================================================================================
        '/-- Draw the Title.
        DrawTexto .Title, True
        '============================================================================================
        '/-- Draw the Text.
        DrawTexto .Text, False, .Title
     End With
     
  End Sub

'======================================================================
'RETRIVES THE SYSTEM DIRECTORY
Private Function GetSystemDir() As String
Dim sSave   As String
Dim Ret     As Long

'Create a buffer
 sSave = Space(255)
'Get the system directory
 Ret = GetSystemDirectory(sSave, 255)
'Remove all unnecessary chr$(0)'s
 GetSystemDir = Left$(sSave, Ret)

End Function
'======================================================================


Public Sub ShowToolTip(ByVal aWnd As Long, ByVal Title As String, ByVal Text As String, ByVal cIcon As Long, Optional TransparentAlpha As Long = 100)
     
  With M_ExWindow
     
          Call StopAll                           '//-- Avoid having the same Window Subclassed again.
                                             
          LWnd = aWnd                            '//-- Pass Parameters to memory
          .Title = Title                         '//-- Pass Parameters to memory
          .Text = Text                           '//-- Pass Parameters to memory
          .Alpha = TransparentAlpha              '//-- Pass Parameters to memory
          .TextH = 0                             '//-- Reset Height Value
          .TextW = 0                             '//-- Reset Widht Value
          .Icon = cIcon                          '//-- Pass Parameters to memory
          Call SetToolTipRect(True)              '//-- Calculate Title Dimensions
          Call SetToolTipRect(False)             '//-- Calculate Text Dimensions
          .Height = M_ExWindow.TextH + IIf(Len(Trim(M_ExWindow.Title)) > 0, 13, 0)
          .Width = M_ExWindow.TextW + 20
                    
          If CreateNewWindow() <> -1 Then mAlive = True   '/-- Start Core Work
                                        
  End With
       
End Sub


Private Function CreateToolTip() As Boolean
    
    Dim XL As Long
    Dim YL As Long
       
   
    With M_ExWindow
           
        'The CreateWindowEx function creates an overlapped, pop-up, or child window with an extended style;
        'otherwise, this function is identical to the CreateWindow function.
        '//-- FYI : #32770: Is the class for dialog boxes.
        '//-- Create a "MFGTOOLTIP" window where the mouse cursor is located.
        '//-- FYI : WS_EX_TOOLWINDOW disables the ShowInTaskBar Window Style.
        .hwnd = CreateWindowEx(WS_EX_TOOLWINDOW, "#32770", "MFG-EXTOOLTIP", 0, CPOS.X, CPOS.Y, .Width, .Height, 0, 0, App.hInstance, 0)
            
        If .hwnd = 0 Then Exit Function '//-- hWnd not found, system error.
                          
        .X = CPOS.X        '//-- Assign X Cursor Coords to Mem.
        .Y = CPOS.Y        '//-- Assign Y Cursor Coords to Mem.
    
        XL = .Width + .X   '//-- ToolTip Rectangle Right-Side.
        YL = .Height + .Y  '//-- ToolTip Rectangle Bottom-Side.
        
        .Width = .Width
        .Height = .Height
        
        '============================================================================
        '//-- Calculate if ToolTip fits on Screen X Coords.
        '============================================================================
        If XL > (Screen.Width / Screen.TwipsPerPixelX) Then
            .bExt = IIf(mToolTipStyle = TTStandard, .aExt = -2, -17)
            .X = .X - .Width
        Else
            .bExt = IIf(mToolTipStyle = TTStandard, .aExt = 2, 17)
            .X = .X + 1
        End If
               
        '============================================================================
        '//-- Calculate if ToolTip fits on Screen Y Coords.
        '============================================================================
        If YL > (Screen.Height / Screen.TwipsPerPixelY) Then
            .aExt = IIf(mToolTipStyle = TTStandard, .aExt = -2, -17)
            .Y = .Y - .Height
        ElseIf .Y - .Height - 20 < 0 Then
            .aExt = IIf(mToolTipStyle = TTStandard, .aExt = 2, 17)
            .Y = .Y + 1
        Else
        '//-- FYI : By default ToolTips always appear on the TOP-LEFT of the screen.
            .aExt = IIf(mToolTipStyle = TTStandard, .aExt = -2, -17)
            .Y = .Y - .Height
        End If
         
        
        '============================================================================
        '//-- Change the window styles, so it can look like a real Tooltip.
        If ChangeWindowStyles(.hwnd) = 0 Then Exit Function
        '============================================================================
        
        '============================================================================
        '//-- Add Shadow to the Tooltip Window, if selected. <Only Layered OS>
        SetClassLong .hwnd, GCL_STYLE, IIf(mShadow = True, GetClassLong(.hwnd, GCL_STYLE) Or CS_DROPSHADOW, 0)
        '============================================================================
                  
        GetWindowRect .hwnd, .hRect                                                    '//-- Get The New ToolTip Coords.
            
        With .hRect
            m_Hdc = ThDC(.Right - .Left, .Bottom - .TOP)                               '//-- Create a Temporary DC, with given Coordinates.
            DCImageDraw m_Hdc, 0, 0, .Right - .Left, .Bottom - .TOP                    '//-- Function that Draws the Tooltip in the temp DC.
            ShowWindow M_ExWindow.hwnd, 4                                              '//-- Displays a window in its most recent size and position. The active window remains active.
            '===============================================================================
            pDraw GetDC(M_ExWindow.hwnd), 0, 0, .Right - .Left, .Bottom - .TOP, 0, 0   '//-- Draw The Temp DC into the Real Window DC.
            '===============================================================================
        End With
        
        pDestroy                '//-- Destroy the Temporary DC.
        ReleaseDC m_Hdc, .hwnd  '//-- The ReleaseDC function releases the window device context (DC), freeing it for use by other applications.
        
        Call Start(.hwnd)                            '//-- SubClass The Tooltip.
        Call AddMsg(.hwnd, WM_DESTROY, MSG_AFTER)
              
    End With
        
       
    CreateToolTip = True         '//-- The MFG ExTooltip was successfully created.
   
       
End Function



Private Sub DCImageDraw( _
            ByVal DstDC As Long, _
            ByVal DstX As Long, _
            ByVal DstY As Long, _
            ByVal DstW As Long, _
            ByVal DstH As Long)
    
    Dim INFO       As BITMAPINFO  'INFO
    Dim Data1()    As RGBQUAD     'RGB
    Dim Data2()    As RGBQUAD     'RGB
    Dim SrcDC      As Long        'DC
    Dim DataDC     As Long        'DC
    Dim DataBmp    As Long        'BMP
    Dim DataObj    As Long        'OBJ
    Dim I          As Long        'COUNTER
    Dim F          As Long        'COUNTER
    Dim A          As Long        'COUNTER
    Dim AlphaD     As Long        'ALPHA PERCENT
    Dim AlphaS     As Long        'ALPHA PERCENT
    
    
    With INFO.bmiHeader
        .biSize = Len(INFO.bmiHeader)   '//-- Specifies the number of bytes required by the structure.
        .biWidth = DstW                 '//-- Specifies the width of the bitmap, in pixels.
        .biHeight = DstH                '//-- Specifies the height of the bitmap, in pixels.
        .biPlanes = 1                   '//-- Specifies the number of planes for the target device. This value must be set to 1.
        .biBitCount = 32                '//-- Specifies the number of bits per pixel.
        .biCompression = 0              '//-- Specifies the type of compression for a compressed bottom-up bitmap.
        .biClrUsed = 0                  '//-- Specifies the number of color indexes in the color table that are actually used by the bitmap.
        .biClrImportant = 0             '//-- Specifies the number of color indexes that are required for displaying the bitmap.
        .biSizeImage = DstW * DstH      '//-- Specifies the size, in bytes, of the image.
    End With
       
    '-----------------------------------------------------------------
    '//-- Fill the Color Table
    '-----------------------------------------------------------------
    For I = 0 To 255
        For F = 0 To 255
            A = I + F
            If A > 255 Then A = 255     '//-- Sanity Color Check
            mCC(I, F) = A
        Next F
    Next I
    
    '-----------------------------------------------------------------
    '//-- Fill the Percent Table
    '-----------------------------------------------------------------
    For I = 0 To 100
        For F = 0 To 255
            mCP(I, F) = F / 100 * I     '//-- Alpha Percent (255 RGB)
        Next F
    Next I
    
    CreateBallonRegion
    DrawToolTipBody
    
    SrcDC = CreateCompatibleDC(0)       '//-- Create a Temp DC.
       
    With M_ExWindow.hRect
       DataObj = CaptureWindow(0, .Left, .TOP, .Right - .Left, .Bottom - .TOP)
    End With
    
    SelectObject SrcDC, DataObj
    DeleteObject DataObj

    GetDCDBitsFromDevice DataDC, DataBmp, DataObj, SrcDC, DstDC, DstH, DstW, Data1, INFO, False
    GetDCDBitsFromDevice DataDC, DataBmp, DataObj, SrcDC, DstDC, DstH, DstW, Data2, INFO, True
    
   '=================================================================================================
   'A device-independent bitmap consists of two distinct parts: a BITMAPINFO structure describing
   'the dimensions and colors of the bitmap, and an array of bytes defining the pixels of the bitmap.
   'The bits in the array are packed together, but each scan line must be padded with zeroes to end
   'on a LONG boundary. If the height is positive, the origin of the bitmap is the lower-left corner.
   'If the height is negative, the origin is the upper-left corner.
   '=================================================================================================
   
   For I = 0 To DstW - 1
        For F = 0 To DstH - 1
            AlphaS = 100 - M_ExWindow.Alpha
            AlphaD = M_ExWindow.Alpha
            Data2(I, F).rgbBlue = mCC(mCP(AlphaD, Data1(I, F).rgbBlue), mCP(AlphaS, Data2(I, F).rgbBlue))
            Data2(I, F).rgbgreen = mCC(mCP(AlphaD, Data1(I, F).rgbgreen), mCP(AlphaS, Data2(I, F).rgbgreen))
            Data2(I, F).rgbred = mCC(mCP(AlphaD, Data1(I, F).rgbred), mCP(AlphaS, Data2(I, F).rgbred))
        Next F
   Next I

   '=================================================================================================
   ' The SetDIBitsToDevice function sets the pixels in the specified rectangle on the device that is
   ' associated with the destination device context using color data from a device-independent bitmap (DIB).
   '=================================================================================================
    
   '//-- Draw the new buffer Bits into the Destination DC.
    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0, 0), INFO, 0
    
    Erase Data1                                     'Free Array
    Erase Data2                                     'Free Array
    DeleteDC SrcDC                                  'Free Mem
             
End Sub

Private Sub GetDCDBitsFromDevice( _
            ByRef DataHDC As Long, _
            ByRef DataBmp As Long, _
            ByRef DataObj As Long, _
            ByRef DataSRC As Long, _
            ByRef DataDST As Long, _
            ByRef DataHGT As Long, _
            ByRef DataWDT As Long, _
            ByRef DataC() As RGBQUAD, _
            ByRef DataINF As BITMAPINFO, _
            ByVal DataXOR As Boolean)
    
    DataHDC = CreateCompatibleDC(DataSRC)                                                   '//--Create a memory device context (DC) compatible with the specified device.
    DataBmp = CreateCompatibleBitmap(DataDST, DataWDT, DataHGT)                             '//--Create a bitmap compatible with the device that is associated with the specified device context.
    DataObj = SelectObject(DataHDC, DataBmp)                                                '//--Select object into the specified device context.
    ReDim DataC(0 To DataWDT - 1, 0 To DataHGT - 1)                                         '//--Pointer to a new safe array bound structure containing the new array bound.
    BitBlt DataHDC, 0, 0, DataWDT, DataHGT, IIf(DataXOR, DataSRC, DataDST), 0, 0, vbSrcCopy '//--Perform a bit-block transfer of the color data corresponding to the rectangle of pixels from the specified DC.
    GetDIBits DataHDC, DataBmp, 0, DataHGT, DataC(0, 0), DataINF, 0                         '//--Retrieve the bits of the specified bitmap and copy them into a buffer using the specified format.
    DataBmp = SelectObject(DataHDC, DataObj)                                                '//--Select object into the specified device context. The new object replaces the previous object of the same type.
    DeleteObject DataBmp                                                                    '//--Delete logical bitmap, freeing all system resources associated with the object. After the object is deleted, the specified handle is no longer valid.
    DeleteDC DataHDC                                                                        '//--Delete the specified device context (DC)

End Sub


Private Sub CreateBallonRegion()

Dim H       As Long
Dim W       As Long
Dim I       As Long


    H = IIf(M_ExWindow.aExt <> 0, Abs(M_ExWindow.aExt) + M_ExWindow.Height, M_ExWindow.Height)
    W = M_ExWindow.Width
    
    If mToolTipStyle = TTStandard Then
       rgn2 = CreateRectRgn(0, 0, W, H)
      Exit Sub
    End If
    
    
    rgnNorm = CreateRectRgn(0, 0, W, H)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
       

    With M_ExWindow
              
        SetAreaRegion IIf(.bExt < 0, .Width, 0), IIf(.aExt > 0, 0, H), IIf(.bExt < 0, .Width - 17, 18), IIf(.aExt > 0, 17, H - 17), rgn1, rgn2, rgnNorm, True
        SetAreaRegion IIf(.bExt < 0, .Width - 35, 35), IIf(.aExt > 0, 0, H), IIf(.bExt < 0, W - .Width, W), IIf(.aExt > 0, 17, H - 17), rgn1, rgn2, rgnNorm, False
    
        For I = 0 To IIf(.bExt < 0, 17, 15)
            SetAreaRegion IIf(.bExt < 0, W - 18 - I, 19 + I), IIf(.aExt > 0, I, H - I), IIf(.bExt < 0, W - 35, 35), IIf(.aExt > 0, I + 1, H - I - 1), rgn1, rgn2, rgnNorm, IIf((I + 1) Mod 2 = 1, True, False)
        Next I
         
      
        '//-- TOP LEFT CORNER ====================================
        SetAreaRegion 0, IIf(.aExt > 0, 17, 0), 5, IIf(.aExt > 0, 18, 1), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 0, IIf(.aExt > 0, 18, 1), 4, IIf(.aExt > 0, 19, 2), rgn1, rgn2, rgnNorm, False
        SetAreaRegion 0, IIf(.aExt > 0, 19, 2), 2, IIf(.aExt > 0, 20, 3), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 0, IIf(.aExt > 0, 20, 3), 1, IIf(.aExt > 0, 22, 5), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
        
        '//-- BOTTOM LEFT CORNER =================================
        SetAreaRegion 0, H - 5 + IIf(.aExt > 0, 0, .aExt), 1, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 1, H - 3 + IIf(.aExt > 0, 0, .aExt), 2, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        SetAreaRegion 2, H - 2 + IIf(.aExt > 0, 0, .aExt), 3, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 3, H + IIf(.aExt > 0, 0, .aExt), 5, H - 1 + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
        
        '//-- TOP RIGHT CORNER ===================================
        SetAreaRegion W - 5, IIf(.aExt > 0, 17, 0), W, IIf(.aExt > 0, 18, 1), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 3, IIf(.aExt > 0, 18, 1), W, IIf(.aExt > 0, 19, 2), rgn1, rgn2, rgnNorm, False
        SetAreaRegion W - 2, IIf(.aExt > 0, 19, 2), W, IIf(.aExt > 0, 20, 3), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 1, IIf(.aExt > 0, 20, 3), W, IIf(.aExt > 0, 22, 5), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
    
        '//-- BOTTOM RIGHT CORNER ================================
        SetAreaRegion W - 1, H - 5 + IIf(.aExt > 0, 0, .aExt), W, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 2, H - 3 + IIf(.aExt > 0, 0, .aExt), W - 1, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        SetAreaRegion W - 3, H - 2 + IIf(.aExt > 0, 0, .aExt), W - 2, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 5, H - 1 + IIf(.aExt > 0, 0, .aExt), W - 3, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
        
    End With
        
    Call SetWindowRgn(M_ExWindow.hwnd, rgnNorm, True)
  

End Sub

Private Sub SetAreaRegion(ByRef A1 As Long, ByRef B1 As Long, ByRef A2 As Long, ByRef B2 As Long, ByRef R1 As Long, ByRef R2 As Long, ByRef RN As Long, ByVal aOR As Boolean)
   
    R1 = CreateRectRgn(A1, B1, A2, B2)
    CombineRgn IIf(aOR, R2, RN), IIf(aOR, RN, R2), R1, RGN_DIFF
    DeleteObject R1
    
End Sub

'======================================================================
'DRAWS THE TEXT ON THE TOOLTIP
Private Sub DrawTexto(ByVal ThisText As String, ByVal Title As Boolean, Optional TitleST As String)
 Dim TextRect As RECT
             
      Set iFnt = Font                             '//--New Font
      iFnt.Bold = IIf(Title, True, False)         '//--Bold State
      hFntOld = SelectObject(m_Hdc, iFnt.hFont)   '//--Use the New Font
          
      DrawText m_Hdc, ThisText, Len(ThisText), TextRect, DT_CALCRECT  '//-- calculate text Dimensions using DT_CALCRECT

      With TextRect
       
           If Title Then
              .TOP = 1 + IIf(M_ExWindow.Icon <> 0, (mIconSize / 2), 0) + IIf(M_ExWindow.aExt > 0, M_ExWindow.aExt, 0)
              .Left = (20 + IIf(M_ExWindow.aExt > 0, M_ExWindow.aExt, 0))
           Else
              DrawText m_Hdc, TitleST, Len(TitleST), TextRect, DT_CALCRECT
              .TOP = IIf(Len(Trim(M_ExWindow.Title)) > 0, .Bottom + IIf(M_ExWindow.Icon <> 0, mIconSize, 3), (.Bottom / 2) - 5) + IIf(M_ExWindow.aExt > 0, M_ExWindow.aExt, 0)
              If Len(Trim(M_ExWindow.Title)) = 0 And M_ExWindow.Icon <> 0 Then .TOP = .TOP + mIconSize
           End If
       
           If M_ExWindow.Icon <> 0 And Title Then
              .Left = .Left + mIconSize
           Else
              .TOP = .TOP + 5
              .Left = 10
           End If
              .Right = .Right + .Left
              .Bottom = .Bottom + .TOP
     
      End With
                
      SetTextColor m_Hdc, GetLngColor(mTextColor)                 '//-- Select Custom Color for Text
      DrawText m_Hdc, ThisText, Len(ThisText), TextRect, &H100    '//-- Draw Text on DC
      SelectObject m_Hdc, hFntOld                                 '//-- Delete the Used Font
      
      
End Sub
'======================================================================

Private Sub SetToolTipRect(ByVal Title As Boolean)
Dim FDC      As Long
Dim TextRect As RECT

      FDC = CreateDC("DISPLAY", "", "", ByVal 0&)  '//--Create a DC
            
      If Title And Len(M_ExWindow.Title) = 0 Then M_ExWindow.Title = " "
      
      Set iFnt = Font                             '//--New Font
      iFnt.Bold = IIf(Title, True, False)         '//--Bold State
      hFntOld = SelectObject(FDC, iFnt.hFont)     '//--Use the New Font

      DrawText FDC, IIf(Title, M_ExWindow.Title, M_ExWindow.Text), Len(IIf(Title, M_ExWindow.Title, M_ExWindow.Text)), TextRect, DT_CALCRECT
      M_ExWindow.TextH = M_ExWindow.TextH + TextRect.Bottom
      M_ExWindow.TextW = IIf(M_ExWindow.TextW > TextRect.Right, M_ExWindow.TextW, TextRect.Right)
                
      If (M_ExWindow.Icon <> 0 And Title) Then
          M_ExWindow.TextW = M_ExWindow.TextW + 25 + mIconSize
          M_ExWindow.TextH = M_ExWindow.TextH + mIconSize
      End If
                       
      SelectObject FDC, hFntOld                   '//--Delete the Used Font
      DeleteDC FDC                                '//--Delete the Used DC
            
End Sub

'======================================================================
'CHECKS-CREATES CORRECT DIMENSIONS OF THE TEMP DC
Private Function ThDC(Width As Long, Height As Long) As Long
   If m_ThDC = 0 Then
      If (Width > 0) And (Height > 0) Then
         pCreate Width, Height
      End If
   Else
      If Width > m_lWidth Or Height > m_lHeight Then
         pCreate Width, Height
      End If
   End If
   ThDC = m_ThDC
End Function
'======================================================================

'======================================================================
'CREATES THE TEMP DC
Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
Dim lhDCC As Long
   pDestroy
   lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
   If Not (lhDCC = 0) Then
      m_ThDC = CreateCompatibleDC(lhDCC)
      If Not (m_ThDC = 0) Then
         m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
         If Not (m_hBmp = 0) Then
            m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
            If Not (m_hBmpOld = 0) Then
               m_lWidth = Width
               m_lHeight = Height
               DeleteDC lhDCC
               Exit Sub
            End If
         End If
      End If
      DeleteDC lhDCC
      pDestroy
   End If
End Sub
'======================================================================

'======================================================================
'DRAWS THE TEMP DC
Private Sub pDraw( _
      ByVal hdc As Long, _
      Optional ByVal XSrc As Long = 0, Optional ByVal YSrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
   If WidthSrc <= 0 Then WidthSrc = M_ExWindow.Width
   If HeightSrc <= 0 Then HeightSrc = M_ExWindow.Height
   BitBlt hdc, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, XSrc, YSrc, vbSrcCopy

End Sub
'======================================================================

'======================================================================
'DESTROYS THE TEMP DC
Private Sub pDestroy()
   If Not m_hBmpOld = 0 Then
      SelectObject m_ThDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If Not m_ThDC = 0 Then
      DeleteDC m_ThDC
      m_ThDC = 0
   End If
   m_lWidth = 0
   m_lHeight = 0
End Sub
'======================================================================

Private Function ChangeWindowStyles(ByVal ThWnd As Long) As Long
  Dim TFlat As Long
  
  '===================================================================================================================
  '//-- Experimenting with some window styles to create a Tooltip look window, this flags give me the best results.
  '===================================================================================================================
  
  TFlat = GetWindowLong(ThWnd, GWL_STYLE)    '//-- Retrieve information about the specified window.
  TFlat = TFlat And WS_BORDER And 1600000    '//-- Parameters of a window that has a flat border.
  SetWindowLongA ThWnd, GWL_STYLE, TFlat     '//-- Change attributes of the specified window.
  
  TFlat = GetWindowLong(ThWnd, GWL_EXSTYLE)  '//-- Retrieve information about the specified window.
  TFlat = TFlat And Not WS_EX_CLIENTEDGE     '//-- Specifies that the window doesn't have a 3D look — that is, Flat Border Style.
  SetWindowLongA ThWnd, GWL_EXSTYLE, TFlat   '//-- Change attributes of the specified window.
      
  '//-- The SetWindowPos function changes the size, position, and Z order of a child, pop-up, or top-level window.
  '//-- Child, pop-up, and top-level windows are ordered according to their appearance on the screen.
  '//-- The topmost window receives the highest rank and is the first window in the Z order.
  
  With M_ExWindow
    ChangeWindowStyles = SetWindowPos(ThWnd, 0, .X - .bExt, .Y + .aExt, .Width, .Height + Abs(.aExt), SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
  End With
  
  '//-- FYI: SWP_FRAMECHANGED Sends a WM_NCCALCSIZE message to the window
  '//-- FYI: SWP_NOZORDER Retains the current Z order.
  '//-- FYI: SWP_NOACTIVATE Does not activate the window.
  
  
End Function

Private Function InsideArea(cHandle As Long) As Boolean
Dim POS As POINTAPI
        
        GetCursorPos POS

        If (WindowFromPoint(POS.X, POS.Y) <> cHandle) Then
            InsideArea = False
            Else
            InsideArea = True
        End If

End Function


Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Long

  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim hDCSrc As Long
  
   hDCSrc = GetDC(hWndSrc) ' Get device context for entire window.

   ' Create a memory device context for the copy process.
   hDCMemory = CreateCompatibleDC(hDCSrc)
   ' Create a bitmap and place it in the memory DC.
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

  ' Copy the on-screen image into the memory DC.
   BitBlt hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy

  ' Remove the new copy of the  on-screen image.
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' Release the device context resources back to the system.
   Call DeleteDC(hDCMemory)
   Call ReleaseDC(hWndSrc, hDCSrc)

   CaptureWindow = hBmp
   
End Function

Private Sub DrawGradient(lStartcolor As Long, lEndColor As Long, ByRef dRect As RECT, ByVal hdc As Long, Optional bH As Boolean)
    ''Draw a Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    lEndColor = GetLngColor(lEndColor)
    lStartcolor = GetLngColor(lStartcolor)

    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    With dRect
    
        sR = (sR - eR) / IIf(bH, .Right, .Bottom)
        sG = (sG - eG) / IIf(bH, .Right, .Bottom)
        sB = (sB - eB) / IIf(bH, .Right, .Bottom)
    
        
        For ni = 0 To IIf(bH, .Right, .Bottom)
        
            If bH Then
                DrawLine .Left + ni, .TOP, .Left + ni, .Bottom, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
            Else
                DrawLine .Left, .TOP + ni, .Right, .TOP + ni, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
            End If
        
        Next ni
    
    End With
    
End Sub

Private Sub Class_Initialize()
    
    Set Font = New StdFont                 '//-- Font Assigned by the current System Font
         
         mDelayTime = 100                  '//-- 1 Sec
         mKillTime = 500                   '//-- 5 Sec
         mBackStyle = 1                    '//-- Default BackStyle is 'Solid'
         mToolTipStyle = TTStandard        '//-- Default Tooltip Style is 'Standard'
         mShadow = False                   '//-- Default Shadow is OFF
         mIconSize = TTIcon16              '//-- Default Icon is 16x16
         mTextColor = vbButtonText         '//-- Default Color
         mBackColor = vbInfoBackground     '//-- Default Color
         mGCStart = vbInfoBackground       '//-- Default Color
         mGCEnd = vbDesktop                '//-- Default Color
         
End Sub

Private Sub Class_Terminate()
StopAll                           '//-- Stop All SubClassed Windows
pDestroy                          '//-- Destroy Temp DC
DestroyWindow M_ExWindow.hwnd     '//-- Destroy API Window
End Sub




