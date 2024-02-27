Attribute VB_Name = "MCsVlgTracing"
' MCsVlgTracing - PLEASE DO NOT DELETE THIS LINE
'                 AND DO NOT ALTER THE METHOD NAMES IN ANY WAY

'CSEH: Skip

Option Explicit

' Specify the position where the tracing instructions are inserted
Public Enum ProcPosition
    ProcEnter
    ProcExit
    ProcInside
End Enum

' Variable which contains a reference to AxTools Visual Logger
Private m_oVLogger As VisualLogger

' Parameters for Visual Logger main call
Private m_Color    As ItemColor
Private m_Bold     As Boolean
Private m_Indent   As Long

Public Const E_ERR_CS_TRACING_INIT = vbObjectError + 1263

Private Const S_ERR_CS_TRACING_INIT = "Could not initiate tracing"

' Member      : AxCsInitiateTrace
' Description : Instantiate the Visual Logger reference
Public Sub AxCsInitiateTrace()

    Static bErrReported As Boolean

    On Error GoTo hErr

    On Error Resume Next
    ' Instantiate a new Visual Logger client
    Set m_oVLogger = New VisualLogger
    On Error GoTo hErr

    If m_oVLogger Is Nothing Then
        If Not bErrReported Then
            bErrReported = True
            Err.Raise E_ERR_CS_TRACING_INIT + 1, "AxCsInitiateTrace", _
                "Could not create Visual Logger object. " & _
                "Please check the AxTools Visual Logger COM server validity."
        End If
        Exit Sub
    End If

    ' Instantiate a new Visual Logger client
    Set m_oVLogger = New VisualLogger
    
    ' Register the new client to the server
    m_oVLogger.Register "TDS_1"
    m_oVLogger.IndentSize = 1
    
    ' Initialize fields
    m_Indent = 0
    m_Color = -1
    m_Bold = False
    
    Exit Sub

hErr:
    Err.Raise E_ERR_CS_TRACING_INIT, "AxCsInitiateTrace", S_ERR_CS_TRACING_INIT

End Sub

' Member      : AxCsTerminateTrace
' Description : Used for cleaning purposes
Public Sub AxCsTerminateTrace()
    
    Set m_oVLogger = Nothing
    
End Sub

' Member      : AxCsTrace
' Description : Send information to the Visual Logger window regarding the current method being processed
' Parameters  : ProjectName   - Name of the project which contains the method
'               ComponentName - Name of the component which contains the method
'               MemberName    - Name of the method being processed
'               TracePosition - Indicates the position within method body, either at start, at exit
'                               or inside it, where the AxCsTraceWatch method is called
' Notes       : For inside-member calls of AxCsTrace method, you can use the ProjectName
'               parameter to send tracing information.
Public Sub AxCsTrace(ByVal ProjectName As String, _
                     Optional ByVal ComponentName As String = "", _
                     Optional ByVal MemberName As String = "", _
                     Optional ByVal TracePosition As ProcPosition = ProcInside)

  Debug.Print MemberName
    
End Sub

' Member      : VLogger
' Description : Get the m_oVLogger object
Public Property Get VLogger() As VisualLogger
    
    If m_oVLogger Is Nothing Then AxCsInitiateTrace

    Set VLogger = m_oVLogger

End Property

' Member      : AxCsDumpParamValue
' Description : Returns a formatted information about each method parameter
' Parameters  : sParamName    - Parameter name
'               vParamValue   - Parameter value
Public Function AxCsDumpParamValue(sParamName As String, vParamValue As Variant) As String

    Dim sRet$
    
    If IsObject(vParamValue) Then
    
        sRet = "Object"
        
    ElseIf IsArray(vParamValue) Then
    
        sRet = "Array"
        
    Else
    
        On Error Resume Next
        sRet = CStr(vParamValue)
        
        If Err.Number <> 0 Then
            sRet = "Not determined"
        Else
            If Len(sRet) > 13 Then sRet = left$(sRet, 10) & "..."
        End If
        
    End If
    
    AxCsDumpParamValue = sParamName & ": [" & sRet & "]"

End Function



