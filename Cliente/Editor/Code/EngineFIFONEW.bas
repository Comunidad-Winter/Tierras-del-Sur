Attribute VB_Name = "ME_Engine_FIFO"
'                  ____________________________________________
'                 /_____/  http://www.arduz.com.ar/ao/   \_____\
'                //            ____   ____   _    _ _____      \\
'               //       /\   |  __ \|  __ \| |  | |___  /      \\
'              //       /  \  | |__) | |  | | |  | |  / /        \\
'             //       / /\ \ |  _  /| |  | | |  | | / /   II     \\
'            //       / ____ \| | \ \| |__| | |__| |/ /__          \\
'           / \_____ /_/    \_\_|  \_\_____/ \____//_____|_________/ \
'           \________________________________________________________/

Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)

Public FileHandleGraficosInd As Long
'Private Declare Sub CalcularNormal Lib "MZEngine.dll" (ByVal A0 As Single, ByVal A1 As Single, ByVal A2 As Single, ByVal A3 As Single, ByRef C1 As D3DVECTOR, ByRef C2 As D3DVECTOR)

Public pakMapas As clsEnpaquetado
Public pakMapasME As clsEnpaquetado

Public grhversion As Long

Public Sub CargarPakMapas(PathPakMapasME As String, Optional pathPakMapasJuego As String = vbNullString)
Set pakMapasME = New clsEnpaquetado

    If Not FileExist(PathPakMapasME) Then
        pakMapasME.CrearVacio PathPakMapasME, 2000
    Else
        pakMapasME.Cargar PathPakMapasME
    End If
    
    If pathPakMapasJuego <> vbNullString Then
        Set pakMapas = New clsEnpaquetado
        If FileExist(pathPakMapasJuego) Then
            If Not pakMapas.Cargar(pathPakMapasJuego) Then
                Set pakMapas = Nothing
                MsgBox "Error al cargar el enpaquetado de mapas del cliente."
            End If
        Else
            pakMapas.CrearVacio pathPakMapasJuego, 2000
        End If
    End If
End Sub


Public Function SwitchMap(ByVal map As Integer) As Boolean

Dim IH As INFOHEADER

If FileExist(app.Path & "\Datos\tmpmap.cache") Then
    Kill app.Path & "\Datos\tmpmap.cache"
End If

'¿Existe el mapa?
If pakMapasME.Cabezal_GetFileSize(map) Then

    pakMapasME.IH_Get map, IH
    'Cargamos el mapa en concreto
    Cargar_Mapa_ME pakMapasME.Path_res, IH.EmpiezaByte ', IH.lngFileSizeUncompressed
    
    SwitchMap = True
Else
    SwitchMap = False
End If


End Function



':) Ulli's VB Code Formatter V2.24.17 (2010-Oct-27 21:29)  Decl: 24  Code: 778  Total: 802 Lines
':) CommentOnly: 189 (23,6%)  Commented: 6 (0,7%)  Filled: 634 (79,1%)  Empty: 168 (20,9%)  Max Logic Depth: 6
