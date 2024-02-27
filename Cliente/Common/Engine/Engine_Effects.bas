Attribute VB_Name = "Engine_PS"
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

Public PS_Glow As Long
'Public Const Glowsrc   As String = "ps.1.1 tex t0 tex t1 add_sat r0, t0, t1"
'Public Const Glowsrc   As String = "vs.1.1 dcl_position v0 dcl_texcoord v3 mov oPos, v0 add oT0, v3, c0 add oT1, v3, c1 add oT2, v3, c2 add oT3, v3, c3"
'Public Const Gussian_Blur   As String = "ps.1.4 def c0, 0.2f, 0.2f, 0.2f, 1.0f texld r0, t0 texld r1, t1 texld r2, t2 texld r3, t3 texld r4, t4 add r0, r0, r1 add r2, r2, r3 add r0, r0, r2 add r0, r0, r4 mul r0, r0, c0"
'Public Const BrightPasssrc   As String = "ps.1.4 def c0,0.561797752,0.561797752,0.561797752,1 def c1,0.78125,0.78125,0.78125,1 def c2,1,1,1,1 def c3,0.1,0.1,0.1,1 texld r0,t0 mul_x4 r0,r0,c0 mul_x2 r1,r0,c1 add r1,r1,c2 mul r0,r0,r1 mov_x4 r2,c2 add r2,r2,c2 sub r0,r0,r2 mul_sat r0,r0,c3"
Public Const Glowsrc   As String = "ps.1.1 " & _
                                    "def c0, 0.4, 0.4, 0.4, 0.4 " & _
                                    "tex t0 " & _
                                    "mul r0, t0, c0"
                                    
Public soporta_pixelShader As Boolean


Public Function mkVec3f(x As Single, y As Single, z As Single) As D3DVECTOR
  With mkVec3f
    .x = x
    .y = y
    .z = z
  End With
End Function


Public Function shCompile(fName As String) As Long

  On Error Resume Next
  shCompile = 0
  
  Static shArray() As Long
  Static shLength As Long
  Static shCode As D3DXBuffer

  Set shCode = D3DX.AssembleShaderFromFile(fName, 0, vbNullString, Nothing)
  shLength = shCode.GetBufferSize() / 4
  
  If Not Err.Number = 0 Then
    Err.Clear
    Set shCode = Nothing
    LogError "No se pudo compilar el PixelShader"
  Else
  
    ReDim shArray(shLength - 1) As Long
    D3DX.BufferGetData shCode, 0, 4, shLength, shArray(0)
    
    shCompile = D3DDevice.CreatePixelShader(shArray(0))
    soporta_pixelShader = True
    If Not Err.Number = 0 Or shCompile = 0 Then
      Err.Clear
      Set shCode = Nothing
      shCompile = 0
      soporta_pixelShader = False
      LogError "No se pudo crear el PixelShader, si se pudo compilar"
    End If
  
  End If

End Function

Public Function shCompileT(t As String) As Long

  On Error Resume Next
  shCompileT = 0
  
  Static shArray() As Long
  Static shLength As Long
  Static shCode As D3DXBuffer

  Set shCode = D3DX.AssembleShader(t, 0, Nothing)
  shLength = shCode.GetBufferSize() / 4
  
  If Not Err.Number = 0 Then
    Err.Clear
    Set shCode = Nothing
    LogError "No se pudo compilar el PixelShader"
    toggle_lights_powa False
  Else
  
    ReDim shArray(shLength - 1) As Long
    D3DX.BufferGetData shCode, 0, 4, shLength, shArray(0)
    
    shCompileT = D3DDevice.CreatePixelShader(shArray(0))
    soporta_pixelShader = True
    
    If Not Err.Number = 0 Or shCompileT = 0 Then
      Err.Clear
      Set shCode = Nothing
      shCompileT = 0
      soporta_pixelShader = False
      LogError "No se pudo crear el PixelShader, si se pudo compilar"
      toggle_lights_powa False
    End If
  
  End If

End Function

'Public Function vsCompileT(t As String) As Long
'
'  'On Error Resume Next
'  shCompileT = 0
'
'  Static shArray() As Long
'  Static shLength As Long
'  Static shCode As D3DXBuffer
'
'  Set shCode = D3DX.AssembleShader(t, 0, Nothing)
'  shLength = shCode.GetBufferSize() / 4
'
'  If Not Err.Number = 0 Then
'    Err.Clear
'    Set shCode = Nothing
'    MsgBox "Could not assemble pixel shader.", vbCritical Or vbOKOnly, "Error"
'  Else
'
'    ReDim shArray(shLength - 1) As Long
'    D3DX.BufferGetData shCode, 0, 4, shLength, shArray(0)
'
'    shCompileT = D3DDevice.CreateVertexShader(shArray(0))
'
'    If Not Err.Number = 0 Or shCompileT = 0 Then
'      Err.Clear
'      Set shCode = Nothing
'      shCompileT = 0
'      MsgBox "Pixel shader was sucessfully assembled, but failed to create." & vbCrLf, vbCritical Or vbOKOnly, "Error"
'    End If
'
'  End If
'
'End Function
