VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUIText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private pVB As Direct3DVertexBuffer8

Private pTextura As Long

Private textoActual As String

Private posx!, posy!

Private pColor As Long
Private pColor2 As Long

Private LargoCadena As Integer

Private LetrasReales As Long

Private Vertices As Long
Private Triangulos As Long

Private TLV() As Box_Vertex
Private vOut() As Box_Vertex

Private ba() As Byte

Private UseVertexBuffer As Boolean

Event Click(x!, y!)
Event MouseMove(x!, y!)

Private ColorMode As eGuiColorFill

Private NecesitoActualizar As Boolean

Private tBox As Box_Vertex

Public Centrar As Boolean

Private offXCentrado As Single

'Private spliteada() As String
'
'Private lineaMasLarga As Long

Public Property Get Color() As Long
    Color = pColor
End Property

Public Property Let Color(ByVal n As Long)
If pColor <> n Then
    pColor = n
    NecesitoActualizar = True
End If
End Property

Public Property Let Color2(ByVal RHS As Long)
If pColor2 <> RHS Then
    pColor2 = RHS
    NecesitoActualizar = True
End If
End Property

Public Property Get Color2() As Long
Color2 = pColor2
End Property

Public Property Let GradientMode(ByVal RHS As eGuiColorFill)
ColorMode = RHS
NecesitoActualizar = True
End Property

Public Property Get GradientMode() As eGuiColorFill
GradientMode = ColorMode
End Property

Public Property Let text(ByVal txt As String)
    If txt <> textoActual Then
        textoActual = txt
'        spliteada = Split(txt, vbLf)
'
'        lineaMasLarga = 0
'
'        Dim i As Integer
'        For i = 0 To UBound(spliteada)
'            If Len(spliteada(i)) > lineaMasLarga Then lineaMasLarga = Len(spliteada(i))
'        Next i
        
        LargoCadena = Len(txt)
        ba() = StrConv(txt, vbFromUnicode)
        NecesitoActualizar = True
    End If
End Property

Public Property Get text() As String
Attribute text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute text.VB_UserMemId = 0
Attribute text.VB_MemberFlags = "200"
    text = textoActual
End Property

Public Sub SetPos(ByVal x%, ByVal y%)
If posx <> x Or posy <> y Then
    posx = x
    posy = y
    
    Recalcular_Posicion
End If
End Sub

Public Sub GetPos(ByRef x%, ByRef y%)
    x = posx
    y = posy
End Sub

Public Sub Render()
    If LargoCadena = 0 Then Exit Sub

    If NecesitoActualizar Then Recalcular_Arrays
    
    If LetrasReales Then
'        If render_tex_cacheada Then
'            D3DDevice.SetTexture 0, pTex
'            last_texture = 0
'            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
'        Else
            Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(pTextura)
            Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
            If IndexBufferEnabled Then
                If UseVertexBuffer Then
                    D3DDevice.SetIndices pIB, 0
                    D3DDevice.SetStreamSource 0, pVB, TL_size
                    D3DDevice.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, INDEX_BUFFER_SIZE, 0, Vertices
                Else
                    D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLELIST, 0, Vertices, Triangulos, StaticIndexBuffer(0), D3DFMT_INDEX16, vOut(0), TL_size
                End If
            Else
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, Vertices - 2, vOut(0), TL_size
            End If
'        End If
    End If
    
End Sub

Private Sub Recalcular_Arrays(Optional Segunda As Boolean)
    Dim i As Integer
    Dim TempColor As Long, TempColor2 As Long
    Dim ResetColor As Byte
    Dim KeyPharse As Byte
    Dim mu As Integer
    Dim TmpY!, TmpX!
    
    TempColor = pColor
    TempColor2 = pColor2
    
    If LargoCadena Then
    
    Triangulos = 0
    Vertices = 0
    LetrasReales = 0
    
    ReDim TLV(LargoCadena)
    
        For i = 0 To LargoCadena - 1
            If ba(i) = 255 Then
                KeyPharse = (Not KeyPharse)
                If KeyPharse Then
                    TempColor = pColor2
                    TempColor2 = pColor
                Else
                    ResetColor = 1
                End If
            ElseIf ba(i) = vbKeyReturn Then
                TmpY = TmpY + Font_Default.CharHeight
                TmpX = 0
            ElseIf ba(i) = 10 Then ' Nueva linea
            ElseIf ba(i) = vbKeySpace Then
                TmpX = TmpX + Font_Default.HeaderInfo.CharWidth(vbKeySpace)
            Else
                DXCopyMemory TLV(LetrasReales), Font_Default.HeaderInfo.CharVA(ba(i)).vertex, BV_size

                TLV(LetrasReales).x0 = TmpX
                TLV(LetrasReales).y0 = TmpY
                
                TLV(LetrasReales).x1 = TLV(LetrasReales).x1 + TmpX '* scalea
                TLV(LetrasReales).y1 = TLV(LetrasReales).y0

                TLV(LetrasReales).x2 = TmpX
                TLV(LetrasReales).y2 = TLV(LetrasReales).y2 + TmpY '* scalea

                TLV(LetrasReales).x3 = TLV(LetrasReales).x1
                TLV(LetrasReales).y3 = TLV(LetrasReales).y2
                
                If ColorMode = dSolid Then
                    TLV(LetrasReales).color0 = TempColor
                    TLV(LetrasReales).Color1 = TempColor
                    TLV(LetrasReales).Color2 = TempColor
                    TLV(LetrasReales).color3 = TempColor
                ElseIf ColorMode = dVertical Then
                    TLV(LetrasReales).color0 = TempColor
                    TLV(LetrasReales).Color1 = TempColor
                    TLV(LetrasReales).Color2 = TempColor2
                    TLV(LetrasReales).color3 = TempColor2
                ElseIf ColorMode = dHorizontal Then
                    mu = (LetrasReales / LargoCadena) * 127
                    Call Lerp44(TLV(LetrasReales).color0, TempColor, TempColor2, mu)
                    TLV(LetrasReales).Color1 = TLV(LetrasReales).color0
                    TLV(LetrasReales).Color2 = TLV(LetrasReales).color0
                    TLV(LetrasReales).color3 = TLV(LetrasReales).color0
                End If
                
                Vertices = Vertices + 4
                Triangulos = Triangulos + 2
                LetrasReales = LetrasReales + 1
                
                TmpX = TmpX + Font_Default.HeaderInfo.CharWidth(ba(i)) '* scalea
            End If
            
            If ResetColor Then
                ResetColor = 0
                TempColor = pColor
                TempColor2 = pColor2
            End If
        Next i

        If LetrasReales Then
            With tBox
                .color0 = mzWhite
                .Color1 = mzWhite
                .Color2 = mzWhite
                .color3 = mzWhite
                
                
                .rhw0 = 1
                .rhw1 = 1
                .rhw2 = 1
                .rhw3 = 1
                
                .x0 = 0
                .y0 = TLV(LetrasReales - 1).y2
                
                .x1 = 0
                .y1 = 0
                
                .x2 = TLV(LetrasReales - 1).x1
                .y2 = .y0
                
                .x3 = .x2
                .y3 = 0
                
                .tu0 = 0
                .tv0 = 1
                .tu1 = 0
                .tv1 = 0
                .tu2 = 1
                .tv2 = 1
                .tu3 = 1
                .tv3 = 0
                
            End With
        End If

        offXCentrado = Engine_GetTextWidth(text) * -0.5
        
        NecesitoActualizar = False
        Recalcular_Posicion True

    End If
End Sub

Private Sub Recalcular_Posicion(Optional ByVal ReCalculado As Boolean = False)
    If LetrasReales Then
        CopiarVertices
        If UseVertexBuffer Then SetVertexBufferData
    End If
End Sub

Private Sub CopiarVertices()
    Dim i As Long
    ReDim vOut(LetrasReales)
    DXCopyMemory vOut(0), TLV(0), BV_size * LetrasReales
    
    If Centrar Then
        For i = 0 To LetrasReales - 1
            vOut(i).x0 = vOut(i).x0 + posx + offXCentrado
            vOut(i).x1 = vOut(i).x1 + posx + offXCentrado
            vOut(i).x2 = vOut(i).x2 + posx + offXCentrado
            vOut(i).x3 = vOut(i).x3 + posx + offXCentrado
            
            vOut(i).y0 = vOut(i).y0 + posy
            vOut(i).y1 = vOut(i).y1 + posy
            vOut(i).y2 = vOut(i).y2 + posy
            vOut(i).y3 = vOut(i).y3 + posy
        Next i
    Else
        For i = 0 To LetrasReales - 1
            vOut(i).x0 = vOut(i).x0 + posx
            vOut(i).x1 = vOut(i).x1 + posx
            vOut(i).x2 = vOut(i).x2 + posx
            vOut(i).x3 = vOut(i).x3 + posx
            
            vOut(i).y0 = vOut(i).y0 + posy
            vOut(i).y1 = vOut(i).y1 + posy
            vOut(i).y2 = vOut(i).y2 + posy
            vOut(i).y3 = vOut(i).y3 + posy
        Next i
    End If
End Sub

Private Sub SetVertexBufferData()
    On Error GoTo errh
        Dim Ptr As Long
        If Not pVB Is Nothing Then
            D3DDevice.SetStreamSource 0, Nothing, 0
            Set pVB = Nothing
        End If
        
        Set pVB = D3DDevice.CreateVertexBuffer(LetrasReales * BV_size, D3DUSAGE_WRITEONLY, FVF, D3DPOOL_MANAGED)
        pVB.Lock 0, 0, Ptr, 0
        DXCopyMemory ByVal Ptr, vOut(0), BV_size * LetrasReales
        pVB.Unlock
        UseVertexBuffer = True
    Exit Sub
errh:
    UseVertexBuffer = False
End Sub


Private Sub Class_Initialize()
'Marce On error resume next
    pTextura = TexturaTexto2
    pColor = &HFFFFFFFF
    pColor2 = &HFFAACCAA

    ColorMode = 0

    
    UseVertexBuffer = False
    ' AGUS: 2015-03-22 A ver si los vertex buffer estan chingando los labels. comento 3 lineas \/
'    Set pVB = D3DDevice.CreateVertexBuffer(BV_size, D3DUSAGE_WRITEONLY, FVF, D3DPOOL_MANAGED)
'    If Not pVB Is Nothing Then UseVertexBuffer = True
'    Set pVB = Nothing
End Sub

Private Sub Class_Terminate()
If Not pVB Is Nothing Then
    Set pVB = Nothing
    If Not D3DDevice Is Nothing Then
        D3DDevice.SetStreamSource 0, Nothing, 0
    End If
End If
End Sub

Public Property Get textura() As Integer
textura = pTextura
End Property

Public Property Let textura(ByVal vNewValue As Integer)
pTextura = vNewValue
End Property

Public Property Get Alineacion() As Variant

End Property

Public Property Let Alineacion(ByVal vNewValue As Variant)

End Property

Public Sub GetSize(ByRef w As Integer, ByRef h As Integer)
w = tBox.x2 - tBox.x0
h = tBox.y2 - tBox.y1
End Sub
