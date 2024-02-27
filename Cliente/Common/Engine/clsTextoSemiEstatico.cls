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
Implements clsGUIObject

Option Explicit

Private pVB As Direct3DVertexBuffer8

Private pTextura As Long

Private TextoActual As String

Private PosX!, PosY!

Private pColor As Long
Private pColor2 As Long

Private LargoCadena As Integer

Private LetrasReales As Long

Private Vertices As Long
Private Triangulos As Long

Private TLV() As Box_Vertex
Private vOut() As Box_Vertex

Private BA() As Byte

Private UseVertexBuffer As Boolean

Event Click(x!, Y!)
Event MouseMove(x!, Y!)

Private ColorMode As eGuiColorFill

Private NecesitoActualizar As Boolean

Private pTex As Direct3DTexture8
Private render_tex_cacheada As Boolean

Private tBox As Box_Vertex

Public Property Get clsGUIObject_Color() As Long
    clsGUIObject_Color = pColor
End Property

Public Property Let clsGUIObject_Color(ByVal n As Long)
    pColor = n
    NecesitoActualizar = True
End Property

Private Property Let clsGUIObject_Color2(ByVal RHS As Long)
pColor2 = RHS
NecesitoActualizar = True
End Property

Private Property Get clsGUIObject_Color2() As Long
clsGUIObject_Color2 = pColor2
End Property

Private Property Let clsGUIObject_GradientMode(ByVal RHS As eGuiColorFill)
ColorMode = RHS
NecesitoActualizar = True
End Property

Private Property Get clsGUIObject_GradientMode() As eGuiColorFill
clsGUIObject_GradientMode = ColorMode
End Property

Public Property Let clsGUIObject_Text(ByVal txt As String)
    If txt <> TextoActual Then
        TextoActual = txt
        LargoCadena = Len(txt)
        BA() = StrConv(txt, vbFromUnicode)
        NecesitoActualizar = True
    End If
End Property

Public Property Get clsGUIObject_Text() As String
    clsGUIObject_Text = TextoActual
End Property

Public Sub clsGUIObject_SetPos(ByVal x!, ByVal Y!)
    PosX = Round(x)
    PosY = Round(Y)
    
    Recalcular_Posicion
End Sub

Public Sub clsGUIObject_GetPos(ByRef x!, ByRef Y!)
    x = PosX
    Y = PosY
End Sub

Public Sub clsGUIObject_Render()
    If LargoCadena > 0 And NecesitoActualizar Then Recalcular_Arrays
    If LetrasReales Then
        If render_tex_cacheada Then
            D3DDevice.SetTexture 0, pTex
            last_texture = 0
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
        Else
            Call GetTexture(pTextura)
            If lColorMod <> D3DTOP_MODULATE Then Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
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
            If lColorMod <> D3DTOP_MODULATE Then Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)
        End If
    End If
    
End Sub

Private Sub Recalcular_Arrays()
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
            If BA(i) = 255 Then
                KeyPharse = (Not KeyPharse)
                If KeyPharse Then
                    TempColor = pColor2
                    TempColor2 = pColor
                Else
                    ResetColor = 1
                End If
            ElseIf BA(i) = 13 Then
                TmpY = TmpY + Font_Default.CharHeight
                TmpX = 0
            ElseIf BA(i) = 10 Then
            ElseIf BA(i) = 32 Then
                TmpX = TmpX + Font_Default.HeaderInfo.CharWidth(32)
            Else
                CopyMemory TLV(LetrasReales), Font_Default.HeaderInfo.CharVA(BA(i)).vertex, BV_size

                TLV(LetrasReales).x0 = TmpX
                TLV(LetrasReales).y0 = TmpY
                
                TLV(LetrasReales).x1 = TLV(LetrasReales).x1 + TmpX '* scalea
                TLV(LetrasReales).y1 = TLV(LetrasReales).y0

                TLV(LetrasReales).X2 = TmpX
                TLV(LetrasReales).y2 = TLV(LetrasReales).y2 + TmpY '* scalea

                TLV(LetrasReales).x3 = TLV(LetrasReales).x1
                TLV(LetrasReales).y3 = TLV(LetrasReales).y2
                
                If ColorMode = dSolid Then
                    TLV(LetrasReales).color0 = TempColor
                    TLV(LetrasReales).color1 = TempColor
                    TLV(LetrasReales).Color2 = TempColor
                    TLV(LetrasReales).color3 = TempColor
                ElseIf ColorMode = dVertical Then
                    TLV(LetrasReales).color0 = TempColor
                    TLV(LetrasReales).color1 = TempColor
                    TLV(LetrasReales).Color2 = TempColor2
                    TLV(LetrasReales).color3 = TempColor2
                ElseIf ColorMode = dHorizontal Then
                    mu = (LetrasReales / LargoCadena) * 127
                    Call Lerp44(TLV(LetrasReales).color0, TempColor, TempColor2, mu)
                    TLV(LetrasReales).color1 = TLV(LetrasReales).color0
                    TLV(LetrasReales).Color2 = TLV(LetrasReales).color0
                    TLV(LetrasReales).color3 = TLV(LetrasReales).color0
                End If
                
                Vertices = Vertices + 4
                Triangulos = Triangulos + 2
                LetrasReales = LetrasReales + 1
                
                TmpX = TmpX + Font_Default.HeaderInfo.CharWidth(BA(i)) '* scalea
            End If
            
            If ResetColor Then
                ResetColor = 0
                TempColor = pColor
                TempColor2 = pColor2
            End If
        Next i

        
        With tBox
            .color0 = mzWhite
            .color1 = mzWhite
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
            
            .X2 = TLV(LetrasReales - 1).x1
            .y2 = .y0
            
            .x3 = .X2
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
        
        NecesitoActualizar = False
        Recalcular_Posicion True
    End If
End Sub

Private Sub Recalcular_Posicion_Textura()
    'Renderizo en la textura a cara de perro, y despues hago un render de esa textura
    On Error GoTo errh:
    
    Dim UseVertexBufferTMP As Boolean
    UseVertexBufferTMP = UseVertexBufferç
    UseVertexBuffer = False
    
    ReDim vOut(LetrasReales)
    CopyMemory vOut(0), TLV(0), BV_size * LetrasReales

    render_tex_cacheada = False

    StartRenderToTexture pTex
    
    clsGUIObject_Render
    EndRenderToTexture
    
    render_tex_cacheada = True
    
    UseVertexBuffer = UseVertexBufferTMP
Exit Sub
errh:
UseVertexBuffer = UseVertexBufferTMP
Recalcular_Posicion

End Sub

Private Function SePuedeCacehar() As Boolean
    SePuedeCacehar = tBox.x3 < 512 And tBox.y0 < 128 And Optimizar_Textos And PuedeRenderToTexture And (Engine.Engine_Escene_Abierta = False)
End Function

Private Sub Recalcular_Posicion(Optional ByVal ReCalculado As Boolean = False)
    If LetrasReales Then
        If SePuedeCacehar Then
            Recalcular_Posicion_Textura
        Else
            CopiarVertices
            If UseVertexBuffer Then SetVertexBufferData
            render_tex_cacheada = False
        End If
    End If
End Sub

Private Sub CopiarVertices()
    Dim i As Long
    ReDim vOut(LetrasReales)
    CopyMemory vOut(0), TLV(0), BV_size * LetrasReales
    
    For i = 0 To LetrasReales - 1
        vOut(i).x0 = vOut(i).x0 + PosX
        vOut(i).x1 = vOut(i).x1 + PosX
        vOut(i).X2 = vOut(i).X2 + PosX
        vOut(i).x3 = vOut(i).x3 + PosX
        
        vOut(i).y0 = vOut(i).y0 + PosY
        vOut(i).y1 = vOut(i).y1 + PosY
        vOut(i).y2 = vOut(i).y2 + PosY
        vOut(i).y3 = vOut(i).y3 + PosY
    Next i
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
        CopyMemory ByVal Ptr, vOut(0), BV_size * LetrasReales
        pVB.Unlock
        UseVertexBuffer = True
    Exit Sub
errh:
    UseVertexBuffer = False
End Sub


Private Sub Class_Initialize()
On Error Resume Next
    pTextura = 9733 '9718
    pColor = &HFFFFFFFF
    pColor2 = &HFFAACCAA

    ColorMode = 0

    UseVertexBuffer = False
    Set pVB = D3DDevice.CreateVertexBuffer(BV_size, D3DUSAGE_WRITEONLY, FVF, D3DPOOL_MANAGED)
    If Not pVB Is Nothing Then UseVertexBuffer = True
    Set pVB = Nothing
End Sub

Private Sub Class_Terminate()
If Not pVB Is Nothing Then
    Set pVB = Nothing
    If Not D3DDevice Is Nothing Then
        D3DDevice.SetStreamSource 0, Nothing, 0
    End If
End If
End Sub


