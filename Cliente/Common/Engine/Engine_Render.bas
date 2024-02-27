Attribute VB_Name = "Engine_Render"
Option Explicit


Private pVB As Direct3DVertexBuffer8
Private Const VERTEX_BUFFER_SIZE As Long = 4000

Private tBox As Box_Vertex


'#################### BUFFFFEEEEEERRRRRRRRRR ####################
Private batch_buffer(0 To 1000 * TL_size)   As Byte
Private batch_count                         As Long
Private batch_texture                       As Integer
Private batch_blend_mode                    As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private PtrVertArray        As Long
Private PtrVertArrayOffset  As Long




'Dibuja y "vacia" el buffer de poligonos
Public Sub batch_render()
    If PtrVertArray Then
        If batch_count Then
            Set_Blend_Mode batch_blend_mode
            Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(batch_texture)
            pVB.Unlock
            Call D3DDevice.SetStreamSource(0, pVB, TL_size)
            D3DDevice.SetIndices pIB, 0
            D3DDevice.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, batch_count * 4, 0, batch_count * 2
            PtrVertArray = 0
            pVB.Lock 0, 0, PtrVertArray, D3DLOCK_DISCARD
            batch_count = 0
        End If
    Else
        pVB.Lock 0, 0, PtrVertArray, 0
        batch_count = 0
    End If
End Sub

'PUSH en el buffer de poligonos
Public Sub batch_add_box(ByRef vertex As Box_Vertex, ByVal texture As Integer, ByVal blend As Byte)
    If (batch_count >= 990 Or batch_texture <> texture Or batch_blend_mode <> blend) Then
        batch_render
        batch_blend_mode = blend
        batch_texture = texture
    End If
    If PtrVertArray Then
        Call CopyMemory(ByVal (PtrVertArray + batch_count * BV_size), vertex, BV_size)
        'D3DVertexBuffer8SetData pVB, batch_count * BV_size, BV_size, 0, vertex
        batch_count = batch_count + 1
    End If
End Sub

Public Sub batch_init(ByRef max As Integer)
    On Error GoTo errh
        ' Create Vertex buffer
        
        Set pVB = D3DDevice.CreateVertexBuffer(VERTEX_BUFFER_SIZE * TL_size, D3DUSAGE_WRITEONLY, FVF, D3DPOOL_DEFAULT)
    
        Call D3DDevice.SetVertexShader(FVF)
        Call D3DDevice.SetStreamSource(0, pVB, TL_size)
        
        pVB.Lock 0, 0, PtrVertArray, 0
        
        
    'IniciarBuffers = True
    Exit Sub
errh:
    MsgBox "No se pudieron iniciar los Vertex Buffer y IndexBuffer"
End Sub
