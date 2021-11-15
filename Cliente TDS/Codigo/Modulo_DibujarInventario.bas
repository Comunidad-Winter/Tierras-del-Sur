Attribute VB_Name = "DibujarInventario"
'Argentum Online 0.11.2
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

'[CODE]:MatuX
'
'  Casi todo recodeado menos los calculos
'
'[END]'

Public Const XCantItems As Byte = 5

Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public mx As Integer
Public my As Integer

Private AuxSurface   As DirectDrawSurface7
Private BoxSurface   As DirectDrawSurface7
Private SelSurface   As DirectDrawSurface7
Private bStaticInit  As Boolean   'Se inicializaron las Statics?
Private r1           As RECT, r2 As RECT, auxr As RECT
Private rBox         As RECT  'Pos del cuadradito rojo
Private rBoxFrame(2) As RECT
Private iFrameMod    As Integer

Function ClicEnItemElegido(ByVal X As Integer, ByVal Y As Integer, picInv As PictureBox) As Boolean
Dim kMx As Integer, kMy As Integer
If X > 0 And Y > 0 And X < picInv.ScaleWidth And Y < picInv.ScaleHeight Then
    'bInvMod = True
    kMx = X \ 32 + 1
    kMy = Y \ 32 + 1
    If ItemElegido = FLAGORO Then
        ClicEnItemElegido = False
    Else
'        ClicEnItemElegido = (UserInventory(ItemElegido).OBJIndex > 0) And (ItemElegido = (kMx + (kMy - 1) * 5) + OffsetDelInv)
        ClicEnItemElegido = (UserInventory(ItemElegido).OBJIndex > 0) And (kMx = mx) And (kMy = my)
    End If
End If
End Function

Sub ItemClick(X As Integer, Y As Integer, picInv As PictureBox)
Dim lPreItem As Long
bInvMod = False
If X > 0 And Y > 0 And X < picInv.ScaleWidth And Y < picInv.ScaleHeight Then
    mx = X \ 32 + 1
    my = Y \ 32 + 1
    lPreItem = (mx + (my - 1) * 5) + OffsetDelInv
    If lPreItem <= MAX_INVENTORY_SLOTS Then
        If UserInventory(lPreItem).GrhIndex > 0 Then
            ItemElegido = lPreItem
            bInvMod = True
        End If
    End If
End If
End Sub

Public Sub DibujarInvBox()
    On Error Resume Next
    If bStaticInit And ItemElegido <> 0 Then
        Call BoxSurface.BltColorFill(auxr, vbBlack)
        Call BoxSurface.BltFast(0, 0, SelSurface, auxr, DDBLTFAST_SRCCOLORKEY)
        With Grh(1)
            .FrameCounter = 2
            Call BoxSurface.BltFast(0, 0, SurfaceDB.GetBMP(GrhData(GrhData(.GrhIndex).Frames(.FrameCounter)).FileNum), rBoxFrame(.FrameCounter - 1), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        End With
        Call BoxSurface.BltToDC(frmMain.picInv.Hdc, auxr, rBox)
        Call frmMain.picInv.Refresh
    End If
End Sub

Sub DibujarInv()
Dim iX As Integer
If Not bStaticInit Then _
    Call InitMem
r1.top = 0: r1.left = 0: r1.right = 32: r1.bottom = 32
r2.top = 0: r2.left = 0: r2.right = 32: r2.bottom = 32
frmMain.picInv.Cls
For iX = OffsetDelInv + 1 To UBound(UserInventory)
    If UserInventory(iX).GrhIndex > 0 Then
        AuxSurface.BltColorFill auxr, vbBlack
        AuxSurface.BltFast 0, 0, SurfaceDB.GetBMP(GrhData(UserInventory(iX).GrhIndex).FileNum, 0), auxr, DDBLTFAST_NOCOLORKEY
        AuxSurface.DrawText 0, 0, UserInventory(iX).Amount, False
        If UserInventory(iX).Equipped Then
            AuxSurface.SetForeColor vbYellow
            AuxSurface.DrawText 20, 20, "+", False
            AuxSurface.SetForeColor vbWhite
        End If
        If ItemElegido = iX Then
            With r2: .left = (mx - 1) * 32: .right = r2.left + 32: .top = (my - 1) * 32: .bottom = r2.top + 32: End With
            Call AuxSurface.BltFast(0, 0, SurfaceDB.GetBMP(GrhData(GrhData(Grh(1).GrhIndex).Frames(2)).FileNum), rBoxFrame(2), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        End If
        AuxSurface.BltToDC frmMain.picInv.Hdc, auxr, r2
    End If
    r2.left = r2.left + 32
    r2.right = r2.right + 32
    r1.left = r1.left + 32
    r1.right = r1.right + 32
    If r2.left >= 160 Then
        r2.left = 0
        r1.left = 0
        r1.right = 32
        r2.right = 32
        r2.top = r2.top + 32
        r1.top = r1.top + 32
        r2.bottom = r2.bottom + 32
        r1.bottom = r1.bottom + 32
    End If
Next iX
bInvMod = False
If ItemElegido = 0 Then _
    Call ItemClick(2, 2, frmMain.picInv)
End Sub

Private Sub InitMem()
    Dim ddck        As DDCOLORKEY
    Dim SurfaceDesc As DDSURFACEDESC2
    
    'Back Buffer Surface
    r1.right = 32: r1.bottom = 32
    r2.right = 32: r2.bottom = 32
    
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = r1.bottom
        .lWidth = r1.right
    End With
    ' Create surface
    Set AuxSurface = DirectDraw.CreateSurface(SurfaceDesc)
    Set BoxSurface = DirectDraw.CreateSurface(SurfaceDesc)
    Set SelSurface = DirectDraw.CreateSurface(SurfaceDesc)
    'Set color key
    AuxSurface.SetColorKey DDCKEY_SRCBLT, ddck
    BoxSurface.SetColorKey DDCKEY_SRCBLT, ddck
    SelSurface.SetColorKey DDCKEY_SRCBLT, ddck
    auxr.right = 32: auxr.bottom = 32
    AuxSurface.SetFontTransparency True
    AuxSurface.SetFont frmMain.Font
    SelSurface.SetFontTransparency True
    SelSurface.SetFont frmMain.Font
    'RedBox Frame Position List
    With rBoxFrame(0): .left = 0:  .top = 0: .right = 32: .bottom = 32: End With
    With rBoxFrame(1): .left = 32: .top = 0: .right = 64: .bottom = 32: End With
    With rBoxFrame(2): .left = 64: .top = 0: .right = 96: .bottom = 32: End With
    iFrameMod = 1
    bStaticInit = True
End Sub
'********************Misery_Ezequiel 28/05/05********************'
