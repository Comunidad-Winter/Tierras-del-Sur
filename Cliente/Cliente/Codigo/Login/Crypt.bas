Attribute VB_Name = "Crypt"
Private x1a0(9) As Long
Private cle(17) As Long
Private x1a2 As Long

Private inter As Long, res As Long, ax As Long, BX As Long
Private CX As Long, DX As Long, si As Long, Tmp As Long
Private i As Long, c As Byte

Option Explicit

Private Sub Assemble()

    x1a0(0) = ((cle(1) * 256) + cle(2)) Mod 65536
    Code
    inter = res
    
    x1a0(1) = x1a0(0) Xor ((cle(3) * 256) + cle(4))
    Code
    inter = inter Xor res
    
    
    x1a0(2) = x1a0(1) Xor ((cle(5) * 256) + cle(6))
    Code
    inter = inter Xor res
    
    x1a0(3) = x1a0(2) Xor ((cle(7) * 256) + cle(8))
    Code
    inter = inter Xor res
    
    x1a0(4) = x1a0(3) Xor ((cle(9) * 256) + cle(10))
    Code
    inter = inter Xor res
    
    x1a0(5) = x1a0(4) Xor ((cle(11) * 256) + cle(12))
    Code
    inter = inter Xor res
    
    x1a0(6) = x1a0(5) Xor ((cle(13) * 256) + cle(14))
    Code
    inter = inter Xor res
    
    x1a0(7) = x1a0(6) Xor ((cle(15) * 256) + cle(16))
    Code
    inter = inter Xor res
    
    i = 0

End Sub

Private Sub Code()
    DX = (x1a2 + i) Mod 65536
    ax = x1a0(i)
    CX = &H15A
    BX = &H4E35
    
    Tmp = ax
    ax = si
    si = Tmp
    
    Tmp = ax
    ax = DX
    DX = Tmp
    
    If (ax <> 0) Then
        ax = (ax * BX) Mod 65536
    End If
    
    Tmp = ax
    ax = CX
    CX = Tmp
    
    If (ax <> 0) Then
        ax = (ax * si) Mod 65536
        CX = (ax + CX) Mod 65536
    End If
    
    Tmp = ax
    ax = si
    si = Tmp
    ax = (ax * BX) Mod 65536
    DX = (CX + DX) Mod 65536
    
    ax = ax + 1
    
    x1a2 = DX
    x1a0(i) = ax
    
    res = ax Xor DX
    i = i + 1

End Sub


Public Function crypt(cadena As String, clave As String)
    Dim encriptado As String
    Dim fois As Integer
    Dim champ1 As String
    Dim lngchamp1 As Integer
    Dim cfc As Integer
    Dim cfd As Integer
    Dim compte As Byte
    Dim e As Byte
    Dim d As Integer
    
    encriptado = ""
    si = 0
    x1a2 = 0
    i = 0
    
    For fois = 1 To 16
        cle(fois) = 0
    Next fois
    
    champ1 = clave
    lngchamp1 = Len(champ1)
    
    For fois = 1 To lngchamp1
        cle(fois) = Asc(mid(champ1, fois, 1))
    Next fois
    
    champ1 = cadena
    lngchamp1 = Len(champ1)
    
    For fois = 1 To lngchamp1
        c = Asc(mid(champ1, fois, 1))
        
        Assemble
        
        If inter > 65535 Then
        inter = inter - 65536
        End If
        
        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256
    
        For compte = 1 To 16
        
            cle(compte) = cle(compte) Xor c
        
        Next compte
    
        c = c Xor (cfc Xor cfd)
        
        d = (((c / 16) * 16) - (c Mod 16)) / 16
        e = c Mod 16
        
        encriptado = encriptado + Chr$(&H61 + d) ' d+&h61 give one letter range from a to p for the 4 high bits of c
        encriptado = encriptado + Chr$(&H61 + e) ' e+&h61 give one letter range from a to p for the 4 low bits of c
    
    
    Next fois
    

    Dim loopCaracter As Integer
    
    'Ponemos una capa adicional de enmasqueramiento
    For loopCaracter = 1 To Len(encriptado)
         crypt = crypt & Chr$(Asc(mid$(encriptado, loopCaracter, 1)) Xor 7)
    Next
    
    'retorno
    crypt = UCase$(crypt)
    'crypt = encriptado
End Function

Public Function decrypt(ByVal cadena As String, clave As String)
    Dim desencriptado As String
    Dim loopCaracter As Integer
    Dim fois As Integer
    Dim champ1 As String
    Dim lngchamp1 As Integer
    Dim cfc As Integer
    Dim cfd As Integer
    Dim compte As Byte
    Dim c As Byte
    Dim d As Byte
    Dim e As Byte
    
    'Quitamos la capa adicional de enmasqueramiento
    cadena = LCase$(cadena)
    
    For loopCaracter = 1 To Len(cadena)
        desencriptado = desencriptado & Chr$(Asc(mid$(cadena, loopCaracter, 1)) Xor 7)
    Next
    
    cadena = desencriptado
    
    desencriptado = ""
    si = 0
    x1a2 = 0
    i = 0

    For fois = 1 To 16
        cle(fois) = 0
    Next fois

    champ1 = clave
    lngchamp1 = Len(champ1)

    For fois = 1 To lngchamp1
        cle(fois) = Asc(mid(champ1, fois, 1))
    Next fois

    champ1 = cadena
    lngchamp1 = Len(champ1)

    For fois = 1 To lngchamp1
    
        d = Asc(mid(champ1, fois, 1))
        If (d - &H61) >= 0 Then
            d = d - &H61  ' to transform the letter to the 4 high bits of c
            If (d >= 0) And (d <= 15) Then
                d = d * 16
            End If
        End If
        
        If (fois <> lngchamp1) Then
            fois = fois + 1
        End If
    
        e = Asc(mid(champ1, fois, 1))
        
        If (e - &H61) >= 0 Then
            e = e - &H61 ' to transform the letter to the 4 low bits of c
            If (e >= 0) And (e <= 15) Then
            c = d + e
            End If
        End If
    
        Assemble
    
        If inter > 65535 Then
            inter = inter - 65536
        End If
    
        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256
    
        c = c Xor (cfc Xor cfd)
    
        For compte = 1 To 16
            cle(compte) = cle(compte) Xor c
        Next compte
    
        desencriptado = desencriptado + Chr$(c)
    
    Next fois
        
        decrypt = desencriptado
End Function


