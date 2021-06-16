Attribute VB_Name = "Cipher"
Dim x1a0(9) As Long
Dim cle(17) As Long
Dim x1a2 As Long

Dim inter As Long, res As Long, ax As Long, bx As Long
Dim cx As Long, dx As Long, si As Long, tmp As Long
Dim i As Long, c As Byte
Sub Assemble()
On Error Resume Next

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
    
If Err Then GrabarLog "Assemble", Err.Number & " " & Err.Description, "Cipher"
End Sub
Sub Code()
On Error Resume Next

    dx = (x1a2 + i) Mod 65536
    ax = x1a0(i)
    cx = &H15A
    bx = &H4E35

    tmp = ax
    ax = si
    si = tmp

    tmp = ax
    ax = dx
    dx = tmp

    If (ax <> 0) Then
        ax = (ax * bx) Mod 65536
    End If

    tmp = ax
    ax = cx
    cx = tmp

    If (ax <> 0) Then
        ax = (ax * si) Mod 65536
        cx = (ax + cx) Mod 65536
    End If

    tmp = ax
    ax = si
    si = tmp
    ax = (ax * bx) Mod 65536
    dx = (cx + dx) Mod 65536

    ax = ax + 1

    x1a2 = dx
    x1a0(i) = ax

    res = ax Xor dx
    i = i + 1
    
If Err Then GrabarLog "Code", Err.Number & " " & Err.Description, "Cipher"
End Sub
Function Encriptar(ByVal inp As String, ByVal Key As String) As String
On Error Resume Next
    
    Encriptar = ""
    si = 0
    x1a2 = 0
    i = 0

    For fois = 1 To 16
        cle(fois) = 0
    Next fois

    champ1 = Key
    lngchamp1 = Len(champ1)

    For fois = 1 To lngchamp1
        cle(fois) = Asc(Mid(champ1, fois, 1))
    Next fois

    champ1 = inp
    lngchamp1 = Len(champ1)
    
    For fois = 1 To lngchamp1
        c = Asc(Mid(champ1, fois, 1))
        Assemble

        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256

        For compte = 1 To 16

            cle(compte) = cle(compte) Xor c

        Next compte

        c = c Xor (cfc Xor cfd)
    
        d = (((c / 16) * 16) - (c Mod 16)) / 16
        e = c Mod 16

        Encriptar = Encriptar + Chr$(&H61 + d) ' d+&h61 give one letter range from a to p for the 4 high bits of c
        Encriptar = Encriptar + Chr$(&H61 + e) ' e+&h61 give one letter range from a to p for the 4 low bits of c


    Next fois
    
If Err Then GrabarLog "Encriptar", Err.Number & " " & Err.Description, "Cipher"
End Function
Function DesEncriptar(ByVal inp As String, ByVal Key As String) As String
On Error Resume Next

    DesEncriptar = ""
    si = 0
    x1a2 = 0
    i = 0

    For fois = 1 To 16
        cle(fois) = 0
    Next fois

    champ1 = Key
    lngchamp1 = Len(champ1)

    For fois = 1 To lngchamp1
        cle(fois) = Asc(Mid(champ1, fois, 1))
    Next fois

    champ1 = inp
    lngchamp1 = Len(champ1)

    For fois = 1 To lngchamp1

        d = Asc(Mid(champ1, fois, 1))
        If (d - &H61) >= 0 Then
            d = d - &H61  ' to transform the letter to the 4 high bits of c
            If (d >= 0) And (d <= 15) Then
                d = d * 16
            End If
        End If
        If (fois <> lngchamp1) Then
            fois = fois + 1
        End If
        e = Asc(Mid(champ1, fois, 1))
        If (e - &H61) >= 0 Then
            e = e - &H61 ' to transform the letter to the 4 low bits of c
            If (e >= 0) And (e <= 15) Then
                c = d + e
            End If
        End If
    
        Assemble

        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256

        c = c Xor (cfc Xor cfd)

        For compte = 1 To 16

            cle(compte) = cle(compte) Xor c

        Next compte

        DesEncriptar = DesEncriptar + Chr$(c)

    Next fois
    
If Err Then GrabarLog "DesEncriptar", Err.Number & " " & Err.Description, "Cipher"
End Function
