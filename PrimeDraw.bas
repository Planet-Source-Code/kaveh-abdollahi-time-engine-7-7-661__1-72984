Attribute VB_Name = "PrimeDraw"

Option Explicit

Private A As Double, B As Double, coc(0 To 9) As Long, cocS(0 To 255) As Long, BStp As Long
Private Stp As Long, X As Double, Y As Double, yen As Long, yst As Long
Private rad As Double, SI1 As Double, SI2 As Double, SI3 As Double, SI4 As Double, SI5 As Double, SI6 As Double
Private pd As Double, cTim As Double, xTm As Double, yTm As Double
Private E As Long, m As Long, Cc As Long, Ti2 As Long, Ti3 As Long, Ti4 As Long, Ti5 As Long
Private xCntr As Double, yCntr As Double, w(1 To 10) As Double, sz As Long
Private bBol(1 To 10) As Byte, tmp1 As Double, tmp2 As Double, tmp3 As Double, tmp4 As Double
Private pot As POINTAPI, co As Long, co2 As Long, co3 As Long, co4 As Long, co5 As Long
Private cA As Long, cR As Long, cG As Long, cb As Long, S As String

Private Type dt
   H1               As String * 30
   H2               As Long
   H3               As Long
   sT(1 To 64000)   As Byte
End Type
Public Keyp As dt


Public Sub DrawP1()
'Dim Elc As Double, Pur As Double, Nut As Double
'
'Elc = 1E-31 * 9.10938188
'Pur = 1E-27 * 1.67262158
'Nut = 1E-27 * 1.67492716
    With frmBase
    
    On Error Resume Next
    If .chkShotAll.Value And .chkAutoShot.Value Then .cmdSF_Click


    LQT = LQT + 1
    If LQT > 148900 Then LQT = 1
    If LQT2 > 148900 Then LQT2 = 1
    .txtspm(13) = LQT
    pd = (.txtspm(2))
    If .chkAvalue Then LQT2 = LQT2 + pd
    .txtspm(11) = LQT2: .txtspm(11).Refresh
    .txtLQT2 = Format$(LQT2, "###,###0") & vbCrLf & _
        Format$(Primes(LQT2), "###,###,###0") & vbCrLf & _
        Format$(PrK(3, LQT2), "####") & vbCrLf & _
        Format$(PrK(2, LQT2), "####") & vbCrLf & _
        Format$(sz, "###,###,###0")
    .txtLQT2.Refresh
    
    .lblFullscr(1) = Round(LQT2, 2): .lblFullscr(1).Refresh
    .lblFullscr(2) = Primes(LQT2): .lblFullscr(2).Refresh
    .lblFullscr(3) = PrK(2, LQT2): .lblFullscr(3).Refresh
    .lblFullscr(4) = PrK(3, LQT2): .lblFullscr(4).Refresh
    
    sz = 0
    '''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''
    
    If .ChkDraw(4) Then BitBlt picTmp.hdc, 0, 0, 1024, 768, picTmp.hdc, 0, 0, 0
    
    '''''''''''''''
    If .chkAgr(0) Then .txtspm(19) = (.txtspm(19) + 1 / (ABass + 1) / Stp) / 5: _
         .txtspm(17) = (.txtspm(17) + 1 / (MidlL - MidlR) / Stp) / 5
    '''''''''''''''
    If .chkAgr(1) Then .txtspm(19) = (.txtspm(19) + 1 / (ABass + 1) / Stp) / 5: _
         .txtspm(16) = (.txtspm(16) + 1 / (MidlL - MidlR) / Stp) / 5

    
    
    If .chkAutoMax Then
          yen = .txtspm(11)
        Else
          yen = .txtspm(21)
    End If
    '''''''''''''''''''''''''''''''''''''''
    If .chkLastP Then
         yst = yen - .txtspm(32)
      Else
         yst = .txtspm(28)
    End If

    cTim = LQT2 / 1000000
    SI1 = 1: SI2 = 1: SI3 = 1: SI4 = 1: SI5 = 1: SI6 = 1
    If IsNumeric(.txtspm(16)) Then SI1 = .txtspm(16) '* 2
    If IsNumeric(.txtspm(17)) Then SI2 = .txtspm(17) '* 2
    If IsNumeric(.txtspm(18)) Then SI3 = .txtspm(18) '* 2
    If IsNumeric(.txtspm(19)) Then SI4 = .txtspm(19) '* 2
    If IsNumeric(.txtspm(20)) Then SI5 = .txtspm(20) '* 2
    If IsNumeric(.txtspm(33)) Then SI6 = .txtspm(33) '* 2

    bBol(1) = .chkCol(1)
    bBol(2) = .chkCol(2)
    bBol(3) = .chkCol(3)
    bBol(4) = .chkCol(4)
    bBol(5) = .chkCol(5)
    bBol(0) = .chkCol(0)

    bBol(6) = .chkBox

    xCntr = .txtspm(22)
    yCntr = .txtspm(23)
                                                                                                                
    co5 = RGB((LQT2 Mod 4096) / 16, (LQT2 Mod 2048) / 8, (LQT2 Mod 8192) / 32)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For B = yst + Stp To yen Step 1
        
        BStp = PrK(3, B)
        Stp = PrK(2, B)
'        If Stp <> 34 Then GoTo enx
        
        m = Stp * (SI2)
        Cc = BStp * (SI3)
        If .chkCM Then Cc = Stp * (SI3)
        Cc = Cc * Cc * (SI4)
        E = (m * Cc) * (SI1)
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        co4 = RGB(Stp, Stp, Stp)
        If bBol(1) Then
            co = RGB(256 - Stp * 7, 256 / Stp * 3, Stp * 5)
            co2 = RGB(256 - Stp * 3, 256 - Stp * 5, Stp * 7)
            co3 = RGB(256 - Stp * 7, 256 / Stp * 3, 256 - Stp * 3)
        ElseIf bBol(2) Then
            co = RGB(256 - Stp * 7, 256 / Stp * 7, Stp * 5)
            co2 = RGB(256 - Stp * 7, 256 / Stp * 7, Stp * 5)
            co3 = RGB(256 - Stp * 7, 256 / Stp * 7, Stp * 5)
        ElseIf bBol(3) Then
            co = RGB(256 - BStp * 7, 256 / BStp * 3, BStp * 5)
            co2 = RGB(256 - BStp * 3, 256 / BStp * 5, BStp * 7)
            co3 = RGB(256 - BStp * 7, 256 / BStp * 3, BStp * 3)
        ElseIf bBol(4) Then
            co = RGB(256 - Stp * 3, 256 / Stp * 5, Stp * 5)
            co2 = RGB(256 - Stp * 7, 256 / Stp * 3, Stp * 7)
            co3 = RGB(256 - Stp * 5, 256 / Stp * 7, 256 - Stp * 3)
        ElseIf bBol(5) Then
            co = RGB(256 - Cc / Stp, Cc / Stp, 256 - Cc / Stp)
            co2 = RGB(Cc / Stp, 256 - E / Cc, Cc / Stp)
            co3 = RGB(E / Stp, E / Stp * 2, 256 - Cc / Stp)
        ElseIf bBol(0) Then
            co = RGB(Cc, E, m)
            co2 = RGB(E, m, Cc)
            co3 = RGB(m, Cc, E)
        Else
            co = RGB(Cc / Stp, Cc / Stp, Cc / Stp)
            co2 = RGB(Cc / Stp, 256 - Cc / Stp * 2, Cc / Stp)
            co3 = RGB(E / Stp, 256 - E / Stp * 2, 256 - Cc / Stp)
        End If
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If .chkPant Then
'           X = Sin(E * cTim + B) * Cos(Cc * cTim - B) * Cc / SI6 + xCntr
'           Y = Cos(E * cTim - B) * Cos(Cc * cTim + B) * Cc / SI6 + yCntr
           X = Sin(E * cTim) * Cos(Cc * cTim) * Cc / SI6 + xCntr
           Y = Cos(E * cTim) * Cos(Cc * cTim) * Cc / SI6 + yCntr
        Else
            X = Sin(E * cTim) * Cos(Cc * cTim) * Cc / SI6 + xCntr
            Y = Cos(E * cTim) * Sin(Cc * cTim) * Cc / SI6 + yCntr
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(1) Then
            picTmp.ForeColor = co2 And co4
            picTmp.FillColor = co And co5
            xTm = X
            yTm = Y
            Ellipse picTmp.hdc, xTm - m \ SI6, yTm - m \ SI6, xTm + m \ SI6, yTm + m \ SI6
        End If
        ''''''''''''''''''''''''''''''''' Strings '''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(0) Then
            If B = yst Then MoveToEx picTmp.hdc, X, Y, pot
            picTmp.ForeColor = co3 And co5
            LineTo picTmp.hdc, X, Y
        End If
        If .chkTimeEnable(2) Then
            If B = yst Then MoveToEx picTmp.hdc, X, Y, pot
            picTmp.ForeColor = co2 Or co And co5
            LineTo picTmp.hdc, X, Y
        End If
        If .chkTimeEnable(4) Then
            If B = yst Then MoveToEx picTmp.hdc, X, Y, pot
            picTmp.ForeColor = co3 Xor co4 And co5  ' RGB(Stp, Stp, Stp)
            LineTo picTmp.hdc, X, Y
        End If
        '''''''''''''''''''''''''''''''' Strings ''''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(3) Or .chkTimeEnable(5) Then
            SetPixel picTmp.hdc, X, Y, co And co4 Xor co5: sz = sz + 1
            For Ti2 = B To B + Stp \ 3
                DoEvents
                m = PrK(2, Ti2) * SI2
                Cc = PrK(3, Ti2) * SI3
                Cc = Cc * Cc * SI4
                E = (m * Cc) * SI1
                xTm = Cos(Cc * Ti2 + cTim / Stp) * Cos(E * Ti2) * m \ SI6 + X
                yTm = Sin(Cc * Ti2) * Cos(E * Ti2 + cTim / Stp) * m \ SI6 + Y
                If .chkTimeEnable(3) Then SetPixel picTmp.hdc, xTm, yTm, co2 And co Xor co: sz = sz + 1
                
                If .chkTimeEnable(5) Then
                    For Ti3 = Ti2 To Ti2 + Stp \ 5 '17
                        m = PrK(2, Ti2) * SI2
                        Cc = PrK(3, Ti2) * SI3
                        Cc = Cc * Cc * SI4
                        E = (m * Cc) * SI1
                        X = Cos(Cc * Ti3) * Cos(E - m * cTim) * PrK(2, B) \ SI6 + X
                        Y = Sin(Cc * Ti3) * Cos(E - Cc * cTim) * PrK(2, B) \ SI6 + Y
                        SetPixel picTmp.hdc, X, Y, co Xor co5: sz = sz + 1
                     Next Ti3
                 End If
            
            Next Ti2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
enx:
      B = B + Stp * SI5
       
    Next B
    
    If .chkAutoFix Then .txtspm(33) = BStp \ 8
    If .chkALog Then Loger
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

CycleST

     Blend.SourceConstantAlpha = Val(.txtspm(7))
     Blend.AlphaFormat = 0
     If .chkAlphaEnable Then Blend.AlphaFormat = 1

     If .chkAlpha Then Blend.SourceConstantAlpha = (Stp + Cc \ m + BStp) \ 3
     CopyMemory BlendPtr, Blend, 4

     i = .txtspm(25)
         StretchBlt .picBuffEE.hdc, 0, 0, .Width / 15 + 1, .Height / 15 + 1, _
              picBuff.hdc, 0, 0, 1024 / i, 768 / i, vbSrcCopy

         AlphaBlend picView.hdc, 0, 0, 1024, 768, _
              .picBuffEE.hdc, 0, 0, 1024, 768, BlendPtr

'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''

     Nvg(1) = .txtspm(36): Nvg(2) = .txtspm(34): Nvg(3) = .txtspm(35)

     If .fraTelo.Visible = True Then
         AlphaBlend .picTele.hdc, 0, 0, 256, 256, _
         .picBuffEE.hdc, Nvg(2), Nvg(3), Nvg(1), Nvg(1), BlendPtr
     End If

     If bBol(6) Then
        picView.ForeColor = vbYellow
        MoveToEx picView.hdc, Nvg(2), Nvg(3), pot '0

        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3)
        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3) + Nvg(1)

        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)

        LineTo picView.hdc, Nvg(2), Nvg(3)
        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)
     End If

     CycleED
     Process(5, 1) = Round(tFa, 2)
     
     
     End With
End Sub

'Public Sub DrawP1()
'
'    Dim A As Double, B As Double, coc(0 To 9) As Long, cocS(0 To 255) As Long, BStp As Long
'    Dim Stp As Long, x As Double, y As Double, yen As Long, yst As Long
'    Dim rad As Double, SI1 As Double, SI2 As Double, SI3 As Double, SI4 As Double, SI5 As Double, SI6 As Double
'    Dim pd As Double, cTim As Double, xTm As Double, yTm As Double
'    Dim E As Long, m As Long, Cc As Long, Ti2 As Long, Ti3 As Long, Ti4 As Long, Ti5 As Long
'    Dim xCntr As Double, yCntr As Double, w(1 To 10) As Double
'    Dim bBol(1 To 10) As Byte, tmp1 As Double, tmp2 As Double, tmp3 As Double, tmp4 As Double
'    Dim pot As POINTAPI, co As Long, co2 As Long, co3 As Long, co4 As Long
'    Dim cA As Long, cR As Long, cG As Long, cb As Long, S As String
'
'    With frmBase
'
'    On Error Resume Next
'    If .chkShotAll.Value And .chkAutoShot.Value Then .cmdSF_Click
'
'    LQT = LQT + 1
'    If LQT > 148900 Then LQT = 1
'    If LQT2 > 148900 Then LQT2 = 1
'    .txtspm(13) = LQT
'    pd = (.txtspm(2))
'    If .chkAvalue Then LQT2 = LQT2 + pd
'    .txtspm(11) = LQT2: .txtspm(11).Refresh
'    .txtLQT2 = Format$(Abs(LQT2), "###,###0") & vbCrLf & _
'        Format$(Primes(Abs(LQT2)), "###,###,###0") & vbCrLf & _
'        Format$(PrK(3, Abs(LQT2)), "####") & vbCrLf & _
'        Format$(PrK(2, Abs(LQT2)), "####")
'    .txtLQT2.Refresh
'    '''''''''''''''''''''''''''''''''''''''
'    '''''''''''''''''''''''''''''''''''''''
'
'    If .ChkDraw(4) Then BitBlt picTmp.hdc, 0, 0, 1024, 768, picTmp.hdc, 0, 0, 0
'
'    If .chkAutoMax Then
'          yen = .txtspm(11)           ' 148932
'        Else
'          yen = .txtspm(21)
'    End If
'    '''''''''''''''''''''''''''''''''''''''
'    If .chkLastP Then
'         yst = yen - .txtspm(32)
'      Else
'         yst = .txtspm(28)
'    End If
'
'    cTim = LQT2 / 100000 ''* 0.0000000035 '
'    SI1 = 1: SI2 = 1: SI3 = 1: SI4 = 1: SI5 = 1: SI6 = 1
'    If IsNumeric(.txtspm(16)) Then SI1 = .txtspm(16)
'    If IsNumeric(.txtspm(17)) Then SI2 = .txtspm(17)
'    If IsNumeric(.txtspm(18)) Then SI3 = .txtspm(18)
'    If IsNumeric(.txtspm(19)) Then SI4 = .txtspm(19)
'    If IsNumeric(.txtspm(20)) Then SI5 = .txtspm(20)
'    If IsNumeric(.txtspm(33)) Then SI6 = .txtspm(33)
'
'    bBol(1) = .chkCol(1)
'    bBol(2) = .chkCol(2)
'    bBol(3) = .chkCol(3)
'    bBol(4) = .chkCol(4)
'
'    bBol(6) = .chkBox
'
'    xCntr = .txtspm(22)
'    yCntr = .txtspm(23)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    For B = yst To yen Step 1
'
'        BStp = PrK(3, B)
'        Stp = PrK(2, B) * 1.618034
'
'        m = Stp * (SI2)
'        Cc = BStp * (SI3)
'        If .chkCM Then Cc = Stp * (SI3)
'        Cc = Cc * Cc * (SI4)
'        E = (m * Cc) * (SI1)
'
'        If B = yen - 1 Then
'            .lblFullscr(1) = Round(B, 2): .lblFullscr(1).Refresh
'            .lblFullscr(2) = Primes(B): .lblFullscr(2).Refresh
'            .lblFullscr(3) = Stp \ 2: .lblFullscr(3).Refresh
'            .lblFullscr(4) = BStp \ 2: .lblFullscr(4).Refresh
'            .lblFullscr(5) = Cc \ 2: .lblFullscr(5).Refresh
'        End If
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If bBol(1) Then
'            co = RGB(256 - Cc / Stp, Cc / Stp, Cc / Stp)
'            co2 = RGB(Cc / Stp, 256 - Cc / Stp * 2, Cc / Stp)
'            co3 = RGB(E / Stp, 256 - E / Stp * 2, 256 - Cc / Stp)
'            co4 = RGB(Stp, Stp, Stp)
'        ElseIf bBol(2) Then
'            co2 = RGB(256 - Cc / Stp, Cc / Stp, Cc / Stp)
'            co3 = RGB(Cc / Stp, 256 - Cc / Stp * 2, Cc / Stp)
'            co = RGB(E / Stp, 256 - E / Stp * 2, 256 - Cc / Stp)
'            co4 = RGB(Stp, Stp, Stp)
'        ElseIf bBol(3) Then
'            co3 = RGB(256 - Cc / Stp, Cc / Stp, Cc / Stp)
'            co2 = RGB(Cc / Stp, 256 - Cc / Stp * 2, Cc / Stp)
'            co = RGB(E / Stp, 256 - E / Stp * 2, 256 - Cc / Stp)
'            co4 = RGB(Stp, Stp, Stp)
'        ElseIf bBol(4) Then
'            co2 = RGB(Cc / Stp, Cc / Stp, Cc / Stp)
'            co3 = RGB(Cc / Stp, 256 - Cc / Stp * 2, Cc / Stp)
'            co = RGB(256 - m / Stp, E / Stp * 2, 256 - Cc / Stp)
'            co4 = RGB(Stp, Stp, Stp)
'        Else
'            co = RGB(Cc / Stp, Cc / Stp, Cc / Stp)
'            co3 = RGB(Cc / Stp, 256 - Cc / Stp * 2, Cc / Stp)
'            co2 = RGB(E / Stp, 256 - E / Stp * 2, 256 - Cc / Stp)
'            co4 = RGB(BStp / Stp + m, BStp / Stp - m, BStp / Stp + m)
'        End If
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        x = Sin(E * cTim) * Cos(Cc * cTim) * Cc / SI6 + xCntr
'        y = Cos(E * cTim) * Sin(Cc * cTim) * Cc / SI6 + yCntr
'        If .chkTimeEnable(1) Then
'            picTmp.ForeColor = co Xor co4
'            picTmp.FillColor = co4
'            xTm = x
'            yTm = y
'            Ellipse picTmp.hdc, xTm - m \ SI6, yTm - m \ SI6, xTm + m \ SI6, yTm + m \ SI6
'        End If
'        If .chkTimeEnable(0) Then
''            If B = yst Then MoveToEx picTmp.hdc, x, y, pot
'            picTmp.ForeColor = co Xor RGB(BStp + m, BStp + m, BStp) Xor co4
'            LineTo picTmp.hdc, x, y
'        End If
'        If .chkTimeEnable(2) Then
''            If B = yst Then MoveToEx picTmp.hdc, x, y, pot
'            picTmp.ForeColor = co Xor co2
'            LineTo picTmp.hdc, x, y
'        End If
'        If .chkTimeEnable(4) Then
''            If B = yst Then MoveToEx picTmp.hdc, x, y, pot
'            picTmp.ForeColor = RGB(Stp, Stp, Stp)
'            LineTo picTmp.hdc, x, y
'        End If
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If .chkTimeEnable(3) Then
'            SetPixel picTmp.hdc, x, y, co Xor co2
'            For Ti2 = B To B + Stp
'                m = PrK(2, Ti2) * SI2
'                Cc = PrK(3, Ti2) * SI3
'                Cc = Cc * Cc * SI4
'                E = (m * Cc) * SI1
'                xTm = Cos(Cc * Ti2 + cTim) * Cos(E * Ti2) * m / SI6 + x
'                yTm = Sin(Cc * Ti2) * Cos(E * cTim - Ti2) * m / SI6 + y
'                SetPixel picTmp.hdc, xTm, yTm, co Xor co3
'                 If .chkTimeEnable(5) Then
'                    For Ti3 = Ti2 To Ti2 + Stp
'                        m = PrK(2, Ti2) * SI2
'                        Cc = PrK(3, Ti2) * SI3
'                        Cc = Cc * Cc * SI4
'                        E = (m * Cc) * SI1
'                        x = Cos(Cc * Ti3) * Cos(E * Ti2) * m / SI6 + x
'                        y = Sin(Cc * Ti2) * Cos(E * Ti3 * cTim) * m / SI6 + y
'                        SetPixel picTmp.hdc, x, y, co Xor RGB(256 - Cc / Stp, Cc / Stp, Cc / Stp)
'                     Next Ti3
'                 End If
'            Next Ti2
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        End If
'      B = B + 1 * SI5
'
'    Next B
'
'    If .chkAutoFix Then .txtspm(33) = BStp \ 8
'    If .chkALog Then Loger
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'CycleST
'
'     Blend.SourceConstantAlpha = Val(.txtspm(7))
'     Blend.AlphaFormat = 0
'     If .chkAlphaEnable Then Blend.AlphaFormat = 1
'
'     If .chkAlpha Then Blend.SourceConstantAlpha = (Stp + Cc / m + BStp) / 3
'     CopyMemory BlendPtr, Blend, 4
'
'     i = .txtspm(25)
'         StretchBlt .picBuffEE.hdc, 0, 0, .Width / 15 + 1, .Height / 15 + 1, _
'              picBuff.hdc, 0, 0, 1024 / i, 768 / i, vbSrcCopy
'
'         AlphaBlend picView.hdc, 0, 0, 1024, 768, _
'              .picBuffEE.hdc, 0, 0, 1024, 768, BlendPtr
'
''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
'
'     Nvg(1) = .txtspm(36): Nvg(2) = .txtspm(34): Nvg(3) = .txtspm(35)
'
'     If .fraTelo.Visible = True Then
'         AlphaBlend .picTele.hdc, 0, 0, 256, 256, _
'         .picBuffEE.hdc, Nvg(2), Nvg(3), Nvg(1), Nvg(1), BlendPtr
'     End If
'
'     If bBol(6) Then
'        picView.ForeColor = vbYellow
'        MoveToEx picView.hdc, Nvg(2), Nvg(3), pot '0
'
'        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3)
'        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3) + Nvg(1)
'
'        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)
'
'        LineTo picView.hdc, Nvg(2), Nvg(3)
'        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)
'     End If
'
'     CycleED
'     Process(5, 1) = Round(tFa, 2)
'
'     End With
'End Sub
