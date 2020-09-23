Attribute VB_Name = "PointsDraw"
Option Explicit
Public Sub DrawP5()

Dim a1 As Double, B As Double, co(0 To 255) As Long, coc(0 To 9) As Long, cocS(0 To 255) As Long, tm As Integer
Dim Tm2 As Long, X As Long, Y As Long, yen As Long, yst As Long


    For a1 = 0 To 255
'        co(a1) = RGB(a1 ^ 2 \ 32, a1 ^ 2 \ 48, a1 ^ 2 \ 24)
'        co(a1) = RGB(a1 ^ 2 \ (8 * txtspm(10)), a1 ^ 2 \ (8 * txtspm(10)), a1 ^ 2 \ (8 * txtspm(10)))
        co(a1) = RGB(a1 ^ 2 \ 3, a1 ^ 2 \ 3, a1 ^ 2 \ 2)
    Next a1
    For Y = yst To yen Step (PrK(3, LQT)) '* Sin(PrK(3, LQT)) + 1
        
        tm = PrK(3, (Y)) * 2
        Tm2 = PrK(2, (Y)) * 2
      
'      tm = Sin((PrK(3, (y)) - PrK(2, (y))) * Rad) * (PrK(3, (y)) - PrK(2, (y))) * 3  ''* 4
'      Tm2 = Sin((PrK(2, (y))) * Rad) * (PrK(2, (y))) * 3   '     PrK(2, (y)) * 2
      SetPixel picTmp.hdc, Sin(tm * rad * LQT * rad) * Tm2 + 512, Cos(tm * rad * Y * rad) * Tm2 + 384, co(Abs(Tm2) \ 4) Xor ColTp(100)
    Next Y
'
'    yen = 148931
'    yst = 1
'
'    For y = yst To yen Step PrK(3, LQT) + 1  '+ PrK(3, LQT) + 1)
'        tm = (PrK(3, (y)) - PrK(2, (y))) ''* 4
'        Tm2 = PrK(2, (y)) * 2
'        SetPixel picTmp.hdc, Cos(LQT * Rad + (y - 256) * Rad) * Tm2 + 512, Sin(LQT * Rad + (y - 256) * Rad) * tm + 384, co(Tm2 \ 4) Xor ColTp(tm \ 4 + 1)
'
''      tm = (PrK(3, (y)) - PrK(2, (y))) * 4 ''* 4
''      Tm2 = PrK(2, (y)) * 4
''      SetPixel picTmp.hdc, Cos(LQT * Rad + (y - 256) * Rad) * tm + 512, Sin(LQT * Rad + (y - 256) * Rad) * tm + 384, co(Tm2 \ 4) 'Xor ColTp(tm \ 4 + 1)
'    Next y
'

End Sub
''
''
'Public Sub DrawP2()
'Dim LAvg As Long, RAvg As Long
'
'If Fst Then InitK: Cof_X = 1: iH = 384    'initial at first time
'
'    ''''''''''''' set parameters variables '''''''''''''
'
'    CycleST
'
'        Cu = frmBase.txtspm(1) + 1
'        Xtmp = frmBase.txtspm(6)
'        z11 = vsY * 4
'        z22 = vsY * 4 '* (Bass / 20 - Treb / 10)
'        z33 = frmBase.txtspm(5) * 4
'        z44 = vsX ' / 2 '* 8
'        zvF = 0: zvT = 1
'        zV = z33   ''''' If frmbase.chkInc is not set Then zV = 1 and not use in loop _
'                          '     Else use  >> Ss?O(x, d) * zV
'        z33 = z33 / 2
'        If frmBase.chkInc.Value <> 0 Then zvF = 1: zvT = 0
'
'    ''''''''''''''''''
'    '''''''''''''''''' set y points of master polyline with data . Pt( , 1) is Master polyline ''''''''''''''''''
'    '''''''''''''''''' x points only set in load in InitK Sub in first time '
''        x=frmBase.txtspm(0)*
'    x2 = (384 - frmBase.txtspm(0) / 2) + frmBase.txtspm(0) / 2 '/ 2
'
'     For x = 0 To 255
'        d = SsPtr - x * Xtmp
'        zV = zvF * (((x + 1) / 32) * z33) + (zvT * zV)
'
'        Pt(255 - x, 1).y = (SsLO(x, d) * zV) - (67 * zV) + x2
'        Pt(256 + x, 1).y = (SsRO(x, d) * zV) - (67 * zV) + x2
'
'     Next x
'
'    p(1).Xx = Cosine((LQT))                   '    p(1).Xx = Cos(LQT * Rad)
'    p(1).Yy = Sine((LQT))                     '    p(1).Yy = Sin(LQT * Rad)
'    p(1).Zz = (Sine((LQT)) - Cosine((LQT))) '/ -0.75                         '    Cosine(Log(LQT) / 4 * LQT + LQT)
'
''     For x = 0 To 511
'''        Pt(x, 1).y = Sin(Pt(x, 1).y * Rad) * 384 '* Cos(LQT * Rad)
'''        Pt(511 - x, 1).x = 0 '(x - 256) * 0.25 + 512 ' Pt(x, 1).x  ' * Pt(x, 1).x * 0.5  '- 512 '* Sin(LQT * Rad)
'''        Pt(x, 1).x = 0
''     Next x
'
'    CycleED
'    Process(12, 1) = Round(tFa, 2)
'
'    '''''''''''''''''' set y points in other polylines with Master polyline ''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''
'
'    CycleST
'  ''''''''''''''''''''''''''
'    bf4 = z22
'    For x2 = 2 To Cu  ' Cu is Count of Scopes
'       z22 = (z22 * 0.95)
'        For x = 0 To 255
'            Pt(255 - x, x2).y = Pt(255 - x, 1).y + z22
'            Pt(256 + x, x2).y = Pt(256 + x, 1).y + z22
'            Pt(255 - x, x2).x = Sin(Pt(255 - x, x2).x * rad) * Cos(Pt(255 - x, x2).y) + 512 - x ' + Pt(255 - x, x2)
'            Pt(256 + x, x2).x = Sin(Pt(256 + x, x2).x * rad) * Cos(Pt(256 + x, x2).y) + 512 + x '+ Pt(256 + x, x2)
'        Next x
'    Next x2
'    ''''''''''''''''''''''''''
'    CycleED
'    Process(13, 1) = Round(tFa, 2)
'
'
'    '''''''''''''''''' find  minY and maxY points of Master polyline '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''
'    CycleST
'
'    minLY = minY: maxLY = maxY
'    minY = 768: maxY = 0
'
'    For x = 0 To 511
'        If maxY < Pt(x, 2).y Then maxY = Pt(x, 2).y
'        If minY > Pt(x, Cu).y Then minY = Pt(x, Cu).y
'    Next x
'    If maxY < 1 Then maxY = 1
'    If minY > 768 Then minY = 768
'
'    ''''''''''''''''''''''''''''''''''''''''''
'    '''''''''''''''''' Set BaseSub Height for scopes '''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''
'
'    If frmBase.chkABalance And ((maxY - minY) < (frmBase.txtspm(0) / 2) Or (maxY - minY) > (frmBase.txtspm(0) * 2)) Then frmBase.chkAHeight.Value = 1
'
'    If frmBase.chkAHeight Then
'        If (maxY - minY) > (frmBase.txtspm(0) * 2) Then
'            frmBase.cmdSmaler_Click (5): frmBase.txtspm(5).Refresh
'        ElseIf (maxY - minY) < (frmBase.txtspm(0) / 2) Then
'            frmBase.cmdLarger_Click (5): frmBase.txtspm(5).Refresh
'        End If
'    End If
'
'    minY = minY - (maxY - minY) / 8 - 64
'    maxY = maxY + (maxY - minY) / 8 + 64
'
'
'    Call CycleED
'    Process(15, 1) = Round(tFa, 2)
'
'
'    '''''''''''''''' Clear last polyline if chkCls1 is checked ''''''''''''''''
'    CycleST
'
''        If frmBase.chkCls1 Then BitBlt picBuff.hdc, 0, 0, 1024, 768, picBuff.hdc, 0, 0, vbBlackness
'
'        If frmBase.ChkDraw(4) Then   '''''   last polyline clear
'            picBuff.ForeColor = vbBlack
'            picBuff.FillStyle = vbSolid
'            picBuff.FillColor = vbBlack
'
'           If frmBase.ChkDraw(2) Then
'                For x = 2 To Cu
'                 Polygon picBuff.hdc, PtL(1, x), 255
'                 Polygon picBuff.hdc, PtL(256, x), 255
'                Next x
'           Else
'                For x = 2 To Cu
'                 Polyline picBuff.hdc, PtL(1, x), 255
'                 Polyline picBuff.hdc, PtL(256, x), 255
'                Next x
'           End If
'        End If
'
'    '''''''''''''''' draw Polylines ''''''''''''''''
'        For x = 2 To Cu
'             picBuff.ForeColor = ColTn(x) Xor vbYellow
'             picBuff.FillStyle = vbSolid
'             picBuff.FillColor = ColTn(x) Xor vbYellow
'
'             If frmBase.ChkDraw(2) Then
'                Polygon picBuff.hdc, Pt(1, x), 255
'                Polygon picBuff.hdc, Pt(256, x), 255
'             Else
'                Polyline picBuff.hdc, Pt(1, x), 255
'                Polyline picBuff.hdc, Pt(256, x), 255
'             End If
'        Next x
'
'    '''''''''''''''''' store data for use in next polylines in next frame ''''''''''''''''''
'
'        For x = 2 To Cu
'            CopyMemory PtL(0, x).x, Pt(0, x).x, 4096
'        Next x
'
'    Call CycleED
'    Process(14, 1) = Round(tFa, 2)
'
'End Sub

Public Sub DrawP3()
Dim ct As Long, Z As Single, d As Single, u As Single, X As Single, Xx As Integer
Dim tm1 As Single, Tm2 As Single, Tm3 As Single, Tm4 As Single, m As Long

CycleST

'''''''''''''''''''''''''''''''''''
    P(1).Xx = Cosine((LQT))                   '    p(1).Xx = Cos(LQT * Rad)
    P(1).Yy = Sine((LQT))                     '    p(1).Yy = Sin(LQT * Rad)
    P(1).Zz = (Sine((LQT)) - Cosine((LQT))) '/ -0.75                         '    Cosine(Log(LQT) / 4 * LQT + LQT)
'''''''''''''''''''''''''''''''''''
    For ct = 1 To 6
    
        If ct Mod 3 = 0 Then
            P(1).Y = P(1).Yy
            P(1).X = P(1).Zz
        ElseIf ct Mod 3 = 1 Then
            P(1).Y = P(1).Zz
            P(1).X = P(1).Xx
        ElseIf ct Mod 3 = 2 Then
            P(1).Y = P(1).Xx
            P(1).X = P(1).Yy
        End If
     ''''''''''''''''''''''''''''''''''''''''''''
     ''''''''''''''''''''''''''''''''''''''''''''
     
        For X = 1 To PrK(3, LQT) '32 '* 2
            If ct Mod 3 = 0 Then
              P(ct).Pt(X).Y = Cos((Log(ABass + 1) * X * LQT) * rad) * (256 \ Log(ct + 1)) * P(1).Yy + 384 '+ 192 * Sin(LQT * Rad) ^ 3
              P(ct).Pt(X).X = Sin((Log(ABass + 1) * X * LQT) * rad) * (256 \ Log(ct + 1)) * P(1).Xx + 512 '+ 256 * Cos(LQT * Rad) ^ 3
            ElseIf ct Mod 3 = 1 Then
              P(ct).Pt(X).Y = Cos((Log(ABass + 1) + X * LQT) * rad) * (256 \ Log(ct + 1)) * P(1).Yy + 384 '+ 192 * Sin(LQT * Rad) ^ 3
              P(ct).Pt(X).X = Sin((Log(ABass + 1) + X * LQT) * rad) * (256 \ Log(ct + 1)) * P(1).Zz + 512 '+ 256 * Cos(LQT * Rad) ^ 3
            ElseIf ct Mod 3 = 2 Then
              P(ct).Pt(X).Y = Cos((Log(ABass + 1) + X * LQT) * rad) * (256 \ Log(ct + 1)) * P(1).Zz + 384 '+ 192 * Sin(LQT * Rad) ^ 3
              P(ct).Pt(X).X = Sin((Log(ABass + 1) + X * LQT) * rad) * (256 \ Log(ct + 1)) * P(1).Xx + 512 '+ 256 * Cos(LQT * Rad) ^ 3
            End If
        Next X
    
    Next ct
    '''''''''''''''''''''''''''''''
    CycleED
    Process(16, 1) = Round(tFa, 2)
    CycleST
    
    If frmBase.ChkDraw(5) Then
    picTmp.ForeColor = vbBlack
    
    For ct = 1 To 6
        Polyline picTmp.hdc, P(ct).PtL(1), PrK(3, LQT)
    Next ct
    
    End If
    '''''''''''''''''''''''''''''''
        
    For ct = 1 To 6
      picTmp.ForeColor = ColTn(ct) 'Xor ColTp(ct) ' Xor vbBlack
      Polyline picTmp.hdc, P(ct).Pt(1), PrK(3, LQT)
    Next ct
    '''''''''''''''''''''''''''''''
    
    For ct = 1 To 6
        CopyMemory P(ct).PtL(1).X, P(ct).Pt(1).X, 2048
    Next ct

    lastSTP = PrK(2, LQT)
    
    CycleED
    Process(18, 1) = Round(tFa, 2)
    
End Sub
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'
Public Sub DrawP4()
Dim ct As Long, Z As Single, d As Single, u As Single, X As Single, Xx As Integer
Dim tm1 As Single, Tm2 As Single, Tm3 As Single, Tm4 As Single, m As Long


    P(1).Col = P(1).Col + P(1).ColV * Cos((BassL - BassR) * (BassR - BassL)) * 2
    If P(1).Col >= 255 Or P(1).Col <= 0 Then P(1).ColV = -P(1).ColV

     P(1).Tmz = Cos((P(1).Col * rad))
     P(1).Zm = Sin((LQT * rad) * (LQT * rad) * rad)
     P(1).mx = P(1).Col * P(1).Zm
     P(1).mY = P(1).Col * P(1).Zm

     P(1).X = Cos(LQT * rad) * P(1).Tmz
     P(1).Y = Sin(LQT * rad) * P(1).Tmz
     P(1).Xx = Sin(LQT * rad) * P(1).Tmz
     P(1).Yy = Cos(LQT * rad) * P(1).Tmz

     P(1).Tm2 = (1 + -2 * frmBase.chkP4Opt(0))
     P(1).Tm4 = (1 + -2 * frmBase.chkP4Opt(1))
     P(1).tm1 = P(1).Tmz * P(1).Tm2                'rotate +-z
     P(1).Tm3 = P(1).Tmz * P(1).Tm4              'rotate -+z


    P(1).Xx = Cos(LQT * rad)
    P(1).Yy = Sin(LQT * rad)
    P(1).Zz = (P(1).Xx) - (P(1).Yy)

    For ct = 30 To 43 Step 1
            If ct Mod 5 = 0 Then
               P(1).Y = P(1).Yy
               P(1).X = P(1).Xx
            ElseIf ct Mod 5 = 1 Then
               P(1).Y = P(1).Zz
               P(1).X = P(1).Xx
            ElseIf ct Mod 5 = 2 Then
               P(1).Y = P(1).Yy
               P(1).X = P(1).Zz
            ElseIf ct Mod 5 = 3 Then
               P(1).Y = P(1).Zz
               P(1).X = P(1).Yy
            Else
               P(1).Y = P(1).Xx
               P(1).X = P(1).Zz
            End If
     ''''''''''''''''''''''''''''''''''''''''''''
     ''''''''''''''''''''''''''''''''''''''''''''

            For Xx = 1 To maxViewP
              P(ct).Pt(Xx).Y = Sin(2 + Xx + LQT * rad * P(1).Tm2) * ((ct - 10) * 8) * (1 + -2 * (Xx Mod 2)) * (P(1).Y) + 384      '* (1 + -2 * (Xx Mod 2))
              P(ct).Pt(Xx).X = Cos(Xx + LQT * rad * P(1).Tm4) * ((ct - 10) * 8) * (1 + -2 * (Xx Mod 2)) * (P(1).X) + 512       '* (1 + -2 * (Xx Mod 2))
            Next Xx

    Next ct
    '''''''''''''''''''''''''''''''

    If frmBase.ChkDraw(7) Then
       picTmp.ForeColor = vbBlack
       For ct = 30 To 43
         Polyline picTmp.hdc, P(ct).PtL(1), maxViewP
       Next ct
    End If
    '''''''''''''''''''''''''''''''

    For ct = 30 To 43
      picTmp.ForeColor = ColTn(ct - 29) And (ColTp(1) Xor vbWhite)
      Polyline picTmp.hdc, P(ct).Pt(1), maxViewP
    Next ct
    '''''''''''''''''''''''''''''''

    For ct = 30 To 43
        CopyMemory P(ct).PtL(1).X, P(ct).Pt(1).X, 1024 * 2
    Next ct


End Sub
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'Public Sub DrawP4()
'Dim ct As Integer, Z As Integer, d As Single, u As Single
'
'    P3P.Col = P3P.Col + P3P.ColV * Cos((BassL - BassR) * (BassR - BassL) / 2)
'    If P3P.Col > 255 Or P3P.Col < 0 Then P3P.ColV = -P3P.ColV
'
'    P4P.Tmz = Sin((P4P.Col) * Rad) * 0.78    '    z = Sin((LQT) * Rad)
'    P4P.Zm = Sin((LQT) * Rad)
'    P4P.x = Cos((LQT) * Rad) * Sin((P4P.Col) * Rad) * 0.78
'    P4P.y = Sin((LQT) * Rad) * Cos((P4P.Col) * Rad) * 0.78
'    P4P.xx = Cos((LQT) * Rad) * Cos((P4P.Col) * Rad) * 0.78
'    P4P.yy = Sin((LQT) * Rad) * Sin((P4P.Col) * Rad) * 0.78
'    P4P.Tm2 = (1 + -2 * frmBase.chkP4Opt(0))               'Height
'    P4P.Tm4 = (1 + -2 * frmBase.chkP4Opt(1))               'width
'    P4P.tm1 = P4P.Tmz * P4P.Tm2 * PI                                 'rotate +-z
'    P4P.Tm3 = P4P.Tmz * P4P.Tm4 * PI                                 'rotate -+z
'
''    For ct = 0 To 511
'''        Pt(ct, 18).Y = Sin(ct + P4P.tm1) * Pt(ct, 1).Y * P4P.Y + 384
'''        Pt(ct, 18).x = Cos(ct + P4P.tm3) * Pt(ct, 1).x * P4P.x + 512
''
'''        Pt(ct, 19).Y = Sin(ct + P3P.tm1) * Pt(ct, 1).Y * P3P.yy + 384
'''        Pt(ct, 19).x = Cos(ct + P3P.tm3) * Pt(ct, 1).x * P3P.xx + 512
''    Next ct
'
'    For ct = 0 To 511
'         Z = Z + d
'        Pt(ct, 18).y = Sin(ct * Rad * PI * (Treb - Bass)) / _
'                       ((PI * MidlL + 1) / (1 + Treb / 1.5)) * _
'                       (Pt(ct, 1).y * 0.75) + 384 + 384 * P4P.y
'
'        Pt(ct, 18).x = Cos(ct * Rad * PI * (Bass - Treb)) / _
'                       ((PI * MidlR + 1) / (1 + Bass / 1.5)) * _
'                       (Pt(ct, 1).x * 1) + 512 + 512 * P4P.xx
''        Pt(ct, 18).Y = Sin(ct + P4P.tm1) * Pt(ct, 1).Y * P4P.Y + 384
''        Pt(ct, 18).x = Cos(ct + P4P.tm3) * Pt(ct, 1).x * P4P.x + 512
'
'    Next ct
'
'        If frmBase.ChkDraw(7) Then
'            picTmp.ForeColor = vbBlack
'            Polyline picTmp.hdc, PtL(32, 18), 32
'            Polyline picTmp.hdc, PtL(128, 18), 128
'        End If
'        picTmp.ForeColor = ColTp(1) Xor ColTn(1) ' ColTp(1) Xor vbCyan
'        Polyline picTmp.hdc, Pt(32, 18), 32
'        picTmp.ForeColor = ColTp(1) Xor ColTn(1) ' ColTp(1) Xor vbMagenta
'        Polyline picTmp.hdc, Pt(128, 18), 128
'
'    CopyMemory PtL(0, 18).x, Pt(0, 18).x, 4096
'
'End Sub
'
'
'
'Public Sub DrawP1()
'
'Dim A As Double, B As Double, coc(0 To 9) As Long, cocS(0 To 255) As Long, BStp As Long
'Dim Stp As Long, x As Double, y As Double, yen As Long, yst As Long
'Dim rad As Double, SI1 As Double, SI2 As Double, SI3 As Double, SI4 As Double, SI5 As Double
'Dim pd As Double, cTim As Double, xTm As Double, yTm As Double
'Dim E As Long, m As Long, Cc As Long, Ti2 As Long, Ti3 As Long, Ti4 As Long, Ti5 As Long
'Dim xCntr As Double, yCntr As Double, w(1 To 10) As Double
'Dim bBol(1 To 10) As Byte, tmp1 As Double, tmp2 As Double, tmp3 As Double, tmp4 As Double
'Dim pot As POINTAPI, co As Long, co2 As Long, co3 As Long
'Dim cA As Long, cR As Long, cG As Long, cb As Long, S As String
'
'On Error Resume Next
'
'    If frmBase.chkShotAll.Value And frmBase.chkAutoShot.Value Then frmBase.cmdSF_Click
'
'    LQT = LQT + 1
'    If LQT > 148900 Then LQT = 1
'    If LQT2 > 148900 Then LQT2 = 1
'    frmBase.txtspm(13) = LQT
'    pd = (frmBase.txtspm(2))
'    If frmBase.chkAvalue Then LQT2 = LQT2 + pd
'
'    frmBase.txtspm(11) = LQT2: frmBase.txtspm(11).Refresh
'
'    With frmBase
'        .txtLQT2 = Format$(Abs(LQT2), "###,###0") & vbCrLf & _
'                Format$(Primes(Abs(LQT2)), "###,###,###0") & vbCrLf & _
'                Format$(PrK(3, Abs(LQT2)), "####") & vbCrLf & _
'                Format$(PrK(2, Abs(LQT2)), "####")
'        .txtLQT2.Refresh
'
'    End With
'
'    '''''''''''''''''''''''''''''''''''''''
'    '''''''''''''''''''''''''''''''''''''''
'
'    If frmBase.ChkDraw(4) Then BitBlt picTmp.hdc, 0, 0, 1024, 768, picTmp.hdc, 0, 0, 0
'
'    If frmBase.chkAutoMax Then
'          yen = frmBase.txtspm(11)           ' 148932
'        Else
'          yen = frmBase.txtspm(21)
'    End If
'    '''''''''''''''''''''''''''''''''''''''
'    If frmBase.chkLastP Then
'         yst = yen - frmBase.txtspm(32)
'        Else
'         yst = frmBase.txtspm(28)
'    End If
'
'    cTim = LQT2 / 100000 ' 86400
'
'    SI1 = frmBase.txtspm(16)
'    SI2 = frmBase.txtspm(17)
'    SI3 = frmBase.txtspm(18)
'    SI4 = frmBase.txtspm(19)
'    SI5 = frmBase.txtspm(20)
'
'
'    bBol(1) = frmBase.chkCol(1)
'    bBol(2) = frmBase.chkCol(2)
'    bBol(3) = frmBase.chkCol(3)
'    bBol(4) = frmBase.chkCol(4)
'
'    bBol(6) = frmBase.chkBox
'
'    xCntr = frmBase.txtspm(22)
'    yCntr = frmBase.txtspm(23)
'
'    co2 = RGB(ColPR, ColPG, ColPB)
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    For B = yst To yen Step 1
'
'        BStp = PrK(3, B)
'        Stp = PrK(2, B)
'
'        m = PrK(2, B) * SI2
'        Cc = PrK(3, B) * SI3
'        Cc = Cc * Cc * SI4
'        E = (m * Cc) * SI1
'        co2 = RGB(m * Log(m), E / Log(E), Cc / Log(Cc)) ' RGB(ColPR, ColPG, ColPB)
'
'        co = RGB(m * Log(m), Cc, E / Log(E))
'
'        If frmBase.chkTimeEnable(0) Then
'            x = Sin(E * cTim) * Cos(m * cTim) * Cc + Stp * Sin(B) + xCntr
'            y = Cos(E * cTim) * Sin(m * cTim) * Cc + Stp * Cos(B) + yCntr
'            co = RGB(m, Cc, E / Log(E))
'            picTmp.ForeColor = co Xor vbGrayed
'            LineTo picTmp.hdc, x, y
'        End If
'
'        x = Sin(E * cTim) * Cos(m * cTim) * Cc + xCntr
'        y = Cos(E * cTim) * Cos(m * cTim) * Cc + yCntr
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        For Ti2 = 1 To PrK(2, B)
'            m = PrK(2, B) * SI2
'            Cc = PrK(3, B) * SI3
'            Cc = Cc * Cc * SI4
'            E = (m * Cc) * SI1
'
'            co2 = RGB(m * Log(m), E / Log(E), Cc / Log(Cc))
'            co = co Xor co2
'            xTm = Cos(E * cTim - Ti2) * Cos(Cc * Ti2) * PrK(2, E) + x
'            yTm = Sin(E * cTim - Ti2) * Cos(Cc * Ti2) * PrK(2, E) + y
'            SetPixel picTmp.hdc, xTm, yTm, co
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            For Ti3 = 1 To PrK(2, Ti2)
'                m = PrK(2, Ti2) * SI2
'                Cc = PrK(3, Ti2) * SI3
'                Cc = Cc * Cc * SI4
'                E = (m * Cc) * SI1
'
'                co2 = RGB(E / Log(E), Cc / Log(Cc), m * Log(m))
'                co = co Xor co2
'                x = Cos(E * Ti3 - Ti2) * Cos(m * Ti3 * cTim) * PrK(2, E) + xTm
'                y = Sin(E * Ti3 - Ti2) * Cos(m * Ti3 * cTim) * PrK(2, E) + yTm
'                SetPixel picTmp.hdc, x, y, co
'            Next Ti3
'          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''        Ti2 = Ti2 + 1 * SI5
'        Next Ti2
'      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       co = RGB(m * Log(m), Cc, E / Log(E))
'       SetPixel picTmp.hdc, x, y, co
'       B = B + 1 * SI5
'    Next B
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    If ColPR >= 255 Then ColPRsgn = -1: ColPR = 255
'    If ColPG >= 255 Then ColPGsgn = -1: ColPG = 255
'    If ColPB >= 255 Then ColPBsgn = -1: ColPB = 255
'    If ColPR <= 1 Then ColPRsgn = 1: ColPR = 1
'    If ColPG <= 1 Then ColPGsgn = 1: ColPG = 1
'    If ColPB <= 1 Then ColPBsgn = 1: ColPB = 1
'    ColPR = ColPR + ColPRsgn * 2 * (Log(PrK(2, LQT2))) / 7
'    ColPG = ColPG + ColPGsgn * 3 * (Log(PrK(2, LQT2))) / 7
'    ColPB = ColPB + ColPBsgn * 5 * (Log(PrK(2, LQT2))) / 7
'
'CycleST
'
'     Blend.SourceConstantAlpha = Val(frmBase.txtspm(7))
'     Blend.AlphaFormat = 0
'     If frmBase.chkAlphaEnable And ((LQT2 * 1234) Mod 3 = 1) Then Blend.AlphaFormat = 1
'
'     If frmBase.chkAlpha Then Blend.SourceConstantAlpha = ((Colv_R + Colv_G + Colv_B)) / 3
'     CopyMemory BlendPtr, Blend, 4
'
'     i = frmBase.txtspm(25)
'         StretchBlt frmBase.picBuffEE.hdc, 0, 0, frmBase.Width \ 15 + 1, frmBase.Height \ 15 + 1, _
'              picBuff.hdc, 0, 0, 1024 / i, 768 / i, vbSrcCopy
'
'         AlphaBlend picView.hdc, 0, 0, 1024, 768, _
'              frmBase.picBuffEE.hdc, 0, 0, 1024, 768, BlendPtr
'
''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
'
'     Nvg(1) = frmBase.txtNavigate(1): Nvg(2) = frmBase.txtNavigate(2): Nvg(3) = frmBase.txtNavigate(3)
'
'     If frmBase.fraTelo.Visible = True Then
'         AlphaBlend frmBase.picTele.hdc, 0, 0, 256, 256, _
'         frmBase.picBuffEE.hdc, Nvg(2), Nvg(3), Nvg(1), Nvg(1), BlendPtr
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
''    If frmBase.chkALog And frmBase.chkAvalue Then Loger
'
'End Sub
''


