Attribute VB_Name = "Declare"
Option Explicit
 
    Public Const PI = 3.14159265358979, rad = 0.0174532925199
    
    Public Vsgn As Integer, Hsgn As Integer
    Public sPath As String, bPath As String
    Public FpS As Single, FPrcS As Single, p3c As Single, p3cV As Single, SaveCount As Long
    Public DoEv As Boolean, DoClickS As Boolean, DoClickL As Boolean, DoS As Integer, DoL As Integer, idxS As Integer, idxL As Integer
    Public Process(0 To 20, 1 To 2), ProcTi(0 To 20, 1 To 2)
    Public aa As Integer, dRF As Integer, dRFlag As Integer, BlurNum As Integer
    Public FFtL(1 To 16) As Single, FFtR(1 To 16) As Single
    Public SsL(0 To 255) As Single, SsR(0 To 255) As Single
    Public FHL As Single, FHR As Single, FLL As Single, FLR As Single, FAvrR As Single, Fsave As Single
    Public SsLtmp(0 To 255) As Single
    Public SsRtmp(0 To 255) As Single
    Public SsLO(0 To 255, -5000 To 0) As Single
    Public SsRO(0 To 255, -5000 To 0) As Single
    Public SsPtr As Integer, Angel As Integer
    Public Bass As Single, BassL As Single, BassR As Single
    Public Midl As Single, MidlL As Single, MidlR As Single
    Public Treb As Single, TrebL As Single, TrebR As Single
    Public Freq As Single, FreqL As Single, FreqR As Single
    Public ABass As Single, AMidl As Single, ATreb As Single, AFreq  As Single
    Public ABass2 As Single, AMidl2 As Single, ATreb2 As Single, AFreq2  As Single
    Public LBass As Single, LMidl As Single, LTreb As Single, LFreq  As Single
    Public Ox As Single, Oy As Single, Ox2 As Single, Oy2  As Single
    Public status As Long, status2 As Boolean, BlrF As Boolean
    Public gX As Long, gY As Long, gX2 As Long, gY2 As Long, Fla2 As Long
    Public RV As Single, GV As Single, BV As Single
    Public RN As Single, GN As Single, BN As Single
    Public PiT1 As Double, PiTAdd1 As Double
    Public PiT2 As Double, PiTAdd2 As Double
    Public ColVGr As Double, Colv_R As Double, Colv_G As Double, Colv_B As Double
    Public ColNGr As Double, ColN1 As Double, ColN2 As Double, ColN3 As Double
    Public cS(0 To 2) As Integer  ' R G B
    Public ColPR As Double, ColPG As Double, ColPB As Double, ColP As Long, ColB As Long
    Public ColPRsgn As Double, ColPGsgn As Double, ColPBsgn As Double
    Public FrqVbr  As Single, KR As Double
    Public Primes(0 To 148932) As Long
    Public PrK(1 To 3, 0 To 148932) As Integer
    Public PK(1 To 5, 0 To 148932) As Long
    Public lastSTP As Integer, maxViewP As Integer
    Public LQT As Double, LQT2 As Double, LQ_ModTime As Long, LQ_Pr_Mod57 As Long, St_Time As Long
    Public MaxC(0 To 2) As Single
    Public MinC(0 To 2) As Single
    Public xMin As Long, xMax As Long
    Public clCpu1 As Single, clCpu2 As Single, clCpu3 As Single
    Public K1 As Integer, K2 As Integer, M1 As Integer, M2 As Integer, z1 As Integer, Z2 As Integer
    Public Ef As Integer, FlgBlur As Integer, Z As Long
    Public xG As Integer, yG As Integer, cG As Long
    Public tx As Single, ty As Single, tZ As Single, stFirst  As Single
    Public Vsxx As Single, Vsyy As Single, Vszz As Single, MVolu  As Single
    Public Nx As Single, sR As Single, sb As Single, sg As Single, txP As Single, tyP As Single, vsZM As Single
    Public ColSt(0 To 256)  As Long
    Public ColTn(1 To 100) As Long, ColTv(0 To 1000) As Long, ColTp(1 To 255)
    Public xCol As Integer, yCol As Integer, xColStp As Integer
    Public vsX As Single, vsY As Single, vsTx As Single, vsTy As Single, TiS As Integer
    
    Public picBuff As PictureBox, picBuffSe As PictureBox, picTmp As PictureBox
    Public picView As PictureBox, picBuffSe2 As PictureBox
    Public Fst As Boolean, iH As Integer, reAl As Boolean
    Public minY As Integer, maxY As Integer, minY2 As Integer, maxY2 As Integer, minLY As Integer, maxLY As Integer
    Public minYC As Integer, maxYC As Integer, Ang As Integer
    Public Pt(0 To 511, 1 To 100) As POINTAPI
    Public PtL(0 To 511, 1 To 100) As POINTAPI
    Public Red As Byte, Green As Byte, Blue As Byte
    Public i As Double, A As Single, X As Integer, x2 As Integer, d As Single, zV As Single, Cu As Integer
    Public cnt As Single, ct As Integer, bT As Integer, cR As Long, Xtmp As Single, cnTmp As Single, Cof_X As Single
    Public direction As Single, R As Double, z11 As Single, z22 As Single, z33 As Single, z44 As Single, zvF As Single, zvT As Single
    Public bf1 As Single, bf2 As Single, bf3 As Single, bf4 As Single
    Public ttz As Integer, PointerX As Integer, PointerY As Integer
    Public aC As Single, iC As Single, ixC As Single, yC As Single, sLC As Single, Src As Single
    Public bC As Single, b2C As Single, xC As Single
    Public Blend As BLENDFUNCTION
    Public BlendPtr As Long
    Public PN(1 To 10000) As Long, PiN As Double
    Public Nvg(1 To 100) As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public PadBytes As Long
    Public BytesPerScanLine As Long
    
    Public m_hDIb As Long, m_hBmpOld As Long
    Public m_hDC As Long, DIBPtr As Long
    
    Public picWidth As Long, picHeight As Long
    
    Public Quality As Long
    Public SelectionQuality As Long
    Public aSelect As Boolean
    Public aSelectDone As Boolean
    Public SelectType As Long
    
    ' Rectangle selection coords
    Public XS1 As Single
    Public YS1 As Single
    Public XS2 As Single
    Public YS2 As Single
    
    Public NumLassoLines As Long
    
    Public fraX As Single
    Public fraY As Single
    
    Public STX As Long
    Public STY As Long

    Public Clk As cCpuClk
    Public C1  As Currency, c2  As Currency, cCPU As Currency, cCycles   As Currency
    Public cOver1 As Currency, cOver2 As Currency
    Public cQpc1 As Currency, cQpc2 As Currency
    Public cLong     As Currency
    Public cInteger  As Currency
    Public cNow      As Currency
    Public cStart    As Currency
    Public cStop     As Currency
    Public tFa    As Double
    
    Public Type POINTAPI
            X As Long
            Y As Long
    End Type
    Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    Public Type kPoint3D
            X As Long
            Y As Long
    End Type
    
    Public Resp As Long
    Public Type RGBQUAD
       rgbBlue As Byte
       rgbGreen As Byte
       rgbRed As Byte
       rgbReserved As Byte
    End Type
    Public Type Bitmap
      bmType As Long
      bmWidth As Long
      bmHeight As Long
      bmWidthBytes As Long
      bmPlanes As Integer
      bmBitsPixel As Integer
      bmBits As Long
    End Type
    
    Public PicInfo As Bitmap         'bitmap info structure
    
    Private Type BITMAPINFOHEADER
       biSize As Long
       biWidth As Long
       biHeight As Long
       biPlanes As Integer
       biBitCount As Integer
       biCompression As Long
       biSizeImage As Long
       biXPelsPerMeter As Long
       biYPelsPerMeter As Long
       biClrUsed As Long
       biClrImportant As Long
    End Type
    
    Public Type BLENDFUNCTION
      BlendOp As Byte
      BlendFlags As Byte
      SourceConstantAlpha As Byte
      AlphaFormat As Byte
    End Type
    
    Public Type prm
        t   As Double
        v   As Double
        ag  As Double
        mx  As Double
        E   As Double
        m   As Double
        c   As Double
        Cc  As Double
        al  As Double
        chk(1 To 10) As Byte
    End Type
    Public Para(0 To 20) As prm
    
    Private Const BIF_STATUSTEXT = &H4&
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const BIF_DONTGOBELOWDOMAIN = 2
    Private Const MAX_PATH = 260
    Private Const WM_USER = &H400
    Private Const BFFM_INITIALIZED = 1
    Private Const BFFM_SELCHANGED = 2
    Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
    Private Const BFFM_SETSELECTION = (WM_USER + 102)
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Type BrowseInfo
      hwndOwner      As Long
      pIDLRoot       As Long
      pszDisplayName As Long
      lpszTitle      As Long
      ulFlags        As Long
      lpfnCallback   As Long
      lParam         As Long
      iImage         As Long
    End Type
    Private m_CurrentDirectory As String
    Public mmmhwnd As Long
        
    Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetInputState Lib "user32" () As Long
    Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc2 As Long, ByVal lInt1 As Long, ByVal lInt2 As Long, ByVal lInt3 As Long, ByVal lInt4 As Long, ByVal hdc1 As Long, ByVal lInt5 As Long, ByVal lInt6 As Long, ByVal lInt7 As Long, ByVal lInt8 As Long, ByVal wFlags As Long) As Long
    Public Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    Public Declare Function TransParentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcnWidth As Long, ByVal srcnHeight As Long, ByVal CrTransparent As Long) As Long
    Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    Public Declare Function InvertRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
    Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Public Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
    Public Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
    Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
    Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
    Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    Public Declare Function PolyBezier Lib "gdi32.dll" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
    Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
    Public Declare Function PolyPolygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
    Public Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
    Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal xW As Long, ByVal yW As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal xW As Long, ByVal yW As Long, lpPoint As POINTAPI) As Long
    Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
    Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
    Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function Rectangel Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Public Declare Function GetTickCount Lib "Kernel32" () As Long
    Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Public Declare Function QueryPerformanceCounter Lib "Kernel32" (lpPerformanceCount As Currency) As Long
    Public Declare Function QueryPerformanceFrequency Lib "Kernel32" (lpFrequency As Currency) As Long
    Public Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Public Declare Sub GetSystemInfo Lib "Kernel32" (lpSystemInfo As SYSTEM_INFO)
    Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
    Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
    Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Public Declare Function UpdateColors Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
    Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
    Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
    Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    
    
Public Sub PrimeBase()     ' Generated first 148932 Prime Numbers in ( 1 to 2,000,000 ) range
Dim Lp1 As Long, Lp2 As Long, sR As Long, sT As Long, sR2 As Long, X As Long
Dim nPR() As Byte
ReDim nPR(1 To 2000001)
   
    sT = 3
    sR2 = 2000001
Rx:
    sR = Sqr((sR2))
 '''''''''''''''''''''''''
    For Lp1 = sT To sR Step 2
     If nPR(Lp1) = True Then nPR(Lp1) = False: GoTo Nx

        For Lp2 = Lp1 To sR2 Step Lp1
          nPR(Lp2) = True
        Next Lp2
Nx:
    Next Lp1
 '''''''''''''''''''''''''
    For Lp1 = 2 To sR
        nPR(Lp1) = False
    Next Lp1

    If sR > 3 Then sR2 = sR: GoTo Rx
 
 '''''''''''''''''''''''''
    Primes(1) = 2
    PK(1, 1) = 2
    PrK(2, 1) = 0: PK(2, 1) = 0                       'first Prime=2 and step for it =0
    PrK(3, 1) = 0: PK(3, 1) = 0
    sT = 2:         X = 2
    frmBase.lstLogs.AddItem 1 & " , " & Primes(1) & " , " & sT \ 2
    For Lp2 = 3 To 1999901 Step 2
        If nPR(Lp2) <> True Then
            Primes(X) = Lp2
            PK(1, X) = Lp2 \ 2
            PrK(2, X) = (Primes(X) - Primes(X - 1))
'            PK(2, x) = PrK(2, x): PK(4, x) = Log(PrK(2, x))
            If sT < PrK(2, X) Then sT = PrK(2, X): frmBase.lstLogs.AddItem X & " , " & Primes(X) & " , " & sT
            PrK(3, X) = sT
            PK(3, X) = sT: PK(5, X) = Log(sT + 1)
            X = X + 1
        End If
    Next Lp2
    
    
ReDim nPR(1 To 1)
End Sub


Public Sub LoadREG()
Exit Sub
Dim X As Integer
    Dim i As Long
    Dim intf As Integer
    Dim S As String
On Error Resume Next
With frmBase
    .chkPant = GetSetting("KV_M_B", "kvvisulation", "chkPant", "0")
    .txtQua = GetSetting("KV_M_B", "kvvisulation", "txtQua", "35")
    .chkAutoMax = GetSetting("KV_M_B", "kvvisulation", "chkAutoMax", 0)
    .chkInc = GetSetting("KV_M_B", "kvvisulation", "chkInc", 1)
    .chkAdjFreq(0) = GetSetting("KV_M_B", "kvvisulation", "fix", 1)
    .chkAdjFreq(1) = GetSetting("KV_M_B", "kvvisulation", "fix2", 1)
    .txtspm(9) = GetSetting("KV_M_B", "kvvisulation", "txtcrlfr", "35")
    .chkAdjFreq(1) = GetSetting("KV_M_B", "kvvisulation", "fixl2", 1)
    .chkClrAlter = GetSetting("KV_M_B", "kvvisulation", "clr2", 0)
    .fraControls.Height = GetSetting("KV_M_B", "kvvisulation", "fraControls", 350)
    .fraBlur.Height = GetSetting("KV_M_B", "kvvisulation", "fraBlur", 350)
    .fraProcess.Height = GetSetting("KV_M_B", "kvvisulation", "fraProcess", 350)
    .fraColors.Height = GetSetting("KV_M_B", "kvvisulation", "fraColors", 350)
    .fraLogs.Height = GetSetting("KV_M_B", "kvvisulation", "fraLogs", 350)
    .txtMaxC(0) = GetSetting("KV_M_B", "kvvisulation", "maxcol0", "255")
    .txtMinC(0) = GetSetting("KV_M_B", "kvvisulation", "mincol0", "0")
    .txtMaxC(1) = GetSetting("KV_M_B", "kvvisulation", "maxcol1", "255")
    .txtMinC(1) = GetSetting("KV_M_B", "kvvisulation", "mincol1", "0")
    .txtMaxC(2) = GetSetting("KV_M_B", "kvvisulation", "maxcol2", "255")
    .txtMinC(2) = GetSetting("KV_M_B", "kvvisulation", "mincol2", "0")
    .chkInverse = GetSetting("KV_M_B", "kvvisulation", "inverse", 0)
    .chkTransparent = GetSetting("KV_M_B", "kvvisulation", "chkTransparent", 1)
    .Combo1.ListIndex = GetSetting("KV_M_B", "kvvisulation", "combo1", 12)
    .chkSortP = GetSetting("KV_M_B", "kvvisulation", "chkSortP", 1)
    .chkAlpha = GetSetting("KV_M_B", "kvvisulation", "chkAlpha", 0)
    .chkAlphaEnable = GetSetting("KV_M_B", "kvvisulation", "chkAlphaEnable", 0)
    .chkFallCol = GetSetting("KV_M_B", "kvvisulation", "chkAlphaEnable", 0)
    .chkAgr(0) = GetSetting("KV_M_B", "kvvisulation", "chkAgr0", 0)
    .chkAgr(1) = GetSetting("KV_M_B", "kvvisulation", "chkAgr1", 0)
    
    SaveCount = GetSetting("KV_M_B", "kvvisulation", "SaveCount", 0)
    sPath = GetSetting("KV_M_B", "kvvisulation", "sPath", "")
    bPath = sPath
     
    For X = 0 To 36
       .txtspm(X) = GetSetting("KV_M_B", "kvvisulation", "SpM" & CStr(X))
    Next X
    For X = 0 To .ChkDraw.count + 1
       .ChkDraw(X) = GetSetting("KV_M_B", "kvvisulation", "ChkDraw" & CStr(X))
    Next X
    For X = 0 To .chkBlur.count + 1
       .chkBlur(X) = GetSetting("KV_M_B", "kvvisulation", "chkBlur" & CStr(X))
    Next X
    For X = 0 To .chkP3Opt.count + 1
       .chkP3Opt(X) = GetSetting("KV_M_B", "kvvisulation", "chkP3Opt" & CStr(X))
    Next X
    For X = 0 To .chkP4Opt.count + 1
       .chkP4Opt(X) = GetSetting("KV_M_B", "kvvisulation", "chkP4Opt" & CStr(X))
    Next X
    For X = 0 To .chkCol.count + 1
       .chkCol(X) = GetSetting("KV_M_B", "kvvisulation", "chkCol" & CStr(X))
    Next X
    For X = 0 To .chkCol.count + 1
       .txtRST(X) = GetSetting("KV_M_B", "kvvisulation", "txtRST" & CStr(X))
    Next X
    For X = 0 To .chkTimeEnable.count + 1
       .chkTimeEnable(X).Value = GetSetting("KV_M_B", "kvvisulation", "chkTimeEnable" & CStr(X), 1)
    Next X

    
'    On Error Resume Next
'      Debug.Print 1 / 0
'      If Err Then
'          MsgBox " . If Compile The Code Before Run . Its Runing About 2 Times Farster!!!", , " LQ_SKYS Present  ..."
'      End If

End With

End Sub
Public Sub SaveREG()
Exit Sub
  
  Dim X As Integer
On Error Resume Next
   
With frmBase
    SaveSetting "KV_M_B", "kvvisulation", "chkPant", .chkPant
    SaveSetting "KV_M_B", "kvvisulation", "txtQua", .txtQua
    SaveSetting "KV_M_B", "kvvisulation", "chkAutoMax", .chkAutoMax
    SaveSetting "KV_M_B", "kvvisulation", "chkInc", .chkInc
    SaveSetting "KV_M_B", "kvvisulation", "inverse", .chkInverse
    SaveSetting "KV_M_B", "kvvisulation", "fix", .chkAdjFreq(0)
    SaveSetting "KV_M_B", "kvvisulation", "fix2", .chkAdjFreq(1)
    SaveSetting "KV_M_B", "kvvisulation", "txtcrlfr", .txtspm(9)
    SaveSetting "KV_M_B", "kvvisulation", "fixl2", .chkAdjFreq(1)
    SaveSetting "KV_M_B", "kvvisulation", "clr2", .chkClrAlter
    SaveSetting "KV_M_B", "kvvisulation", "fraControls", .fraControls.Height
    SaveSetting "KV_M_B", "kvvisulation", "fraBlur", .fraBlur.Height
    SaveSetting "KV_M_B", "kvvisulation", "fraProcess", .fraProcess.Height
    SaveSetting "KV_M_B", "kvvisulation", "fraColors", .fraColors.Height
    SaveSetting "KV_M_B", "kvvisulation", "fraLogs", .fraLogs.Height
    SaveSetting "KV_M_B", "kvvisulation", "maxcol0", .txtMaxC(0)
    SaveSetting "KV_M_B", "kvvisulation", "mincol0", .txtMinC(0)
    SaveSetting "KV_M_B", "kvvisulation", "maxcol1", .txtMaxC(1)
    SaveSetting "KV_M_B", "kvvisulation", "mincol1", .txtMinC(1)
    SaveSetting "KV_M_B", "kvvisulation", "maxcol2", .txtMaxC(2)
    SaveSetting "KV_M_B", "kvvisulation", "mincol2", .txtMinC(2)
    SaveSetting "KV_M_B", "kvvisulation", "SaveCount", SaveCount
    SaveSetting "KV_M_B", "kvvisulation", "combo1", .Combo1.ListIndex
    SaveSetting "KV_M_B", "kvvisulation", "chkTransparent", .chkTransparent
    SaveSetting "KV_M_B", "kvvisulation", "chkSortP", .chkSortP
    SaveSetting "KV_M_B", "kvvisulation", "chkAlpha", .chkAlpha
    SaveSetting "KV_M_B", "kvvisulation", "chkAlphaEnable", .chkAlphaEnable
    SaveSetting "KV_M_B", "kvvisulation", "chkFallCol", .chkFallCol
    SaveSetting "KV_M_B", "kvvisulation", "chkAgr0", .chkAgr(0)
    SaveSetting "KV_M_B", "kvvisulation", "chkAgr1", .chkAgr(1)
    
    For X = 0 To 36 ' .txtspm.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "spm" & CStr(X), .txtspm(X)
    Next X
    For X = 0 To .ChkDraw.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "ChkDraw" & CStr(X), .ChkDraw(X)
    Next X
    For X = 0 To .chkBlur.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "chkBlur" & CStr(X), .chkBlur(X)
    Next X
    For X = 0 To .chkP3Opt.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "chkP3Opt" & CStr(X), .chkP3Opt(X)
    Next X
    For X = 0 To .chkP4Opt.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "chkP4Opt" & CStr(X), .chkP4Opt(X)
    Next X
    For X = 0 To .chkCol.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "chkCol" & CStr(X), .chkCol(X)
    Next X
    For X = 0 To .txtRST.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "txtRST" & CStr(X), .txtRST(X)
    Next X
    For X = 0 To .chkTimeEnable.count + 1
       SaveSetting "KV_M_B", "kvvisulation", "chkTimeEnable" & CStr(X), .chkTimeEnable(X).Value
    Next X
    
        
End With
   
End Sub


Public Sub SETBMI()
Dim SBI As BITMAPINFOHEADER
   
   With SBI
      .biSize = 40
      .biWidth = picView.Width \ Screen.TwipsPerPixelX ' picWidth
      .biHeight = picView.Height \ Screen.TwipsPerPixelY   ' picHeight
      .biPlanes = 1
      .biBitCount = 32 '24
      .biCompression = 0
   
      BytesPerScanLine = (((.biWidth * .biBitCount) + 31) \ 32) * 4
      PadBytes = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)
      
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   
   m_hDC = CreateCompatibleDC(0)
   m_hDIb = CreateDIBSection(m_hDC, SBI, 0, DIBPtr, 0, 0)
   m_hBmpOld = SelectObject(m_hDC, m_hDIb)
End Sub

Public Sub SaveJpeg(FSpec$, ByVal TheQuality As Long, APIC As PictureBox)
   Dim pvGDI As GDIPlusJPGConvertor
   
   SETBMI
   
   BitBlt m_hDC, 0, 0, picView.Width / Screen.TwipsPerPixelX, picView.Height / Screen.TwipsPerPixelY, APIC.hdc, 0, 0, vbSrcCopy
  
   Set pvGDI = New GDIPlusJPGConvertor
   
   pvGDI.SaveDIB picView.Width / Screen.TwipsPerPixelX, picView.Height / Screen.TwipsPerPixelY, DIBPtr, FSpec$, TheQuality
 
   Set pvGDI = Nothing
    
   SelectObject m_hDC, m_hBmpOld
   DeleteObject m_hDIb
   DeleteDC m_hDC
End Sub

Public Sub SaveJpegTele(FSpec$, ByVal TheQuality As Long, APIC As PictureBox)
   Dim pvGDI2 As GDIPlusJPGConvertor
   
   SETBMI2
   
   BitBlt m_hDC, 0, 0, 256, 256, APIC.hdc, 0, 0, vbSrcCopy
  
   Set pvGDI2 = New GDIPlusJPGConvertor
   
   pvGDI2.SaveDIB 256, 256, DIBPtr, FSpec$, TheQuality
 
   Set pvGDI2 = Nothing
    
   SelectObject m_hDC, m_hBmpOld
   DeleteObject m_hDIb
   DeleteDC m_hDC
End Sub
Public Sub SETBMI2()
Dim sib2 As BITMAPINFOHEADER
   
   With sib2
      .biSize = 40
      .biWidth = picView.Width \ Screen.TwipsPerPixelX ' picWidth
      .biHeight = picView.Height \ Screen.TwipsPerPixelY   ' picHeight
      .biPlanes = 1
      .biBitCount = 32 '24
      .biCompression = 0
   
      BytesPerScanLine = (((.biWidth * .biBitCount) + 31) \ 32) * 4
      PadBytes = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)
      
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   
   m_hDC = CreateCompatibleDC(0)
   m_hDIb = CreateDIBSection(m_hDC, sib2, 0, DIBPtr, 0, 0)
   m_hBmpOld = SelectObject(m_hDC, m_hDIb)
End Sub

Public Function BrowseForFolder(StartDir As String, owner As Long, title As String) As String
  Dim lpIDList As Long
  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
  m_CurrentDirectory = StartDir & vbNullChar
  szTitle = title
  With tBrowseInfo
    .hwndOwner = owner
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
    .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  Else
    BrowseForFolder = ""
  End If
End Function
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
Dim lpIDList As Long
Dim ret As Long
Dim sBuffer As String
On Error Resume Next
mmmhwnd = hWnd
Select Case uMsg
  Case BFFM_INITIALIZED
    Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
  Case BFFM_SELCHANGED
    sBuffer = Space(MAX_PATH)
    ret = SHGetPathFromIDList(lp, sBuffer)
    If ret = 1 Then
      Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
    End If
End Select
BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function

